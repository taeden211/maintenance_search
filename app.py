from __future__ import annotations

import json
import re
import sqlite3
import threading
import traceback
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path
from queue import Empty, Queue
from typing import Optional

import joblib
import numpy as np
from openpyxl import load_workbook
from sklearn.feature_extraction.text import TfidfVectorizer

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


TOKEN_RE = re.compile(r"[0-9A-Za-z가-힣]+")

YEAR_RE = re.compile(r"(20\d{2})")
FILE_RE = re.compile(r"(20\d{2})(0[1-9]|1[0-2])")

APP_TITLE = "유지보수 사례 검색기"
CACHE_DIR_NAME = ".maintenance_search_cache"
CACHE_DB_NAME = "maintenance_cases.sqlite"
CACHE_ARTIFACT_NAME = "search_artifacts.joblib"
CACHE_REPORT_NAME = "index_report.json"
DEFAULT_TOP_N = 20
DEFAULT_BM25_STAGE = 120
SEARCH_INDEX_VERSION = 3
QUERY_HINT = "검색할 장애 증상이나 조치 내용을 입력하세요"
EXCLUDE_FILE_KEYWORDS = (
    "검색결과",
    "보고서",
    "분석",
    "result",
    "output",
    "backup",
    "bak",
)
CHECKED_VALUES = {"O", "Y", "YES", "TRUE", "1", "○", "예"}


@dataclass(slots=True)
class MaintenanceCase:
    case_id: int
    source_file: str
    source_sheet: str
    row_num: int
    year: int
    month: int
    date_text: str
    department: str
    user: str
    issue_text: str
    action_text: str
    apc: str
    pc_filter: str
    utmp: str
    sheet_title: str
    search_text: str


@dataclass(slots=True)
class SearchFilters:
    year_from: Optional[int] = None
    year_to: Optional[int] = None
    department: str = ""
    user: str = ""
    require_apc: bool = False
    require_pc_filter: bool = False
    require_utmp: bool = False


class BM25Index:
    def __init__(self, tokenized_docs: list[list[str]], k1: float = 1.5, b: float = 0.75) -> None:
        self.k1 = k1
        self.b = b
        self.tokenized_docs = tokenized_docs
        self.term_freqs = [Counter(doc) for doc in tokenized_docs]
        self.doc_len = np.array([len(doc) for doc in tokenized_docs], dtype=np.float32)
        self.avgdl = float(self.doc_len.mean()) if len(self.doc_len) else 0.0
        self.doc_freq: dict[str, int] = {}
        for doc in tokenized_docs:
            for term in set(doc):
                self.doc_freq[term] = self.doc_freq.get(term, 0) + 1
        self.idf: dict[str, float] = {}
        n_docs = len(tokenized_docs)
        for term, freq in self.doc_freq.items():
            self.idf[term] = float(np.log(1.0 + (n_docs - freq + 0.5) / (freq + 0.5)))

    def score(self, query_tokens: list[str]) -> np.ndarray:
        if not self.tokenized_docs or not query_tokens:
            return np.zeros(len(self.tokenized_docs), dtype=np.float32)

        scores = np.zeros(len(self.tokenized_docs), dtype=np.float32)
        for term in query_tokens:
            if term not in self.idf:
                continue
            idf = self.idf[term]
            for idx, term_freq in enumerate(self.term_freqs):
                tf = term_freq.get(term, 0)
                if tf == 0:
                    continue
                denom = tf + self.k1 * (1.0 - self.b + self.b * (self.doc_len[idx] / self.avgdl if self.avgdl else 0.0))
                scores[idx] += idf * (tf * (self.k1 + 1.0)) / denom
        return scores


@dataclass(slots=True)
class SearchArtifacts:
    records: list[MaintenanceCase]
    vectorizer: TfidfVectorizer
    tfidf_matrix: object
    bm25: BM25Index
    searchable_texts: list[str]
    tokenized_texts: list[list[str]]
    years: list[int]


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).replace("\u3000", " ").strip()


def tokenize(text: str) -> list[str]:
    return [token for token in TOKEN_RE.findall(text.lower()) if token.strip()]


def parse_year_month_from_path(path: Path) -> tuple[int, int]:
    for candidate in [path.stem, path.name, path.parent.name]:
        match = FILE_RE.search(candidate)
        if match:
            return int(match.group(1)), int(match.group(2))
        match = YEAR_RE.search(candidate)
        if match:
            return int(match.group(1)), 1
    return 0, 0


def safe_cell(value: object) -> str:
    return normalize_text(value)


def is_checked(value: str) -> bool:
    return normalize_text(value).strip().upper() in CHECKED_VALUES


def get_excel_skip_reason(path: Path) -> str:
    if path.name.startswith("~$"):
        return "엑셀 임시 파일"
    lowered = path.name.lower()
    for keyword in EXCLUDE_FILE_KEYWORDS:
        if keyword.lower() in lowered:
            return f"제외 키워드 포함: {keyword}"
    return ""


def looks_like_maintenance_sheet(sheet: object) -> bool:
    texts: list[str] = []
    for row in sheet.iter_rows(min_row=1, max_row=8, max_col=10, values_only=True):
        for value in row:
            text = safe_cell(value).replace(" ", "")
            if text:
                texts.append(text)

    has_issue = any("장애내용" in text or text == "장애" for text in texts)
    has_action = any("조치내용" in text or "조치" in text for text in texts)
    has_department = any("부서" in text for text in texts)
    has_date = any("날짜" in text for text in texts)
    return has_issue and has_action and (has_department or has_date)


def create_empty_index_report() -> dict[str, object]:
    return {
        "total_files": 0,
        "read_files": 0,
        "excluded_files": 0,
        "skipped_files": 0,
        "total_rows": 0,
        "valid_cases": 0,
        "excluded_rows": 0,
        "multi_sheet_files": 0,
        "missing_departments": 0,
        "missing_users": 0,
        "missing_apc": 0,
        "missing_pc_filter": 0,
        "missing_utmp": 0,
        "excluded_file_details": [],
        "skipped_file_details": [],
        "multi_sheet_details": [],
        "sheet_details": [],
    }


class MaintenanceRepository:
    def __init__(self, folder: Path) -> None:
        self.folder = folder
        self.cache_dir = folder / CACHE_DIR_NAME
        self.db_path = self.cache_dir / CACHE_DB_NAME
        self.artifact_path = self.cache_dir / CACHE_ARTIFACT_NAME
        self.report_path = self.cache_dir / CACHE_REPORT_NAME

    def ensure_cache_dir(self) -> None:
        self.cache_dir.mkdir(parents=True, exist_ok=True)

    def clear(self) -> None:
        self.ensure_cache_dir()
        if self.db_path.exists():
            self.db_path.unlink()
        if self.artifact_path.exists():
            self.artifact_path.unlink()
        if self.report_path.exists():
            self.report_path.unlink()

    def save_records(self, records: list[MaintenanceCase]) -> None:
        self.ensure_cache_dir()
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS cases (
                    case_id INTEGER PRIMARY KEY,
                    source_file TEXT,
                    source_sheet TEXT,
                    row_num INTEGER,
                    year INTEGER,
                    month INTEGER,
                    date_text TEXT,
                    department TEXT,
                    user TEXT,
                    issue_text TEXT,
                    action_text TEXT,
                    apc TEXT,
                    pc_filter TEXT,
                    utmp TEXT,
                    sheet_title TEXT,
                    search_text TEXT
                )
                """
            )
            conn.execute("DELETE FROM cases")
            conn.executemany(
                """
                INSERT INTO cases (
                    case_id, source_file, source_sheet, row_num, year, month, date_text,
                    department, user, issue_text, action_text, apc, pc_filter, utmp,
                    sheet_title, search_text
                ) VALUES (
                    :case_id, :source_file, :source_sheet, :row_num, :year, :month, :date_text,
                    :department, :user, :issue_text, :action_text, :apc, :pc_filter, :utmp,
                    :sheet_title, :search_text
                )
                """,
                [asdict(record) for record in records],
            )
            conn.commit()

    def load_records(self) -> list[MaintenanceCase]:
        if not self.db_path.exists():
            return []
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute("SELECT * FROM cases ORDER BY case_id").fetchall()
        return [MaintenanceCase(**dict(row)) for row in rows]

    def save_artifacts(self, artifacts: SearchArtifacts) -> None:
        self.ensure_cache_dir()
        joblib.dump(
            {
                "index_version": SEARCH_INDEX_VERSION,
                "records": artifacts.records,
                "vectorizer": artifacts.vectorizer,
                "tfidf_matrix": artifacts.tfidf_matrix,
                "bm25": artifacts.bm25,
                "searchable_texts": artifacts.searchable_texts,
                "tokenized_texts": artifacts.tokenized_texts,
                "years": artifacts.years,
            },
            self.artifact_path,
        )

    def load_artifacts(self) -> Optional[SearchArtifacts]:
        if not self.artifact_path.exists():
            return None
        payload = joblib.load(self.artifact_path)
        if payload.get("index_version") != SEARCH_INDEX_VERSION:
            return None
        return SearchArtifacts(
            records=payload["records"],
            vectorizer=payload["vectorizer"],
            tfidf_matrix=payload["tfidf_matrix"],
            bm25=payload["bm25"],
            searchable_texts=payload["searchable_texts"],
            tokenized_texts=payload["tokenized_texts"],
            years=payload["years"],
        )

    def save_report(self, report: dict[str, object]) -> None:
        self.ensure_cache_dir()
        self.report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    def load_report(self) -> Optional[dict[str, object]]:
        if not self.report_path.exists():
            return None
        return json.loads(self.report_path.read_text(encoding="utf-8"))


class MaintenanceSearchEngine:
    def __init__(self) -> None:
        self.repository: Optional[MaintenanceRepository] = None
        self.artifacts: Optional[SearchArtifacts] = None
        self.records: list[MaintenanceCase] = []
        self.loaded_folder: Optional[Path] = None
        self.index_report: Optional[dict[str, object]] = None

    @property
    def is_ready(self) -> bool:
        return self.artifacts is not None and bool(self.records)

    def load_folder(self, folder: Path) -> bool:
        self.repository = MaintenanceRepository(folder)
        artifacts = self.repository.load_artifacts()
        if artifacts is None:
            return False
        self.artifacts = artifacts
        self.records = artifacts.records
        self.loaded_folder = folder
        self.index_report = self.repository.load_report() or self._legacy_report(len(self.records))
        return True

    def build_from_folder(self, folder: Path) -> tuple[int, list[int], dict[str, object]]:
        repository = MaintenanceRepository(folder)
        repository.clear()
        records, report = self._collect_cases(folder)
        if not records:
            raise RuntimeError("엑셀 파일에서 유효한 장애 사례를 찾지 못했습니다.")

        searchable_texts = [record.search_text for record in records]
        tokenized_texts = [tokenize(text) for text in searchable_texts]

        vectorizer = TfidfVectorizer(
            analyzer="char_wb",
            ngram_range=(2, 5),
            min_df=1,
            sublinear_tf=True,
            norm="l2",
        )
        tfidf_matrix = vectorizer.fit_transform(searchable_texts)
        bm25 = BM25Index(tokenized_texts)
        years = sorted({record.year for record in records if record.year})

        artifacts = SearchArtifacts(
            records=records,
            vectorizer=vectorizer,
            tfidf_matrix=tfidf_matrix,
            bm25=bm25,
            searchable_texts=searchable_texts,
            tokenized_texts=tokenized_texts,
            years=years,
        )
        repository.save_records(records)
        repository.save_artifacts(artifacts)
        report["valid_cases"] = len(records)
        repository.save_report(report)
        self.repository = repository
        self.artifacts = artifacts
        self.records = records
        self.loaded_folder = folder
        self.index_report = report
        return len(records), years, report

    def _collect_cases(self, folder: Path) -> tuple[list[MaintenanceCase], dict[str, object]]:
        records: list[MaintenanceCase] = []
        report = create_empty_index_report()
        case_id = 1

        xlsx_paths = sorted(folder.rglob("*.xlsx"))
        report["total_files"] = len(xlsx_paths)
        for xlsx_path in xlsx_paths:
            skip_reason = get_excel_skip_reason(xlsx_path)
            if skip_reason:
                report["excluded_files"] = int(report["excluded_files"]) + 1
                report["excluded_file_details"].append({"file": xlsx_path.name, "reason": skip_reason})
                continue
            year, month = parse_year_month_from_path(xlsx_path)
            try:
                workbook = load_workbook(xlsx_path, data_only=True, read_only=True)
            except Exception as exc:
                report["skipped_files"] = int(report["skipped_files"]) + 1
                report["skipped_file_details"].append({"file": xlsx_path.name, "reason": f"엑셀 로딩 실패: {exc}"})
                continue

            try:
                sheet_names = workbook.sheetnames
                read_sheet_count = 0
                file_valid_cases = 0

                for sheet_name in sheet_names:
                    sheet = workbook[sheet_name]
                    if not looks_like_maintenance_sheet(sheet):
                        report["sheet_details"].append(
                            {
                                "file": xlsx_path.name,
                                "sheet": sheet_name,
                                "status": "제외",
                                "reason": "유지보수 양식으로 판단되지 않음",
                                "rows": 0,
                                "valid_cases": 0,
                                "excluded_rows": 0,
                            }
                        )
                        continue

                    read_sheet_count += 1
                    sheet_title = safe_cell(sheet.cell(1, 1).value)
                    source_sheet = safe_cell(sheet.title)
                    last_date = ""
                    sheet_rows = 0
                    sheet_valid_cases = 0
                    sheet_excluded_rows = 0

                    for row_num, row_values in enumerate(
                        sheet.iter_rows(min_row=5, max_col=9, values_only=True),
                        start=5,
                    ):
                        sheet_rows += 1
                        report["total_rows"] = int(report["total_rows"]) + 1
                        row_values = tuple(row_values) + (None,) * (9 - len(row_values))

                        seq = row_values[0]
                        date_value = safe_cell(row_values[1])
                        if date_value:
                            last_date = date_value
                        date_text = date_value or last_date
                        department = safe_cell(row_values[2])
                        user = safe_cell(row_values[3])
                        issue_text = safe_cell(row_values[4])
                        action_text = safe_cell(row_values[5])
                        apc = safe_cell(row_values[6])
                        pc_filter = safe_cell(row_values[7])
                        utmp = safe_cell(row_values[8])

                        if not issue_text and not action_text:
                            sheet_excluded_rows += 1
                            report["excluded_rows"] = int(report["excluded_rows"]) + 1
                            continue

                        if not department:
                            report["missing_departments"] = int(report["missing_departments"]) + 1
                        if not user:
                            report["missing_users"] = int(report["missing_users"]) + 1
                        if not apc:
                            report["missing_apc"] = int(report["missing_apc"]) + 1
                        if not pc_filter:
                            report["missing_pc_filter"] = int(report["missing_pc_filter"]) + 1
                        if not utmp:
                            report["missing_utmp"] = int(report["missing_utmp"]) + 1

                        searchable_text = " ".join(
                            part
                            for part in [
                                issue_text,
                                action_text,
                                department,
                                user,
                                date_text,
                                sheet_title,
                                xlsx_path.stem,
                            ]
                            if part
                        )

                        records.append(
                            MaintenanceCase(
                                case_id=case_id,
                                source_file=xlsx_path.name,
                                source_sheet=source_sheet,
                                row_num=row_num,
                                year=year,
                                month=month,
                                date_text=date_text,
                                department=department,
                                user=user,
                                issue_text=issue_text,
                                action_text=action_text,
                                apc=apc,
                                pc_filter=pc_filter,
                                utmp=utmp,
                                sheet_title=sheet_title,
                                search_text=searchable_text,
                            )
                        )
                        case_id += 1
                        sheet_valid_cases += 1
                        file_valid_cases += 1

                    report["sheet_details"].append(
                        {
                            "file": xlsx_path.name,
                            "sheet": source_sheet,
                            "status": "읽음",
                            "reason": "",
                            "rows": sheet_rows,
                            "valid_cases": sheet_valid_cases,
                            "excluded_rows": sheet_excluded_rows,
                        }
                    )

                if len(sheet_names) > 1:
                    report["multi_sheet_files"] = int(report["multi_sheet_files"]) + 1
                    report["multi_sheet_details"].append(
                        {
                            "file": xlsx_path.name,
                            "total_sheets": len(sheet_names),
                            "read_sheets": read_sheet_count,
                            "note": "유지보수 양식 시트만 읽음" if read_sheet_count < len(sheet_names) else "전체 시트 읽음",
                        }
                    )

                if read_sheet_count:
                    report["read_files"] = int(report["read_files"]) + 1
                else:
                    report["skipped_files"] = int(report["skipped_files"]) + 1
                    report["skipped_file_details"].append(
                        {"file": xlsx_path.name, "reason": "유지보수 양식 시트 없음"}
                    )
            finally:
                workbook.close()

        report["valid_cases"] = len(records)
        return records, report

    @staticmethod
    def _legacy_report(valid_cases: int) -> dict[str, object]:
        report = create_empty_index_report()
        report["valid_cases"] = valid_cases
        return report

    def _apply_filters(self, filters: SearchFilters) -> list[int]:
        indices: list[int] = []
        department_term = filters.department.strip().lower()
        user_term = filters.user.strip().lower()

        for idx, record in enumerate(self.records):
            if filters.year_from is not None:
                if not record.year or record.year < filters.year_from:
                    continue
            if filters.year_to is not None:
                if not record.year or record.year > filters.year_to:
                    continue
            if department_term and department_term not in record.department.lower():
                continue
            if user_term and user_term not in record.user.lower():
                continue
            if filters.require_apc and not is_checked(record.apc):
                continue
            if filters.require_pc_filter and not is_checked(record.pc_filter):
                continue
            if filters.require_utmp and not is_checked(record.utmp):
                continue
            indices.append(idx)
        return indices

    def search(
        self,
        query: str,
        filters: SearchFilters,
        top_n: int = DEFAULT_TOP_N,
        bm25_stage: int = DEFAULT_BM25_STAGE,
    ) -> list[dict[str, object]]:
        if not self.is_ready or self.artifacts is None:
            raise RuntimeError("검색 인덱스가 아직 준비되지 않았습니다.")

        query = query.strip()
        if not query:
            return []

        candidate_indices = self._apply_filters(filters)
        if not candidate_indices:
            return []

        query_tokens = tokenize(query)
        bm25_scores = self.artifacts.bm25.score(query_tokens)
        candidate_scores = bm25_scores[candidate_indices]

        query_vector = self.artifacts.vectorizer.transform([query])
        candidate_matrix = self.artifacts.tfidf_matrix[candidate_indices]
        vector_scores = np.asarray(candidate_matrix.dot(query_vector.T).toarray()).ravel()

        bm25_norm = self._normalize_scores(candidate_scores)
        vec_norm = self._normalize_scores(vector_scores)
        final_scores = (0.45 * bm25_norm) + (0.55 * vec_norm)
        positive_indices = np.where(final_scores > 1e-8)[0]
        if positive_indices.size == 0:
            return []

        ranked = positive_indices[np.argsort(-final_scores[positive_indices])][:top_n]
        results: list[dict[str, object]] = []
        for rank, rel_idx in enumerate(ranked, start=1):
            case = self.records[candidate_indices[rel_idx]]
            results.append(
                {
                    "rank": rank,
                    "score": float(final_scores[rel_idx]),
                    "bm25": float(candidate_scores[rel_idx]),
                    "vector": float(vector_scores[rel_idx]),
                    "keyword_score": float(bm25_norm[rel_idx]),
                    "similarity_score": float(vec_norm[rel_idx]),
                    "case": case,
                }
            )
        return results

    @staticmethod
    def _normalize_scores(values: np.ndarray) -> np.ndarray:
        if values.size == 0:
            return values.astype(np.float32)
        max_value = float(values.max())
        min_value = float(values.min())
        if np.isclose(max_value, min_value):
            if max_value > 0:
                return np.ones_like(values, dtype=np.float32)
            return np.zeros_like(values, dtype=np.float32)
        return ((values - min_value) / (max_value - min_value)).astype(np.float32)


class MaintenanceSearchApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1400x860")
        self.root.minsize(1200, 760)

        self.engine = MaintenanceSearchEngine()
        self.queue: Queue[tuple[str, object]] = Queue()
        self.search_results: list[dict[str, object]] = []
        self.year_values: list[int] = []
        self._build_busy = False

        self.folder_var = tk.StringVar(value=str(Path.cwd() / "유지보수내역서 25.01~26.04"))
        self.query_var = tk.StringVar()
        self.top_n_var = tk.IntVar(value=20)
        self.year_from_var = tk.StringVar(value="전체")
        self.year_to_var = tk.StringVar(value="전체")
        self.department_var = tk.StringVar()
        self.user_var = tk.StringVar()
        self.apc_var = tk.BooleanVar(value=False)
        self.pc_filter_var = tk.BooleanVar(value=False)
        self.utmp_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="인덱스를 불러오거나 새로 구축하세요.")

        self._build_ui()
        self._show_query_hint()
        self._try_load_existing_index()

    def _build_ui(self) -> None:
        style = ttk.Style()
        style.configure("Search.TEntry", foreground="#111111")
        style.configure("SearchHint.TEntry", foreground="#777777")

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        top = ttk.Frame(self.root, padding=10)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="데이터 폴더").grid(row=0, column=0, sticky="w")
        folder_entry = ttk.Entry(top, textvariable=self.folder_var)
        folder_entry.grid(row=0, column=1, sticky="ew", padx=(8, 8))
        ttk.Button(top, text="찾기", command=self._browse_folder).grid(row=0, column=2, padx=(0, 6))
        ttk.Button(top, text="인덱스 구축", command=self._prepare_index).grid(row=0, column=3, padx=(0, 6))
        ttk.Button(top, text="인덱스 리포트 보기", command=self._show_index_report).grid(row=0, column=4)

        search = ttk.LabelFrame(self.root, text="검색", padding=10)
        search.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        search.columnconfigure(1, weight=1)

        ttk.Label(search, text="장애내용 / 조치내용").grid(row=0, column=0, sticky="w")
        self.query_entry = ttk.Entry(search, textvariable=self.query_var, style="Search.TEntry")
        self.query_entry.grid(row=0, column=1, sticky="ew", padx=(8, 8))
        self.query_entry.bind("<FocusIn>", self._hide_query_hint)
        self.query_entry.bind("<FocusOut>", self._show_query_hint)
        self.query_entry.bind("<Return>", lambda _: self._run_search())
        ttk.Label(search, text="결과 수").grid(row=0, column=2, sticky="e")
        ttk.Spinbox(search, from_=5, to=100, textvariable=self.top_n_var, width=6).grid(row=0, column=3, sticky="w", padx=(8, 0))
        ttk.Button(search, text="검색", command=self._run_search).grid(row=0, column=4, padx=(10, 0))

        filter_frame = ttk.Frame(search)
        filter_frame.grid(row=1, column=0, columnspan=5, sticky="ew", pady=(10, 8))
        for col in range(10):
            filter_frame.columnconfigure(col, weight=0)

        ttk.Label(filter_frame, text="연도 from").grid(row=0, column=0, sticky="w")
        self.year_from_combo = ttk.Combobox(filter_frame, textvariable=self.year_from_var, values=["전체"], width=10, state="readonly")
        self.year_from_combo.grid(row=0, column=1, sticky="w", padx=(6, 12))

        ttk.Label(filter_frame, text="연도 to").grid(row=0, column=2, sticky="w")
        self.year_to_combo = ttk.Combobox(filter_frame, textvariable=self.year_to_var, values=["전체"], width=10, state="readonly")
        self.year_to_combo.grid(row=0, column=3, sticky="w", padx=(6, 12))

        ttk.Label(filter_frame, text="부서 포함").grid(row=0, column=4, sticky="w")
        ttk.Entry(filter_frame, textvariable=self.department_var, width=16).grid(
            row=0, column=5, sticky="w", padx=(6, 16)
        )

        ttk.Label(filter_frame, text="사용자 포함").grid(row=0, column=6, sticky="w")
        ttk.Entry(filter_frame, textvariable=self.user_var, width=14).grid(
            row=0, column=7, sticky="w", padx=(6, 16)
        )

        flags = ttk.Frame(filter_frame)
        flags.grid(row=0, column=8, columnspan=2, sticky="w")
        ttk.Checkbutton(flags, text="APC", variable=self.apc_var).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Checkbutton(flags, text="PC filter", variable=self.pc_filter_var).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Checkbutton(flags, text="UTMP", variable=self.utmp_var).pack(side=tk.LEFT)

        body = ttk.Panedwindow(self.root, orient=tk.VERTICAL)
        body.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.root.rowconfigure(2, weight=1)

        results_frame = ttk.Labelframe(body, text="유사 사례 목록", padding=6)
        detail_frame = ttk.Labelframe(body, text="선택 항목 상세", padding=6)
        body.add(results_frame, weight=3)
        body.add(detail_frame, weight=2)

        results_frame.rowconfigure(0, weight=1)
        results_frame.columnconfigure(0, weight=1)
        columns = ("rank", "score", "year", "date", "dept", "user", "issue", "action", "file")
        self.tree = ttk.Treeview(results_frame, columns=columns, show="headings", height=16)
        headings = {
            "rank": "순위",
            "score": "관련도",
            "year": "연도",
            "date": "날짜",
            "dept": "부서",
            "user": "사용자",
            "issue": "장애내용",
            "action": "조치내용",
            "file": "원본파일",
        }
        widths = {
            "rank": 60,
            "score": 90,
            "year": 70,
            "date": 100,
            "dept": 180,
            "user": 120,
            "issue": 260,
            "action": 320,
            "file": 180,
        }
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], anchor="w")
        vsb = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<<TreeviewSelect>>", self._show_selected_detail)

        ttk.Button(results_frame, text="검색 결과 엑셀 저장", command=self._export_results).grid(
            row=2, column=0, sticky="e", pady=(8, 0)
        )

        detail_frame.rowconfigure(0, weight=1)
        detail_frame.columnconfigure(0, weight=1)
        self.detail_text = tk.Text(detail_frame, wrap="word", height=12)
        self.detail_text.grid(row=0, column=0, sticky="nsew")
        detail_scroll = ttk.Scrollbar(detail_frame, orient=tk.VERTICAL, command=self.detail_text.yview)
        self.detail_text.configure(yscrollcommand=detail_scroll.set)
        detail_scroll.grid(row=0, column=1, sticky="ns")

        bottom = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        bottom.grid(row=3, column=0, sticky="ew")
        bottom.columnconfigure(0, weight=1)
        self.progress = ttk.Progressbar(bottom, mode="indeterminate")
        self.progress.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        ttk.Label(bottom, textvariable=self.status_var).grid(row=1, column=0, sticky="w", pady=(6, 0))

    def _browse_folder(self) -> None:
        selected = filedialog.askdirectory(title="유지보수 엑셀 폴더 선택", initialdir=self.folder_var.get() or str(Path.cwd()))
        if selected:
            self.folder_var.set(selected)
            self.status_var.set(f"폴더 선택 완료: {selected}")
            self._load_index()

    def _show_query_hint(self, event: object | None = None) -> None:
        if not self.query_var.get().strip():
            self.query_var.set(QUERY_HINT)
            self.query_entry.configure(style="SearchHint.TEntry")

    def _hide_query_hint(self, event: object | None = None) -> None:
        if self.query_var.get() == QUERY_HINT:
            self.query_var.set("")
            self.query_entry.configure(style="Search.TEntry")

    def _prepare_index(self) -> None:
        folder = Path(self.folder_var.get()).expanduser()
        if not folder.exists():
            messagebox.showerror(APP_TITLE, "데이터 폴더를 찾을 수 없습니다.")
            return
        if self._build_busy:
            return
        try:
            if self.engine.load_folder(folder):
                self.year_values = self.engine.artifacts.years if self.engine.artifacts else []
                self._refresh_year_combos()
                self.status_var.set(f"저장된 인덱스 로드 완료: {len(self.engine.records)}건")
                messagebox.showinfo(APP_TITLE, f"저장된 인덱스를 불러왔습니다.\n총 {len(self.engine.records)}건")
                return
        except Exception:
            self.status_var.set("저장된 인덱스를 불러올 수 없어 새로 구축합니다.")
        self._start_build(folder)

    def _start_build(self, folder: Optional[Path] = None) -> None:
        folder = folder or Path(self.folder_var.get()).expanduser()
        if not folder.exists():
            messagebox.showerror(APP_TITLE, "데이터 폴더를 찾을 수 없습니다.")
            return
        self._set_busy(True)
        self._build_busy = True
        self.status_var.set("인덱스를 구축 중입니다. 잠시만 기다려 주세요.")
        threading.Thread(target=self._build_worker, args=(folder,), daemon=True).start()
        self.root.after(100, self._poll_queue)

    def _build_worker(self, folder: Path) -> None:
        try:
            count, years, report = self.engine.build_from_folder(folder)
            self.queue.put(("build_done", (count, years, folder, report)))
        except Exception as exc:
            self.queue.put(("error", (exc, traceback.format_exc())))

    def _load_index(self) -> None:
        folder = Path(self.folder_var.get()).expanduser()
        if not folder.exists():
            self.status_var.set("데이터 폴더가 없습니다.")
            return
        try:
            if self.engine.load_folder(folder):
                self.year_values = self.engine.artifacts.years if self.engine.artifacts else []
                self._refresh_year_combos()
                self.status_var.set(f"인덱스 로드 완료: {len(self.engine.records)}건")
            else:
                self.status_var.set("저장된 인덱스가 없습니다. 인덱스 구축을 실행하세요.")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"인덱스 로드 실패\n\n{exc}")

    def _try_load_existing_index(self) -> None:
        self._load_index()

    def _poll_queue(self) -> None:
        try:
            kind, payload = self.queue.get_nowait()
        except Empty:
            if self._build_busy:
                self.root.after(100, self._poll_queue)
            return

        if kind == "build_done":
            count, years, folder, report = payload
            self.year_values = years
            self._refresh_year_combos()
            self.status_var.set(f"인덱스 구축 완료: {count}건 / 폴더: {folder}")
            messagebox.showinfo(APP_TITLE, self._format_build_summary(report))
        elif kind == "error":
            exc, tb = payload
            messagebox.showerror(APP_TITLE, f"작업 중 오류가 발생했습니다.\n\n{exc}\n\n{tb}")
            self.status_var.set("오류가 발생했습니다.")
        self._build_busy = False
        self._set_busy(False)

    def _format_build_summary(self, report: dict[str, object]) -> str:
        return (
            "인덱스 구축 완료\n"
            f"읽은 파일: {report.get('read_files', 0)}개 / 전체 {report.get('total_files', 0)}개\n"
            f"유효 사례: {report.get('valid_cases', 0)}건\n"
            f"제외 행: {report.get('excluded_rows', 0)}건\n"
            f"제외 파일: {report.get('excluded_files', 0)}개\n"
            f"다중 시트 파일: {report.get('multi_sheet_files', 0)}개\n"
            "자세한 내용은 [인덱스 리포트 보기]에서 확인하세요."
        )

    def _show_index_report(self) -> None:
        report = self.engine.index_report
        if report is None:
            folder = Path(self.folder_var.get()).expanduser()
            if folder.exists():
                report = MaintenanceRepository(folder).load_report()
                self.engine.index_report = report
        if report is None:
            messagebox.showwarning(APP_TITLE, "확인할 인덱스 리포트가 없습니다.")
            return

        window = tk.Toplevel(self.root)
        window.title("인덱스 리포트")
        window.geometry("780x520")
        window.minsize(680, 440)

        notebook = ttk.Notebook(window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        summary_rows = [
            ("전체 파일", report.get("total_files", 0)),
            ("읽은 파일", report.get("read_files", 0)),
            ("제외 파일", report.get("excluded_files", 0)),
            ("건너뛴 파일", report.get("skipped_files", 0)),
            ("총 행 수", report.get("total_rows", 0)),
            ("유효 사례", report.get("valid_cases", 0)),
            ("제외 행", report.get("excluded_rows", 0)),
            ("다중 시트 파일", report.get("multi_sheet_files", 0)),
        ]
        summary_frame = ttk.Frame(notebook, padding=8)
        notebook.add(summary_frame, text="요약")
        self._populate_tree(summary_frame, ("item", "value"), {"item": "항목", "value": "값"}, summary_rows)

        excluded_rows = [
            (item.get("file", ""), "제외", item.get("reason", ""))
            for item in report.get("excluded_file_details", [])
        ]
        excluded_rows.extend(
            (item.get("file", ""), "건너뜀", item.get("reason", ""))
            for item in report.get("skipped_file_details", [])
        )
        excluded_frame = ttk.Frame(notebook, padding=8)
        notebook.add(excluded_frame, text="제외 파일")
        self._populate_tree(
            excluded_frame,
            ("file", "status", "reason"),
            {"file": "파일명", "status": "구분", "reason": "사유"},
            excluded_rows,
        )

        multi_sheet_rows = [
            (
                item.get("file", ""),
                item.get("total_sheets", 0),
                item.get("read_sheets", 0),
                item.get("note", ""),
            )
            for item in report.get("multi_sheet_details", [])
        ]
        multi_frame = ttk.Frame(notebook, padding=8)
        notebook.add(multi_frame, text="다중 시트")
        self._populate_tree(
            multi_frame,
            ("file", "total", "read", "note"),
            {"file": "파일명", "total": "전체 시트 수", "read": "읽은 시트 수", "note": "비고"},
            multi_sheet_rows,
        )

        quality_rows = [
            ("유효 사례", report.get("valid_cases", 0)),
            ("제외 행", report.get("excluded_rows", 0)),
            ("부서 누락", report.get("missing_departments", 0)),
            ("사용자 누락", report.get("missing_users", 0)),
            ("APC 누락", report.get("missing_apc", 0)),
            ("PC filter 누락", report.get("missing_pc_filter", 0)),
            ("UTMP 누락", report.get("missing_utmp", 0)),
        ]
        quality_frame = ttk.Frame(notebook, padding=8)
        notebook.add(quality_frame, text="데이터 품질")
        self._populate_tree(quality_frame, ("item", "value"), {"item": "항목", "value": "건수"}, quality_rows)

    def _populate_tree(
        self,
        parent: ttk.Frame,
        columns: tuple[str, ...],
        headings: dict[str, str],
        rows: list[tuple[object, ...]],
    ) -> None:
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        tree = ttk.Treeview(parent, columns=columns, show="headings")
        for column in columns:
            tree.heading(column, text=headings[column])
            tree.column(column, width=180 if column == "file" else 120, anchor="w")
        if rows:
            for row in rows:
                tree.insert("", tk.END, values=row)
        else:
            tree.insert("", tk.END, values=("표시할 내용 없음",) + ("",) * (len(columns) - 1))

        vsb = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=tree.yview)
        hsb = ttk.Scrollbar(parent, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

    def _set_busy(self, busy: bool) -> None:
        if busy:
            self.progress.start(10)
        else:
            self.progress.stop()
        state = ["disabled"] if busy else ["!disabled"]
        for child in self.root.winfo_children():
            self._set_state_recursive(child, state)
        # keep bottom bar usable
        self.progress.configure(mode="indeterminate")

    def _set_state_recursive(self, widget: tk.Widget, state: list[str]) -> None:
        try:
            if isinstance(widget, (ttk.Entry, ttk.Button, ttk.Combobox, ttk.Checkbutton, ttk.Spinbox)):
                widget.state(state)
        except Exception:
            pass
        for child in widget.winfo_children():
            self._set_state_recursive(child, state)

    def _refresh_year_combos(self) -> None:
        values = ["전체"] + [str(year) for year in self.year_values]
        self.year_from_combo.configure(values=values)
        self.year_to_combo.configure(values=values)
        if self.year_from_var.get() not in values:
            self.year_from_var.set("전체")
        if self.year_to_var.get() not in values:
            self.year_to_var.set("전체")

    def _collect_filters(self) -> SearchFilters:
        def parse_year(value: str) -> Optional[int]:
            value = value.strip()
            if not value or value == "전체":
                return None
            try:
                return int(value)
            except ValueError:
                return None

        return SearchFilters(
            year_from=parse_year(self.year_from_var.get()),
            year_to=parse_year(self.year_to_var.get()),
            department=self.department_var.get(),
            user=self.user_var.get(),
            require_apc=self.apc_var.get(),
            require_pc_filter=self.pc_filter_var.get(),
            require_utmp=self.utmp_var.get(),
        )

    def _run_search(self) -> None:
        if not self.engine.is_ready:
            messagebox.showwarning(APP_TITLE, "먼저 인덱스를 구축하거나 불러오세요.")
            return
        query = self.query_var.get().strip()
        if not query or query == QUERY_HINT:
            messagebox.showwarning(APP_TITLE, "검색어를 입력하세요.")
            return

        try:
            filters = self._collect_filters()
            results = self.engine.search(query, filters, top_n=self.top_n_var.get())
            self._show_results(results)
            self.status_var.set(f"검색 완료: {len(results)}건")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"검색 실패\n\n{exc}")

    def _show_results(self, results: list[dict[str, object]]) -> None:
        self.search_results = results
        self.tree.delete(*self.tree.get_children())
        self.detail_text.delete("1.0", tk.END)
        if not results:
            self.detail_text.insert(tk.END, "검색 결과가 없습니다.")
            return

        for result in results:
            case: MaintenanceCase = result["case"]
            self.tree.insert(
                "",
                tk.END,
                values=(
                    result["rank"],
                    self._format_percent(result["score"]),
                    case.year,
                    case.date_text,
                    case.department,
                    case.user,
                    self._shorten(case.issue_text, 28),
                    self._shorten(case.action_text, 40),
                    case.source_file,
                ),
            )
        self.tree.selection_set(self.tree.get_children()[0])
        self.tree.focus(self.tree.get_children()[0])
        self._show_selected_detail()

    def _shorten(self, text: str, limit: int) -> str:
        return text if len(text) <= limit else text[: limit - 1] + "…"

    def _show_selected_detail(self, event: object | None = None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        idx = self.tree.index(selection[0])
        if idx >= len(self.search_results):
            return
        result = self.search_results[idx]
        case: MaintenanceCase = result["case"]
        detail = {
            "순위": result["rank"],
            "관련도": self._format_percent(result["score"]),
            "키워드 일치": self._format_percent(result["keyword_score"]),
            "내용 유사도": self._format_percent(result["similarity_score"]),
            "연도": case.year,
            "월": case.month,
            "날짜": case.date_text,
            "부서": case.department,
            "사용자": case.user,
            "장애내용": case.issue_text,
            "조치내용": case.action_text,
            "APC": case.apc,
            "PC filter": case.pc_filter,
            "UTMP": case.utmp,
            "원본파일": case.source_file,
            "시트": case.source_sheet,
            "행번호": case.row_num,
            "시트제목": case.sheet_title,
        }
        self.detail_text.delete("1.0", tk.END)
        self.detail_text.insert(tk.END, json.dumps(detail, ensure_ascii=False, indent=2))

    @staticmethod
    def _format_percent(value: object) -> str:
        number = float(value)
        return f"{max(0.0, min(number, 1.0)) * 100:.1f}%"

    def _export_results(self) -> None:
        if not self.search_results:
            messagebox.showwarning(APP_TITLE, "먼저 검색을 실행하세요.")
            return

        output_path = filedialog.asksaveasfilename(
            title="검색 결과 엑셀 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel 통합 문서", "*.xlsx")],
            initialfile="유지보수_유사사례_검색결과.xlsx",
        )
        if not output_path:
            return

        try:
            from openpyxl import Workbook
            from openpyxl.styles import Alignment, Font, PatternFill
            from openpyxl.utils import get_column_letter

            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "유사사례 검색결과"
            headers = [
                "순위",
                "관련도",
                "키워드 일치",
                "내용 유사도",
                "연도",
                "월",
                "날짜",
                "부서",
                "사용자",
                "장애내용",
                "조치내용",
                "APC",
                "PC filter",
                "UTMP",
                "원본파일",
                "시트",
                "행번호",
            ]
            worksheet.append(headers)

            for result in self.search_results:
                case: MaintenanceCase = result["case"]
                worksheet.append(
                    [
                        result["rank"],
                        self._format_percent(result["score"]),
                        self._format_percent(result["keyword_score"]),
                        self._format_percent(result["similarity_score"]),
                        case.year,
                        case.month,
                        case.date_text,
                        case.department,
                        case.user,
                        case.issue_text,
                        case.action_text,
                        case.apc,
                        case.pc_filter,
                        case.utmp,
                        case.source_file,
                        case.source_sheet,
                        case.row_num,
                    ]
                )

            header_fill = PatternFill("solid", fgColor="D9EAF7")
            for cell in worksheet[1]:
                cell.font = Font(bold=True)
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")

            widths = [8, 12, 14, 14, 8, 8, 12, 18, 14, 42, 55, 9, 11, 9, 24, 16, 10]
            for index, width in enumerate(widths, start=1):
                worksheet.column_dimensions[get_column_letter(index)].width = width
            for row in worksheet.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
            workbook.save(output_path)
            self.status_var.set(f"엑셀 저장 완료: {output_path}")
            messagebox.showinfo(APP_TITLE, "검색 결과를 엑셀 파일로 저장했습니다.")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"엑셀 저장 실패\n\n{exc}")


def main() -> None:
    root = tk.Tk()
    try:
        ttk.Style().theme_use("clam")
    except Exception:
        pass
    MaintenanceSearchApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
