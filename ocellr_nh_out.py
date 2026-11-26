"""
꽁순이 OCellR - NH Out!
은행 거래내역서·증명서 PDF에서 표 안쪽 핵심 거래내역만 추출하여
미리보기 확인 및 엑셀(.xlsx)로 변환하는 PyQt6 데스크톱 애플리케이션입니다.
"""

from __future__ import annotations

from datetime import date
import importlib
import os
import re
import subprocess
import sys
import threading
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple
if (
    __name__ == "__main__"
    and not getattr(sys, "frozen", False)
    and sys.version_info[:2] != (3, 12)
    and os.environ.get("OCELLR_PY_REEXEC") != "1"
):
    try:
        _probe = subprocess.run(
            ["py", "-3.12", "-c", "import sys; print(sys.executable)"],
            capture_output=True,
            text=True,
            timeout=5,
        )
        _py312 = _probe.stdout.strip()
        if _probe.returncode == 0 and _py312:
            _env = dict(os.environ)
            _env["OCELLR_PY_REEXEC"] = "1"
            raise SystemExit(subprocess.call([_py312, *sys.argv], env=_env))
    except SystemExit:
        raise
    except Exception:
        pass

APP_TITLE = "꽁순이 OCellR v0.6 - NH Out!"


# =============================================================================
# 1. 필수 패키지 자동 점검 및 설치 (GUI 실행 전 PyQt6 확보)
# =============================================================================
def _pip_install(package_spec: str) -> Tuple[bool, str]:
    """현재 실행 중인 Python 인터프리터에 pip로 패키지를 설치합니다."""
    cmd = [sys.executable, "-m", "pip", "install", "--quiet", package_spec]
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
        if proc.returncode == 0:
            return True, ""
        err = (proc.stderr or proc.stdout or "알 수 없는 설치 오류").strip()
        last_line = err.splitlines()[-1] if err.splitlines() else err
        return False, last_line
    except Exception as exc:
        return False, str(exc)


def _ensure_pyqt6() -> None:
    """PyQt6가 없으면 자동 설치를 시도합니다."""
    try:
        import PyQt6  # noqa: F401
        return
    except ImportError:
        print("📦 'PyQt6' 설치 중...")
        ok, err = _pip_install("PyQt6")
        if not ok:
            print(f"❌ PyQt6 설치 실패: {err}")
            input("종료하려면 Enter 키를 누르세요...")
            sys.exit(1)


_ensure_pyqt6()

from PyQt6.QtCore import (  # noqa: E402
    Qt,
    QThread,
    pyqtSignal,
)
from PyQt6.QtGui import (  # noqa: E402
    QColor,
    QDragEnterEvent,
    QDropEvent,
    QFont,
    QIcon,
    QPainter,
    QPalette,
    QPen,
    QPixmap,
)
from PyQt6.QtWidgets import (  # noqa: E402
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


# =============================================================================
# 2. 데이터 모델 및 변환 옵션
# =============================================================================
@dataclass
class ConversionOptions:
    """PDF -> 엑셀 변환 옵션."""

    merge_into_single_sheet: bool = True
    force_ocr: bool = False
    normalize_dates: bool = True
    parse_numbers: bool = True
    dpi_scale: float = 2.6


@dataclass
class ExtractedTable:
    """한 페이지에서 추출된 단일 표 데이터."""

    page_num: int
    table_index: int
    rows: List[List[str]]
    source: str = "TEXT"  # TEXT-GRID, TEXT, OCR
    preview_rows: List[Tuple[List[str], bool]] = field(default_factory=list)


@dataclass
class ConversionResult:
    """전체 PDF 변환 결과."""

    pdf_path: str
    tables: List[ExtractedTable] = field(default_factory=list)
    merged_rows: List[List[str]] = field(default_factory=list)
    preview_rows: List[Tuple[List[str], bool]] = field(default_factory=list)
    page_count: int = 0
    used_ocr_pages: List[int] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


# =============================================================================
# 3. 텍스트 정제, 핵심 표 필터링 및 데이터 정규화 유틸리티
# =============================================================================
_TABLE_HEADER_KEYWORDS = (
    "거래일자",
    "거래일시",
    "거래일",
    "순번",
    "거래구분",
    "구분",
    "적요",
    "거래내용",
    "출금액",
    "입금액",
    "출금",
    "입금",
    "지급금액",
    "맡기신금액",
    "찾으신금액",
    "거래금액",
    "잔액",
    "거래후잔액",
    "거래점",
    "취급점",
    "기재내용",
    "기록사항",
    "거래기록사항",
    "상대계좌",
    "비고",
    "회차",
    "결제일",
    "수납일",
    "청구내역",
    "납입내역",
    "원금",
    "이자",
    "지연이자",
    "중도해지수수료",
    "취급수수료",
    "송무비",
    "부가세",
    "청구합계",
    "납입합계",
    "대출후잔액",
    "상태",
    "사용일자",
    "카드번호",
    "가맹점상품명",
    "가맹점",
    "이용금액",
    "승인번호",
    "이용구분",
    "결제일자",
    "본인가족",
    "사업자번호",
    "봉사료",
    "할부개월",
    "입금일자",
    "고객입금경로",
    "입금유형",
)

_OUTSIDE_NOISE_KEYWORDS = (
    "고객센터",
    "입출금거래내역",
    "거래내역조회",
    "거래내역증명서",
    "발급일시",
    "발급번호",
    "조회기간",
    "조회기준",
    "계좌번호",
    "예금주",
    "고객명",
    "상품명",
    "대출과목",
    "유의사항",
    "본증명서",
    "본확인증",
    "위내용은",
    "고객님이",
    "요청하신대로",
    "처리되었습니다",
    "인터넷뱅킹",
    "법적효력",
    "효력은없습니다",
    "대단히감사합니다",
    "편의를위하여",
    "농협은행",
    "하나은행",
    "국민은행",
    "신한은행",
    "우리은행",
    "기업은행",
    "합계",
    "소계",
    "누계",
    "총건수",
    "총합계",
)

_AMOUNT_HEADER_KEYWORDS = (
    "금액",
    "출금",
    "입금",
    "지급",
    "맡기신",
    "찾으신",
    "잔액",
    "원금",
    "이자",
    "수수료",
    "상환",
    "변제",
    "인출",
    "예입",
    "차변",
    "대변",
)

_DATE_HEADER_KEYWORDS = (
    "일자",
    "일시",
    "거래일",
    "결제일",
    "수납일",
    "사용일자",
    "결제일자",
    "입금일자",
    "날짜",
    "기산일",
    "만기일",
)

_SUMMARY_HEADER_REJECT_KEYWORDS = (
    "대출원금",
    "차량설정금액",
    "연체가지급금",
    "건수합계",
    "금액합계",
    "입금건수",
    "출금건수",
    "입금금액합계",
    "출금금액합계",
    "현재통장잔액",
    "현재지급가능잔액",
    "계좌번호",
    "예금주",
    "예금종류",
    "조회기간",
    "발급일시",
)

_DATE_FLEX_RE = re.compile(
    r"^\s*(\d{4})\s*[./\-년]\s*(\d{1,2})\s*[./\-월]\s*(\d{1,2})\s*일?\.?"
    r"(?:[\sT\(\[]*\d{1,2}:\d{2}(?::\d{2})?[\)\]]?)?\s*$"
)
_DATE_SHORT_YEAR_RE = re.compile(
    r"^\s*(\d{2})\s*[./\-]\s*(\d{1,2})\s*[./\-]\s*(\d{1,2})\.?"
    r"(?:[\sT\(\[]*\d{1,2}:\d{2}(?::\d{2})?[\)\]]?)?\s*$"
)
_DATE_COMPACT_RE = re.compile(
    r"^\s*((?:19|20)\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])"
    r"(?:\s+\d{2}:?\d{2}(?::?\d{2})?)?\s*$"
)
_ISO_DATE_RE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})$")
_NUMBER_RE = re.compile(r"^\s*([+-]?\d{1,3}(?:,\d{3})+|[+-]?\d+)(?:\.(\d+))?\s*원?\s*$")


def _clean_cell_text(val: Any) -> str:
    """셀 문자열의 줄바꿈 및 중복 공백을 정리합니다."""
    if val is None:
        return ""
    s = str(val).replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
    return re.sub(r"\s+", " ", s).strip()


def _compact_alnum(text: str) -> str:
    """비교용으로 공백과 특수문자를 제거한 소문자 문자열을 반환합니다."""
    return re.sub(r"[^0-9a-zA-Z가-힣]", "", _clean_cell_text(text)).lower()


def _cluster_coords(coords: Sequence[float], tolerance: float = 3.0) -> List[float]:
    """인접한 좌표들을 허용 오차 내에서 하나로 병합해 정렬된 대표 좌표 목록을 반환합니다."""
    if not coords:
        return []
    sorted_vals = sorted(float(v) for v in coords)
    clusters: List[List[float]] = [[sorted_vals[0]]]
    for v in sorted_vals[1:]:
        if abs(v - clusters[-1][-1]) <= tolerance:
            clusters[-1].append(v)
        else:
            clusters.append([v])
    return [sum(c) / len(c) for c in clusters]


def _trim_empty_rows_cols(
    rows: Sequence[Sequence[str]],
    drop_empty_cols: bool = True,
) -> List[List[str]]:
    """완전히 빈 행(및 옵션에 따라 빈 열)을 제거하고 모든 행의 열 길이를 맞춥니다."""
    cleaned_rows: List[List[str]] = []
    for row in rows:
        cleaned = [_clean_cell_text(c) for c in row]
        if any(cleaned):
            cleaned_rows.append(cleaned)
    if not cleaned_rows:
        return []

    max_cols = max(len(r) for r in cleaned_rows)
    if max_cols == 0:
        return []

    normalized = [r + [""] * (max_cols - len(r)) for r in cleaned_rows]
    if not drop_empty_cols:
        return normalized

    non_empty_col_indices = [
        c_idx
        for c_idx in range(max_cols)
        if any(row[c_idx] for row in normalized)
    ]
    if not non_empty_col_indices:
        return []

    return [[row[c_idx] for c_idx in non_empty_col_indices] for row in normalized]


def _is_likely_table_header_row(row: Sequence[str]) -> bool:
    """실제 거래내역 표의 컬럼 헤더 행인지 판별합니다 (상단 요약 박스 제외)."""
    non_empty = [_clean_cell_text(c) for c in row if _clean_cell_text(c)]
    if len(non_empty) < 3:
        return False

    compacts = [_compact_alnum(c) for c in non_empty]

    # 상단 계좌 요약 박스(입금건수합계, 계좌번호, 현재통장잔액 등) 배제
    for cp in compacts:
        if any(reject in cp for reject in _SUMMARY_HEADER_REJECT_KEYWORDS):
            return False

    # 날짜·금액·건수 같은 실제 데이터 값이 섞여 있으면 헤더가 아님
    for raw in non_empty:
        if re.search(r"\d{4}[./-]\d{1,2}[./-]\d{1,2}", raw):
            return False
        if re.search(r"\d{4}\s*년\s*\d{1,2}\s*월", raw):
            return False
        if raw.rstrip().endswith(":"):
            return False
        if re.match(r"^\s*[+-]?\d[\d,]*(?:\.\d+)?\s*(?:원|건)\s*$", raw):
            return False
        if _NUMBER_RE.match(raw):
            return False

    matched_cells = 0
    has_date_kw = False
    has_amount_kw = False
    for cp in compacts:
        if not cp:
            continue
        if any(kw in cp for kw in _TABLE_HEADER_KEYWORDS):
            matched_cells += 1
        else:
            cp_dedup = re.sub(r"(.)\1", r"\1", cp)
            if cp_dedup != cp and any(kw in cp_dedup for kw in _TABLE_HEADER_KEYWORDS):
                matched_cells += 1
        if any(kw in cp for kw in _DATE_HEADER_KEYWORDS):
            has_date_kw = True
        else:
            cp_dedup = re.sub(r"(.)\1", r"\1", cp)
            if cp_dedup != cp and any(kw in cp_dedup for kw in _DATE_HEADER_KEYWORDS):
                has_date_kw = True
        if any(kw in cp for kw in _AMOUNT_HEADER_KEYWORDS):
            has_amount_kw = True
        else:
            cp_dedup = re.sub(r"(.)\1", r"\1", cp)
            if cp_dedup != cp and any(kw in cp_dedup for kw in _AMOUNT_HEADER_KEYWORDS):
                has_amount_kw = True

    return matched_cells >= 3 or (matched_cells >= 2 and has_date_kw and has_amount_kw)


_SUMMARY_TOTAL_CONTAINS_KEYWORDS = (
    "합계",
    "소계",
    "누계",
    "총계",
    "월계",
    "일계",
    "total",
    "subtotal",
)


def _is_summary_total_row(row: Sequence[str]) -> bool:
    """어느 셀에든 '합계'·'소계'·'누계'·'총계' 등이 포함된 행은 무조건 제외합니다."""
    non_empty = [_clean_cell_text(c) for c in row if _clean_cell_text(c)]
    if not non_empty:
        return False
    if _is_likely_table_header_row(row):
        return False

    compacts = [_compact_alnum(c) for c in non_empty]

    # 1) 행 안의 어느 셀에든 합계/소계/누계/총계/월계/일계 문구가 포함되면 즉시 제외
    for cp in compacts:
        if not cp:
            continue
        if any(kw in cp for kw in _SUMMARY_TOTAL_CONTAINS_KEYWORDS):
            return True

    # 2) '합' | '계'처럼 인접 셀로 분리된 경우나 앞쪽 셀의 단독 '계' 표기 제외
    head_joined = "".join(compacts[:3])
    if head_joined == "계" or any(kw in head_joined for kw in _SUMMARY_TOTAL_CONTAINS_KEYWORDS):
        return True
    for cp in compacts[:3]:
        if cp in ("계", "총", "sum"):
            return True

    return False


def _is_outside_noise_row(row: Sequence[str]) -> bool:
    """표 바깥쪽의 제목/계좌요약/합계/안내문구/페이지번호 행인지 판별합니다."""
    non_empty = [c.strip() for c in row if c and c.strip()]
    if not non_empty:
        return True

    # 합계·소계·누계·총계 행은 열 개수와 상관없이 무조건 제외
    if _is_summary_total_row(row):
        return True

    joined_space = " ".join(non_empty)
    joined_compact = _compact_alnum("".join(non_empty))
    if not joined_compact:
        return True

    # 페이지 번호 형태 (예: 1 / 6, - 2 -, Page 1 of 3)
    if len(non_empty) <= 2 and re.match(
        r"^(?:page\s*)?\d+\s*(?:/|of|-)\s*\d+$|^-\s*\d+\s*-$",
        joined_space.strip(),
        re.IGNORECASE,
    ):
        return True

    # 확인증/안내문/감사문구 등은 열 개수와 무관하게 즉시 제외
    footer_notice_keywords = (
        "본확인증",
        "본증명서",
        "고객님이",
        "요청하신대로",
        "처리되었습니다",
        "인터넷뱅킹",
        "법적효력",
        "효력은없습니다",
        "대단히감사합니다",
        "편의를위하여",
        "유의사항",
        "위내용은",
    )
    if any(kw in joined_compact for kw in footer_notice_keywords):
        return True

    # 날짜가 포함된 실제 거래 데이터 행은 보존 (단, 발급일시/조회기간 행은 제외)
    if re.search(r"\d{4}[./-]\d{1,2}[./-]\d{1,2}", joined_space):
        if any(kw in joined_compact for kw in ("발급일시", "조회기간", "조회기준", "출력일시")):
            return True
        return False

    if len(non_empty) <= 3 and any(kw in joined_compact for kw in _OUTSIDE_NOISE_KEYWORDS):
        return True

    if any(
        joined_compact.startswith(kw)
        for kw in (
            "합계",
            "소계",
            "누계",
            "총건수",
            "총합계",
            "유의사항",
            "발급일시",
            "조회기간",
            "계좌번호",
            "입출금거래내역",
        )
    ):
        return True

    return False


def _get_amount_col_indices(header_row: Sequence[str]) -> List[int]:
    """헤더 행에서 금액/잔액 관련 열 인덱스를 찾습니다."""
    indices: List[int] = []
    for idx, cell in enumerate(header_row):
        cp = _compact_alnum(cell)
        if cp and any(kw in cp for kw in _AMOUNT_HEADER_KEYWORDS):
            indices.append(idx)
    return indices


def _get_date_col_indices(header_row: Sequence[str]) -> List[int]:
    """헤더 행에서 날짜/일시 관련 열 인덱스를 찾습니다."""
    indices: List[int] = []
    for idx, cell in enumerate(header_row):
        cp = _compact_alnum(cell)
        if cp and any(kw in cp for kw in _DATE_HEADER_KEYWORDS):
            indices.append(idx)
    return indices


def _is_amount_value(text: str) -> bool:
    """셀 값이 유효한 금액/숫자인지 판별합니다."""
    s = _clean_cell_text(text)
    if not s:
        return False
    return bool(_NUMBER_RE.match(s))


def _row_has_amount(row: Sequence[str], amount_col_indices: Sequence[int]) -> bool:
    """행에 유효한 금액 값이 하나 이상 있는지 판별합니다."""
    if amount_col_indices:
        for c_idx in amount_col_indices:
            if c_idx < len(row):
                cell_val = row[c_idx]
                if _is_amount_value(cell_val) or any(_is_amount_value(tok) for tok in cell_val.split()):
                    return True
        return False
    return any(_is_amount_value(cell) for cell in row)


def _filter_to_core_table_rows(
    rows: List[List[str]],
    preserve_columns: bool = False,
    deleted_out: Optional[List[Tuple[List[str], bool]]] = None,
) -> List[List[str]]:
    """표 바깥쪽 상·하단 메타 정보와 금액이 없는 불필요한 행을 제거합니다."""
    cleaned = _trim_empty_rows_cols(rows, drop_empty_cols=not preserve_columns)
    if not cleaned:
        return []

    # 1) 실제 표 헤더 행 위치를 찾아 그 위쪽(문서 제목, 계좌 요약 박스 등) 제거
    header_idx: Optional[int] = None
    for idx, r in enumerate(cleaned[: min(len(cleaned), 15)]):
        if _is_likely_table_header_row(r):
            header_idx = idx
            break
    if deleted_out is not None and header_idx is not None and header_idx > 0:
        for r in cleaned[:header_idx]:
            deleted_out.append((r, True))

    if header_idx is not None and header_idx > 0:
        cleaned = cleaned[header_idx:]

    # 2) 하단부터 역순으로 표 바깥쪽 안내/합계/페이지번호 행 제거
    while len(cleaned) > 1 and _is_outside_noise_row(cleaned[-1]):
        cleaned.pop()

    # 3) 헤더가 없는 이어지는 페이지의 상단 노이즈 행 제거
    while (
        len(cleaned) > 1
        and not _is_likely_table_header_row(cleaned[0])
        and _is_outside_noise_row(cleaned[0])
    ):
        cleaned.pop(0)

    # 4) 헤더 행을 제외한 데이터 행 중 안내문구 또는 금액이 없는 행 제거
    has_header = bool(cleaned and _is_likely_table_header_row(cleaned[0]))
    amount_cols = _get_amount_col_indices(cleaned[0]) if has_header else []
    filtered_rows: List[List[str]] = []
    for r_idx, row in enumerate(cleaned):
        if r_idx == 0 and has_header:
            filtered_rows.append(row)
            continue
        if _is_outside_noise_row(row):
            continue
        if not _row_has_amount(row, amount_cols):
            continue
        filtered_rows.append(row)
    if deleted_out is not None:
        kept_set = {id(r) for r in filtered_rows}
        for r in cleaned:
            deleted_out.append((r, id(r) not in kept_set))

    # 이전 페이지 열 경계를 상속한 경우(preserve_columns=True)에는 빈 열을 유지해 열 밀림 방지
    return _trim_empty_rows_cols(filtered_rows, drop_empty_cols=not preserve_columns)


def _select_main_data_tables(tables: List[ExtractedTable]) -> List[ExtractedTable]:
    """한 페이지에서 추출된 여러 표 후보 중 핵심 거래내역 표만 선택합니다."""
    valid = [t for t in tables if t.rows and len(t.rows) >= 1 and len(t.rows[0]) >= 2]
    if not valid:
        return []
    if len(valid) == 1:
        return valid

    def _score(tbl: ExtractedTable) -> Tuple[int, int, int]:
        has_hdr = 1 if _is_likely_table_header_row(tbl.rows[0]) else 0
        n_rows = len(tbl.rows)
        n_cols = len(tbl.rows[0])
        return (has_hdr, n_rows, n_cols)

    best = max(valid, key=_score)
    return [best]


def _parse_normalized_date(text: str, allow_compact: bool = False) -> Optional[str]:
    """다양한 은행 날짜/일시 문자열을 YYYY-MM-DD 형식으로 정규화합니다."""
    s = _clean_cell_text(text)
    if not s:
        return None

    m = _DATE_FLEX_RE.match(s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return date(y, mo, d).strftime("%Y-%m-%d")
        except ValueError:
            return None

    m_short = _DATE_SHORT_YEAR_RE.match(s)
    if m_short:
        y2, mo, d = int(m_short.group(1)), int(m_short.group(2)), int(m_short.group(3))
        y = 2000 + y2 if y2 <= 79 else 1900 + y2
        try:
            return date(y, mo, d).strftime("%Y-%m-%d")
        except ValueError:
            return None

    if allow_compact:
        m_compact = _DATE_COMPACT_RE.match(s)
        if m_compact:
            y, mo, d = int(m_compact.group(1)), int(m_compact.group(2)), int(m_compact.group(3))
            try:
                return date(y, mo, d).strftime("%Y-%m-%d")
            except ValueError:
                return None
        m_compact6 = re.match(r"^\s*(\d{2})(\d{2})(\d{2})\s*$", s)
        if m_compact6:
            y2, mo, d = int(m_compact6.group(1)), int(m_compact6.group(2)), int(m_compact6.group(3))
            y = 2000 + y2 if y2 <= 79 else 1900 + y2
            try:
                return date(y, mo, d).strftime("%Y-%m-%d")
            except ValueError:
                return None

    return None


def _normalize_cell_value(
    text: str,
    options: ConversionOptions,
    is_date_col: bool = False,
) -> str:
    """날짜 포맷(YYYY-MM-DD) 등 셀 표시 값을 정규화합니다."""
    s = _clean_cell_text(text)
    if not s:
        return ""
    if options.normalize_dates:
        parsed_date = _parse_normalized_date(s, allow_compact=is_date_col)
        if parsed_date:
            return parsed_date
    return s


def _parse_excel_cell_value(
    text: str,
    is_header: bool,
    options: ConversionOptions,
    is_date_col: bool = False,
) -> Tuple[Any, Optional[str]]:
    """엑셀 저장 시 날짜/숫자/문자 타입과 표시 서식을 결정합니다."""
    s = _normalize_cell_value(text, options, is_date_col=is_date_col)
    if is_header or not s:
        return s, None

    if options.normalize_dates:
        m_iso = _ISO_DATE_RE.match(s)
        if m_iso:
            try:
                dt_val = date(int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3)))
                return dt_val, "yyyy-mm-dd"
            except ValueError:
                pass

    if options.parse_numbers:
        # 0으로 시작하는 코드/계좌번호/순번(예: 00123)은 문자열 유지
        raw_digits = s.replace(",", "").replace("원", "").strip()
        if not (len(raw_digits) > 1 and raw_digits.startswith("0") and "." not in raw_digits):
            m = _NUMBER_RE.match(s)
            if m:
                int_part = m.group(1).replace(",", "")
                frac_part = m.group(2)
                try:
                    if frac_part is not None:
                        return float(f"{int_part}.{frac_part}"), "#,##0.##"
                    return int(int_part), "#,##0"
                except ValueError:
                    pass
    return s, None


def _rows_similar_header(row_a: Sequence[str], row_b: Sequence[str]) -> bool:
    """두 행이 동일한 표 헤더의 반복인지 비교합니다."""
    if not row_a or not row_b:
        return False
    comp_a = [_compact_alnum(c) for c in row_a if c.strip()]
    comp_b = [_compact_alnum(c) for c in row_b if c.strip()]
    if not comp_a or not comp_b:
        return False
    if comp_a == comp_b:
        return True
    overlap = len(set(comp_a) & set(comp_b))
    return overlap >= max(3, min(len(comp_a), len(comp_b)) - 1)


def _merge_tables_into_single_grid(
    tables: List[ExtractedTable],
    options: ConversionOptions,
) -> List[List[str]]:
    """여러 페이지에서 추출된 표를 하나의 연속된 통합 거래내역 표로 병합합니다."""
    if not tables:
        return []

    max_cols = max((len(t.rows[0]) for t in tables if t.rows), default=0)
    if max_cols == 0:
        return []

    merged: List[List[str]] = []
    first_header: Optional[List[str]] = None
    amount_cols: List[int] = []
    date_cols: List[int] = []

    for tbl in tables:
        if not tbl.rows:
            continue
        # 페이지 표가 first_header보다 열이 많고 빈 헤더 열이 끼어 있으면 앞 열에 병합
        tbl_rows = tbl.rows
        if first_header and _clean_cell_text(first_header[0]):
            while tbl_rows and len(tbl_rows[0]) > 1 and all(not _clean_cell_text(r[0]) for r in tbl_rows if r):
                tbl_rows = [r[1:] for r in tbl_rows]
        if first_header and not _clean_cell_text(first_header[-1]):
            fh_non_empty = sum(1 for c in first_header if _clean_cell_text(c))
            keep_cs = [ci for ci in range(len(tbl_rows[0])) if any(_clean_cell_text(r[ci]) for r in tbl_rows if ci < len(r))]
            expected_cols = fh_non_empty + (0 if _clean_cell_text(first_header[0]) else 1)
            if len(keep_cs) == expected_cols:
                tbl_rows = [[r[ci] if ci < len(r) else "" for ci in keep_cs] for r in tbl_rows]
        if _is_likely_table_header_row(tbl_rows[0]) and (
            not first_header or len(tbl_rows[0]) >= len(first_header)
        ):
            drop_cs = [
                ci for ci in range(1, len(tbl_rows[0]))
                if not _clean_cell_text(tbl_rows[0][ci]) and _clean_cell_text(tbl_rows[0][ci - 1])
            ]
            if drop_cs:
                new_rows: List[List[str]] = []
                for r in tbl_rows:
                    rr = list(r)
                    for ci in reversed(drop_cs):
                        if ci < len(rr):
                            if not _clean_cell_text(rr[ci - 1]) and _clean_cell_text(rr[ci]):
                                rr[ci - 1] = rr[ci]
                            rr.pop(ci)
                    new_rows.append(rr)
                tbl_rows = new_rows
        for r_idx, raw_row in enumerate(tbl_rows):
            padded_row = list(raw_row)
            if len(padded_row) < max_cols:
                padded_row += [""] * (max_cols - len(padded_row))
            elif len(padded_row) > max_cols:
                padded_row = padded_row[:max_cols]

            if not any(_clean_cell_text(c) for c in padded_row):
                continue

            if not merged:
                clean_first = [_clean_cell_text(c) for c in padded_row]
                if _is_likely_table_header_row(clean_first):
                    clean_first = [
                        re.sub(r"([가-힣A-Za-z])\1", r"\1", c)
                        if re.search(r"([가-힣])\1.*([가-힣])\2", c)
                        else c
                        for c in clean_first
                    ]
                    first_header = clean_first
                    amount_cols = _get_amount_col_indices(first_header)
                    date_cols = _get_date_col_indices(first_header)
                    merged.append(clean_first)
                    continue

            norm_row = [
                _normalize_cell_value(c, options, is_date_col=(c_idx in date_cols))
                for c_idx, c in enumerate(padded_row)
            ]
            if (
                merged
                and _parse_normalized_date(merged[-1][0])
                and len(norm_row) > 1
                and not _clean_cell_text(norm_row[0])
                and _parse_normalized_date(norm_row[1])
            ):
                norm_row = norm_row[1:] + [""]

            # 2페이지 이후 반복되는 헤더 행 자동 스킵
            if r_idx == 0 and first_header and _rows_similar_header(first_header, norm_row):
                continue
            if (
                merged
                and r_idx == 0
                and _is_likely_table_header_row(norm_row)
                and _rows_similar_header(merged[0], norm_row)
            ):
                continue

            # 안내문구 및 금액 없는 행 최종 필터링
            if _is_outside_noise_row(norm_row):
                continue

            if (
                merged
                and not _is_likely_table_header_row(merged[-1])
                and not _clean_cell_text(norm_row[date_cols[0]] if date_cols else norm_row[0])
                and sum(1 for c in norm_row if _clean_cell_text(c)) <= 3
                and _clean_cell_text(merged[-1][date_cols[0]] if date_cols else merged[-1][0])
            ):
                for ci in range(len(norm_row)):
                    val = _clean_cell_text(norm_row[ci])
                    if val:
                        prev_val = _clean_cell_text(merged[-1][ci])
                        merged[-1][ci] = f"{prev_val} {val}".strip() if prev_val else val
                continue
            if not _row_has_amount(norm_row, amount_cols):
                continue
            non_empty_vals = [_clean_cell_text(c) for c in norm_row if _clean_cell_text(c)]
            if non_empty_vals and all(v in ("0", "0.0", "-") for v in non_empty_vals):
                continue
            if any("고객센터" in _clean_cell_text(c) for c in norm_row):
                continue
            merged.append(norm_row)

    result_rows = _trim_empty_rows_cols(merged, drop_empty_cols=True)
    result_rows = _merge_two_line_card_rows(result_rows, options)
    result_rows = _merge_pg_submall_into_merchant(result_rows)
    result_rows = _split_daegu_interest_columns(result_rows)
    for ri in range(1, len(result_rows)):
        r = result_rows[ri]
        if (
            len(r) > 1
            and not _clean_cell_text(r[0])
            and _parse_normalized_date(r[1])
            and _parse_normalized_date(result_rows[ri - 1][0])
        ):
            result_rows[ri] = r[1:] + [""]
    result_rows = _trim_empty_rows_cols(result_rows, drop_empty_cols=True)
    return _fix_lotte_split_prefixes(result_rows)


def _merge_two_line_card_rows(
    rows: List[List[str]],
    options: ConversionOptions,
) -> List[List[str]]:
    """씨티카드 등 1건의 거래가 2줄(상단/하단)로 나뉜 표를 1줄(12열)로 병합합니다."""
    if len(rows) < 2 or len(rows[0]) != 6:
        return rows
    h0 = "".join(_compact_alnum(c) for c in rows[0])
    if not ("사용일자" in h0 and "이용금액" in h0):
        return rows

    h1 = "".join(_compact_alnum(c) for c in rows[1])
    if "결제일자" in h1:
        out: List[List[str]] = [rows[0] + rows[1]]
        i = 2
    else:
        out = [rows[0] + ["결제일자", "본인/가족", "사업자번호", "부가세", "봉사료", "할부개월"]]
        i = 1
    while i < len(rows):
        r1 = rows[i]
        if i + 1 < len(rows) and _compact_alnum(rows[i + 1][1]) in ("본인", "가족"):
            r2 = list(rows[i + 1])
            if options.normalize_dates and r2[0]:
                r2[0] = _normalize_cell_value(r2[0], options, is_date_col=True)
            out.append(r1 + r2)
            i += 2
        else:
            out.append(r1 + [""] * 6)
            i += 1
    return out


def _merge_pg_submall_into_merchant(rows: List[List[str]]) -> List[List[str]]:
    """롯데카드 등에서 PG하위몰명이 있으면 가맹점명을 'PG하위몰명(가맹점명)'으로 반영합니다."""
    if not rows or len(rows[0]) < 7:
        return rows
    sample = rows[1] if len(rows) > 1 else rows[0]
    if len(sample) == 10 and re.search(r"\d{4}-\d{4}-\*{4}-", sample[1]) and re.search(r"\d{3}-\d{2}-", sample[7]):
        out10: List[List[str]] = []
        for r_idx, row in enumerate(rows):
            rr = list(row)
            if r_idx == 0 and _is_likely_table_header_row(rr):
                cleaned_hdr = [re.sub(r"([가-힣])\1", r"\1", _clean_cell_text(c)) for c in rr[:9]]
                out10.append(cleaned_hdr)
                continue
            merchant = _clean_cell_text(rr[8])
            submall = _clean_cell_text(rr[9])
            if submall and merchant:
                rr[8] = f"{submall}({merchant})"
            out10.append(rr[:9])
        return out10
    if len(sample) != 7:
        return rows
    if not (re.search(r"\d{4}-\d{4}-\*{4}-", sample[1]) and re.search(r"\d{3}-\d{2}-", sample[4])):
        return rows

    merged_col2_n = 0
    for r in rows[1:]:
        if len(r) >= 3:
            n_tok = len(_clean_cell_text(r[2]).split())
            if n_tok in (3, 4):
                merged_col2_n = n_tok
                break
    out: List[List[str]] = []
    for r_idx, row in enumerate(rows):
        rr = list(row)
        if r_idx == 0 and _is_likely_table_header_row(rr):
            base = [re.sub(r"([가-힣])\1", r"\1", _clean_cell_text(c)) for c in rr[:6]]
            if merged_col2_n == 4:
                base = base[:2] + ["승인번호", "매출금액", "부가세", "할부"] + base[3:]
            elif merged_col2_n == 3:
                base = base[:2] + ["승인번호", "매출금액", "할부"] + base[3:]
            out.append(base)
            continue
        merchant = _clean_cell_text(rr[5])
        submall = _clean_cell_text(rr[6])
        if submall and merchant:
            rr[5] = f"{submall}({merchant})"
        base = rr[:6]
        if merged_col2_n in (3, 4):
            parts = _clean_cell_text(base[2]).split()
            if len(parts) == merged_col2_n:
                base = base[:2] + parts + base[3:]
            else:
                base = base[:2] + [base[2]] + [""] * (merged_col2_n - 1) + base[3:]
        out.append(base)
    return out


def _split_daegu_interest_columns(rows: List[List[str]]) -> List[List[str]]:
    """대구은행 이자납입내역서의 합쳐진 열(통화/거래금액/이자금액, 거래점/지점명/이자계산내용)을 분리합니다."""
    if len(rows) < 2 or len(rows[0]) != 7:
        return rows
    h0 = _compact_alnum(rows[0][0])
    h4 = _compact_alnum(rows[0][4])
    h6 = _compact_alnum(rows[0][6])
    if not ("실행명" in h0 and "이자금액" in h4 and "이자계산내용" in h6):
        return rows

    out: List[List[str]] = [[
        "실행명세번호", "세부명세번호", "거래일자", "거래구분",
        "통화코드", "거래금액", "이자금액", "이자계산대상일수",
        "거래점번호", "지점명", "이자계산내용",
    ]]
    for row in rows[1:]:
        rr = list(row)
        c4_toks = _clean_cell_text(rr[4]).split()
        curr = "KRW" if ("KR" in c4_toks or "KRW" in c4_toks) else ""
        num_toks = [t for t in c4_toks if t not in ("KR", "W", "KRW")]
        if len(num_toks) == 2:
            amt1, amt2 = num_toks[0], num_toks[1]
        elif len(num_toks) == 3:
            amt1, amt2 = num_toks[0], num_toks[1] + num_toks[2]
        elif len(num_toks) == 4:
            amt1, amt2 = num_toks[0] + num_toks[2], num_toks[1] + num_toks[3]
        else:
            amt1, amt2 = " ".join(num_toks), ""

        c6_parts = _clean_cell_text(rr[6]).split(None, 2)
        br_no = c6_parts[0] if len(c6_parts) > 0 else ""
        br_name = c6_parts[1] if len(c6_parts) > 1 else ""
        calc_desc = c6_parts[2] if len(c6_parts) > 2 else ""

        out.append([
            rr[0], rr[1], rr[2], rr[3],
            curr, amt1, amt2, rr[5],
            br_no, br_name, calc_desc,
        ])
    return out


def _fix_lotte_split_prefixes(rows: List[List[str]]) -> List[List[str]]:
    """롯데카드 페이지 경계에서 앞부분이 잘린 카드번호(294)·사업자번호(01095)의 접두사를 복원합니다."""
    if len(rows) < 2 or len(rows[0]) != 8:
        return rows
    card_prefix = ""
    biz_by_merchant: Dict[str, str] = {}
    for r in rows[1:]:
        m_card = re.match(r"^(\d{4}-\d{4}-\*{4}-)\s*\d{3,4}$", _clean_cell_text(r[1]))
        if m_card and not card_prefix:
            card_prefix = m_card.group(1)
        merch = _clean_cell_text(r[7])
        biz = _clean_cell_text(r[6])
        if merch and re.match(r"^\d{3}-\d{2}-\s*\d{5}$", biz):
            biz_by_merchant[merch] = biz
    if not card_prefix:
        return rows
    for r in rows[1:]:
        if re.match(r"^\d{3,4}$", _clean_cell_text(r[1])):
            r[1] = f"{card_prefix} {_clean_cell_text(r[1])}"
        merch = _clean_cell_text(r[7])
        if re.match(r"^\d{5}$", _clean_cell_text(r[6])) and merch in biz_by_merchant:
            r[6] = biz_by_merchant[merch]
    return rows


# =============================================================================
# 4. 의존성 부트스트랩 및 OCR 엔진 관리 (RapidOCR 한국어 + 안전 폴백)
# =============================================================================
class DependencyManager:
    """PDF/이미지/엑셀/OCR 라이브러리를 지연 설치 및 로드합니다."""

    REQUIRED_PACKAGES: List[Tuple[str, str]] = [
        ("pdfplumber", "pdfplumber"),
        ("openpyxl", "openpyxl"),
        ("pypdfium2", "pypdfium2"),
        ("numpy", "numpy"),
        ("cv2", "opencv-python-headless"),
        ("PIL", "Pillow"),
    ]

    def __init__(self) -> None:
        self.ocr_engine: Any = None
        self.ocr_engine_name: str = "NONE"
        self.ocr_available: bool = False
        self._ocr_initialized: bool = False

    def ensure_core_packages(
        self,
        log_cb: Callable[[str], None],
    ) -> Tuple[bool, List[str]]:
        """핵심 라이브러리와 선택적 RapidOCR을 점검·설치합니다."""
        warnings: List[str] = []
        py_ver = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
        arch = "64bit" if sys.maxsize > 2**32 else "32bit"
        log_cb(f"ℹ 실행 환경: Python {py_ver} ({arch})")

        for mod_name, pip_name in self.REQUIRED_PACKAGES:
            try:
                importlib.import_module(mod_name)
            except ImportError:
                log_cb(f"📦 '{pip_name}' 설치 중...")
                ok, err = _pip_install(pip_name)
                if not ok:
                    msg = f"필수 패키지 '{pip_name}' 설치 실패: {err}"
                    log_cb(f"❌ {msg}")
                    return False, [msg]
                log_cb(f"✅ '{pip_name}' 설치 완료")

        # RapidOCR 점검 (Python 3.15 등 미지원 환경에서는 실패해도 앱이 멈추지 않도록 폴백)
        if not self._try_import_rapidocr():
            log_cb("📦 'rapidocr' 설치 확인 중...")
            ok, err = _pip_install("rapidocr")
            if ok:
                # 추론 백엔드 onnxruntime 설치 시도
                if not self._has_module("onnxruntime") and not self._has_module("openvino"):
                    log_cb("📦 'onnxruntime' 설치 중...")
                    ok_ort, err_ort = _pip_install("onnxruntime")
                    if not ok_ort:
                        log_cb(f"⚠ 'onnxruntime' 설치 실패 ({err_ort}), 'openvino' 대체 시도...")
                        _pip_install("openvino")
            else:
                warn = f"'rapidocr' 설치 건너뜀 ({err}) - 텍스트 PDF는 정상 변환됩니다."
                log_cb(f"⚠ {warn}")
                warnings.append(warn)

        return True, warnings

    @staticmethod
    def _has_module(name: str) -> bool:
        try:
            importlib.import_module(name)
            return True
        except ImportError:
            return False

    @staticmethod
    def _try_import_rapidocr() -> bool:
        for mod in ("rapidocr", "rapidocr_onnxruntime"):
            try:
                importlib.import_module(mod)
                return True
            except ImportError:
                continue
        return False

    def get_ocr_engine(self, log_cb: Callable[[str], None]) -> Any:
        """RapidOCR 한국어 엔진을 최초 1회 초기화해 반환합니다."""
        if self._ocr_initialized:
            return self.ocr_engine
        self._ocr_initialized = True

        # 1) 신형 rapidocr 패키지 (한국어 모델 지원)
        try:
            from rapidocr import EngineType, LangRec, OCRVersion, RapidOCR

            engine_type = (
                EngineType.ONNXRUNTIME
                if self._has_module("onnxruntime")
                else EngineType.OPENVINO
            )
            log_cb("🔍 RapidOCR 한국어 모델 초기화 중...")
            try:
                self.ocr_engine = RapidOCR(
                    params={
                        "Det.engine_type": engine_type,
                        "Cls.engine_type": engine_type,
                        "Rec.engine_type": engine_type,
                        "Rec.lang_type": LangRec.KOREAN,
                        "Rec.ocr_version": OCRVersion.PPOCRV5,
                        "Rec.model_type": "mobile",
                    }
                )
            except Exception:
                try:
                    self.ocr_engine = RapidOCR(
                        params={
                            "Rec.lang_type": LangRec.KOREAN,
                            "Rec.ocr_version": OCRVersion.PPOCRV5,
                            "Rec.model_type": "mobile",
                        }
                    )
                except Exception:
                    self.ocr_engine = RapidOCR()
            self.ocr_engine_name = "RapidOCR-KR"
            self.ocr_available = True
            log_cb("✅ RapidOCR 한국어 엔진 준비 완료")
            return self.ocr_engine
        except Exception as exc:
            log_cb(f"ℹ rapidocr 초기화 폴백: {exc}")

        # 2) 구형 rapidocr_onnxruntime 패키지
        try:
            from rapidocr_onnxruntime import RapidOCR as LegacyRapidOCR

            self.ocr_engine = LegacyRapidOCR()
            self.ocr_engine_name = "RapidOCR-ONNX"
            self.ocr_available = True
            log_cb("✅ RapidOCR(ONNX) 엔진 준비 완료")
            return self.ocr_engine
        except Exception:
            pass

        self.ocr_engine = None
        self.ocr_engine_name = "NONE"
        self.ocr_available = False
        log_cb("⚠ 사용 가능한 OCR 엔진이 없어 스캔 이미지 페이지 인식은 제한됩니다. (Python 3.12~3.13 권장)")
        return None


# =============================================================================
# 5. PDF 표 추출 엔진 (텍스트 격자 복원 + 이미지 표 OCR)
# =============================================================================
def _is_password_error(exc: Exception) -> bool:
    """예외가 PDF 비밀번호/암호화 관련 오류인지 판별합니다."""
    cls_name = type(exc).__name__.lower()
    msg = str(exc).lower()
    keywords = ("password", "encrypt", "decrypt", "incorrect password", "pdfpasswordincorrect")
    return any(k in cls_name or k in msg for k in keywords)


def _find_header_line_and_tokens(
    words: Sequence[Dict[str, Any]],
    v_lines: Optional[Sequence[Tuple[float, float, float]]] = None,
) -> Optional[Tuple[float, float, List[Tuple[float, float, str]]]]:
    """페이지 단어 목록에서 실제 거래내역 표 헤더 줄의 Y범위와 각 컬럼 토큰(x0, x1, text)을 찾습니다."""
    if not words:
        return None

    sorted_words = sorted(words, key=lambda w: (float(w["top"]), float(w["x0"])))
    lines: List[List[Dict[str, Any]]] = []
    for w in sorted_words:
        w_top = float(w["top"])
        if not lines or abs(w_top - float(lines[-1][0]["top"])) > 4.5:
            lines.append([w])
        else:
            lines[-1].append(w)

    for line in lines:
        line.sort(key=lambda w: float(w["x0"]))
        line_top = min(float(w["top"]) for w in line)
        line_bot = max(float(w["bottom"]) for w in line)

        # 헤더 줄을 가로지르는 세로선 X좌표 수집 (세로선 양쪽의 단어는 절대 병합하지 않음)
        line_vxs: List[float] = []
        if v_lines:
            line_vxs = [
                vx
                for vx, vy0, vy1 in v_lines
                if vy0 <= (line_bot + 4.0) and vy1 >= (line_top - 4.0)
            ]

        # 같은 컬럼명 안의 가까운 단어(예: '거래 후 잔액')만 세로선이 없을 때 병합
        tokens: List[Tuple[float, float, str]] = []
        for w in line:
            wx0, wx1, txt = float(w["x0"]), float(w["x1"]), _clean_cell_text(w["text"])
            if not txt:
                continue
            if tokens and (wx0 - tokens[-1][1]) <= 4.5:
                prev_x1 = tokens[-1][1]
                has_vline_between = any((prev_x1 - 0.5) <= vx <= (wx0 + 0.5) for vx in line_vxs)
                if not has_vline_between:
                    px0, _, ptxt = tokens[-1]
                    tokens[-1] = (px0, wx1, f"{ptxt} {txt}")
                    continue
            tokens.append((wx0, wx1, txt))

        if _is_likely_table_header_row([t[2] for t in tokens]):
            return line_top, line_bot, tokens

    return None


def _build_col_bounds_from_header_and_vlines(
    tx0: float,
    tx1: float,
    header_tokens: List[Tuple[float, float, str]],
    inner_vs: Sequence[float],
) -> List[float]:
    """헤더 컬럼 위치와 세로선을 결합해 출금/입금 및 다열 표의 정확한 열 경계(col_xs)를 만듭니다."""
    v_bounds = _cluster_coords([tx0, tx1] + list(inner_vs), tolerance=3.0)
    # 실제 세로선이 헤더 열 수 이상으로 뚜렷하게 존재하면 세로선 격자를 최우선 사용
    if len(v_bounds) - 1 >= len(header_tokens) and len(inner_vs) >= max(3, len(header_tokens) - 1):
        return v_bounds

    bounds: List[float] = [tx0]
    sorted_vs = sorted(inner_vs)
    for i in range(len(header_tokens) - 1):
        left_x1 = header_tokens[i][1]
        right_x0 = header_tokens[i + 1][0]
        # 두 헤더 토큰 사이에 실제 세로선이 있으면 우선 사용
        candidates = [vx for vx in sorted_vs if (left_x1 - 2.0) <= vx <= (right_x0 + 2.0)]
        if candidates:
            mid_target = (left_x1 + right_x0) / 2.0
            best_v = min(candidates, key=lambda vx: abs(vx - mid_target))
            bounds.append(best_v)
        else:
            bounds.append((left_x1 + right_x0) / 2.0)
    bounds.append(tx1)
    return bounds


def _detect_date_col_index(
    words: Sequence[Dict[str, Any]],
    col_xs: Sequence[float],
    ty0: float,
    ty1: float,
) -> int:
    """헤더가 없는 페이지에서 날짜 문자열이 가장 많이 등장하는 열 인덱스를 자동 탐지합니다."""
    n_cols = len(col_xs) - 1
    if n_cols <= 0:
        return 0
    counts = [0] * n_cols
    for w in words:
        txt = _clean_cell_text(w["text"])
        if not re.match(r"^\d{4}[./-]\d{1,2}[./-]\d{1,2}", txt):
            continue
        cx = (float(w["x0"]) + float(w["x1"])) / 2.0
        cy = (float(w["top"]) + float(w["bottom"])) / 2.0
        if not (ty0 - 2.0 <= cy <= ty1 + 2.0):
            continue
        for ci in range(n_cols):
            if col_xs[ci] - 2.0 <= cx <= col_xs[ci + 1] + 2.0:
                counts[ci] += 1
                break
    best_ci = max(range(n_cols), key=lambda i: counts[i])
    return best_ci if counts[best_ci] > 0 else 0


def _split_row_ys_by_dates(
    row_ys: List[float],
    words: Sequence[Dict[str, Any]],
    date_col_x0: float,
    date_col_x1: float,
) -> List[float]:
    """가로선이 누락되어 한 행에 2건 이상의 거래 날짜가 뭉친 경우 날짜 Y좌표로 행을 분할합니다."""
    if len(row_ys) < 2:
        return row_ys

    date_words = [
        w
        for w in words
        if (date_col_x0 - 3.0) <= ((float(w["x0"]) + float(w["x1"])) / 2.0) <= (date_col_x1 + 3.0)
        and re.match(r"^\d{4}[./-]\d{1,2}[./-]\d{1,2}", _clean_cell_text(w["text"]))
    ]
    if not date_words:
        return row_ys

    refined: List[float] = [row_ys[0]]
    for ri in range(len(row_ys) - 1):
        y0, y1 = row_ys[ri], row_ys[ri + 1]
        in_band = sorted(
            [w for w in date_words if (y0 - 1.5) <= ((float(w["top"]) + float(w["bottom"])) / 2.0) <= (y1 + 1.5)],
            key=lambda w: float(w["top"]),
        )
        # Y좌표가 5.5pt 이상 떨어진 날짜가 2개 이상이면 각각의 거래 행으로 분리 (촘촘한 다열 표 대응)
        distinct_dates: List[Dict[str, Any]] = []
        for dw in in_band:
            if not distinct_dates or (float(dw["top"]) - float(distinct_dates[-1]["top"])) >= 5.5:
                distinct_dates.append(dw)

        if len(distinct_dates) >= 2:
            for di in range(len(distinct_dates) - 1):
                # 두 날짜의 중간 지점을 새 행 경계로 지정
                mid_y = (float(distinct_dates[di]["bottom"]) + float(distinct_dates[di + 1]["top"])) / 2.0
                if refined[-1] + 3.0 < mid_y < y1 - 3.0:
                    refined.append(mid_y)
        refined.append(y1)

    return refined


def _extract_tables_with_pdfplumber(
    page: Any,
    page_num: int,
    options: ConversionOptions,
    prev_col_xs: Optional[List[float]] = None,
    prev_date_col_idx: Optional[int] = None,
) -> Tuple[List[ExtractedTable], Optional[List[float]], Optional[int]]:
    """실제 거래내역 헤더 기준으로 상단 요약 박스와 하단 안내문을 제외하고 메인 표만 정밀 복원합니다."""
    tables_out: List[ExtractedTable] = []

    h_lines: List[Tuple[float, float, float]] = []
    v_lines: List[Tuple[float, float, float]] = []

    for ln in getattr(page, "lines", []) or []:
        x0, x1 = float(ln["x0"]), float(ln["x1"])
        top, bottom = float(ln["top"]), float(ln["bottom"])
        if abs(bottom - top) <= 2.5 and abs(x1 - x0) >= 20.0:
            h_lines.append((min(x0, x1), max(x0, x1), (top + bottom) / 2.0))
        elif abs(x1 - x0) <= 2.5 and abs(bottom - top) >= 8.0:
            v_lines.append(((x0 + x1) / 2.0, min(top, bottom), max(top, bottom)))

    for rect in getattr(page, "rects", []) or []:
        x0, x1 = float(rect["x0"]), float(rect["x1"])
        top, bottom = float(rect["top"]), float(rect["bottom"])
        w, h = abs(x1 - x0), abs(bottom - top)
        if w >= 20.0 and h <= 3.0:
            h_lines.append((min(x0, x1), max(x0, x1), (top + bottom) / 2.0))
        elif h >= 8.0 and w <= 3.0:
            v_lines.append(((x0 + x1) / 2.0, min(top, bottom), max(top, bottom)))
        elif w >= 40.0 and h >= 10.0:
            h_lines.append((min(x0, x1), max(x0, x1), min(top, bottom)))
            h_lines.append((min(x0, x1), max(x0, x1), max(top, bottom)))
            v_lines.append((min(x0, x1), min(top, bottom), max(top, bottom)))
            v_lines.append((max(x0, x1), min(top, bottom), max(top, bottom)))

    words = page.extract_words(
        x_tolerance=2.0,
        y_tolerance=2.5,
        keep_blank_chars=False,
        use_text_flow=False,
    ) or []

    # 1) 페이지 내에서 실제 거래내역 헤더 줄 탐색 (세로선 반영, 상단 계좌 요약 박스 자동 배제)
    header_info = _find_header_line_and_tokens(words, v_lines=v_lines)

    # 하단 안내문·회사정보 시작 Y좌표 탐색
    footer_limit_y = float(page.height)
    for w in words:
        cp = _compact_alnum(w["text"])
        if any(fk in cp for fk in ("고객님이", "본확인증", "본증명서", "법적효력", "인터넷뱅킹을", "회사명")):
            footer_limit_y = min(footer_limit_y, float(w["top"]) - 2.0)

    if header_info is not None:
        hdr_top, hdr_bot, hdr_tokens = header_info
        # 헤더 상단 경계선(헤더 바로 위 가로선) 탐색
        above_hs = [y for _, _, y in h_lines if (hdr_top - 16.0) <= y <= (hdr_top + 2.0)]
        ty0 = max(above_hs) if above_hs else (hdr_top - 4.0)

        # 헤더 이후 ~ 하단 안내문 이전의 가로선만 사용
        table_h_lines = [
            (x0, x1, y)
            for x0, x1, y in h_lines
            if (ty0 - 2.0) <= y <= footer_limit_y
        ]
        if table_h_lines:
            tx0 = min(min(x0 for x0, _, _ in table_h_lines), hdr_tokens[0][0] - 4.0)
            tx1 = max(max(x1 for _, x1, _ in table_h_lines), hdr_tokens[-1][1] + 4.0)
            ty1 = max(y for _, _, y in table_h_lines)
        else:
            tx0 = min(hdr_tokens[0][0] - 6.0, float(page.width) * 0.04)
            tx1 = max(hdr_tokens[-1][1] + 6.0, float(page.width) * 0.96)
            data_ys = [float(w["bottom"]) for w in words if hdr_bot < float(w["top"]) < footer_limit_y]
            ty1 = max(data_ys) + 4.0 if data_ys else footer_limit_y

        # 헤더 영역 아래의 세로선만 추출 (상단 요약 박스의 세로선 완전 차단)
        inner_vs = _cluster_coords(
            [
                vx
                for vx, vy0, vy1 in v_lines
                if (tx0 - 2.0) <= vx <= (tx1 + 2.0)
                and vy1 >= (hdr_bot - 2.0)
                and vy0 <= (ty1 + 2.0)
                and (min(vy1, ty1) - max(vy0, ty0)) >= 6.0
            ],
            tolerance=3.0,
        )

        # 헤더 토큰과 세로선을 결합해 정확한 열 경계(출금/입금 및 다열 구조 유지) 생성
        col_xs = _build_col_bounds_from_header_and_vlines(tx0, tx1, hdr_tokens, inner_vs)
        if prev_col_xs and len(prev_col_xs) == len(col_xs):
            if all(abs(a - b) <= 18.0 for a, b in zip(prev_col_xs, col_xs)):
                col_xs = prev_col_xs

        # 행 경계(가로선) 산출 및 헤더 아래 경계 보강
        raw_row_ys = [ty0, ty1] + [y for _, _, y in table_h_lines]
        below_hdr_hs = [y for _, _, y in table_h_lines if (hdr_bot - 1.0) <= y <= (hdr_bot + 14.0)]
        if not below_hdr_hs:
            raw_row_ys.append(hdr_bot + 3.0)
        row_ys = _cluster_coords(raw_row_ys, tolerance=3.0)

        # 날짜 열 인덱스 파악 후 가로선 누락 병합 행 분할
        date_cols = _get_date_col_indices([t[2] for t in hdr_tokens])
        if date_cols and date_cols[0] + 1 < len(col_xs):
            date_ci = date_cols[0]
        else:
            date_ci = _detect_date_col_index(words, col_xs, ty0, ty1)

        if date_ci + 1 < len(col_xs):
            row_ys = _split_row_ys_by_dates(
                row_ys,
                words,
                col_xs[date_ci],
                col_xs[date_ci + 1],
            )

        if len(row_ys) >= 2 and len(col_xs) >= 5:
            n_rows = len(row_ys) - 1
            n_cols = len(col_xs) - 1
            grid: List[List[List[Tuple[float, float, str]]]] = [
                [[] for _ in range(n_cols)] for _ in range(n_rows)
            ]

            for w in words:
                wx0, wx1 = float(w["x0"]), float(w["x1"])
                wtop, wbot = float(w["top"]), float(w["bottom"])
                cx = (wx0 + wx1) / 2.0
                cy = (wtop + wbot) / 2.0
                if not (col_xs[0] - 4.0 <= cx <= col_xs[-1] + 4.0):
                    continue
                if not (ty0 - 2.0 <= cy <= ty1 + 2.0):
                    continue

                r_idx = -1
                for ri in range(n_rows):
                    if row_ys[ri] - 1.5 <= cy <= row_ys[ri + 1] + 1.5:
                        r_idx = ri
                        break
                c_idx = -1
                for ci in range(n_cols):
                    if col_xs[ci] - 1.5 <= cx <= col_xs[ci + 1] + 1.5:
                        c_idx = ci
                        break
                if r_idx >= 0 and c_idx >= 0:
                    grid[r_idx][c_idx].append((round(wtop, 1), wx0, w["text"]))

            parsed_rows: List[List[str]] = []
            for ri in range(n_rows):
                row_cells: List[str] = []
                for ci in range(n_cols):
                    items = sorted(grid[ri][ci], key=lambda t: (t[0], t[1]))
                    row_cells.append(_clean_cell_text(" ".join(t[2] for t in items)))
                parsed_rows.append(row_cells)

            preview_rows: List[Tuple[List[str], bool]] = []
            cleaned = _filter_to_core_table_rows(parsed_rows, preserve_columns=True, deleted_out=preview_rows)
            if cleaned and len(cleaned) >= 1 and len(cleaned[0]) >= 2:
                return [
                    ExtractedTable(
                        page_num=page_num,
                        table_index=1,
                        rows=cleaned,
                        preview_rows=preview_rows,
                        source="TEXT-GRID",
                    )
                ], col_xs, date_ci

    # 2) 헤더가 반복되지 않는 이어지는 페이지: 이전 페이지의 열 경계(prev_col_xs)를 최우선 상속
    if len(h_lines) >= 2:
        all_row_ys = _cluster_coords(
            [y for _, _, y in h_lines if y <= footer_limit_y],
            tolerance=3.5,
        )
        y_blocks: List[List[float]] = []
        for y in all_row_ys:
            if not y_blocks:
                y_blocks.append([y])
            else:
                prev_y = y_blocks[-1][-1]
                bridged_by_vline = any(
                    vy0 <= prev_y + 3.0 and vy1 >= y - 3.0
                    for _, vy0, vy1 in v_lines
                )
                if (y - prev_y) <= 28.0 or bridged_by_vline:
                    y_blocks[-1].append(y)
                else:
                    y_blocks.append([y])

        best_col_xs: Optional[List[float]] = prev_col_xs
        best_date_ci: Optional[int] = prev_date_col_idx
        for b_idx, row_ys in enumerate(y_blocks, start=1):
            if len(row_ys) < 2:
                continue
            ty0, ty1 = row_ys[0], row_ys[-1]
            if (ty1 - ty0) < 15.0:
                continue

            block_h_lines = [
                (x0, x1, y)
                for x0, x1, y in h_lines
                if (ty0 - 3.0) <= y <= (ty1 + 3.0)
            ]
            if not block_h_lines:
                continue

            tx0 = min(x0 for x0, _, _ in block_h_lines)
            tx1 = max(x1 for _, x1, _ in block_h_lines)
            if (tx1 - tx0) < float(page.width) * 0.45:
                continue

            # 이전 페이지와 표 좌우 폭이 일치하면 1페이지의 열 경계를 그대로 상속해 위·아래 열 완벽 정렬
            if (
                prev_col_xs
                and len(prev_col_xs) >= 5
                and abs(tx0 - prev_col_xs[0]) <= 25.0
                and abs(tx1 - prev_col_xs[-1]) <= 25.0
            ):
                col_xs = prev_col_xs
            else:
                inner_vs = [
                    vx
                    for vx, vy0, vy1 in v_lines
                    if (tx0 - 2.0) <= vx <= (tx1 + 2.0)
                    and min(vy1, ty1) - max(vy0, ty0) >= 4.0
                ]
                col_xs = _cluster_coords([tx0, tx1] + inner_vs, tolerance=3.0)

            if len(col_xs) < 4:
                continue

            # 이전 페이지의 날짜 열 인덱스(또는 자동 탐지)로 병합 행 분할
            date_ci = (
                prev_date_col_idx
                if (prev_date_col_idx is not None and prev_date_col_idx + 1 < len(col_xs))
                else _detect_date_col_index(words, col_xs, ty0, ty1)
            )
            if date_ci + 1 < len(col_xs):
                row_ys = _split_row_ys_by_dates(
                    row_ys,
                    words,
                    col_xs[date_ci],
                    col_xs[date_ci + 1],
                )

            n_rows = len(row_ys) - 1
            n_cols = len(col_xs) - 1
            grid = [[[] for _ in range(n_cols)] for _ in range(n_rows)]

            for w in words:
                wx0, wx1 = float(w["x0"]), float(w["x1"])
                wtop, wbot = float(w["top"]), float(w["bottom"])
                cx = (wx0 + wx1) / 2.0
                cy = (wtop + wbot) / 2.0
                if not (col_xs[0] - 3.0 <= cx <= col_xs[-1] + 3.0):
                    continue
                if not (ty0 - 2.0 <= cy <= ty1 + 2.0):
                    continue

                r_idx = -1
                for ri in range(n_rows):
                    if row_ys[ri] - 1.5 <= cy <= row_ys[ri + 1] + 1.5:
                        r_idx = ri
                        break
                c_idx = -1
                for ci in range(n_cols):
                    if col_xs[ci] - 1.5 <= cx <= col_xs[ci + 1] + 1.5:
                        c_idx = ci
                        break

                if r_idx >= 0 and c_idx >= 0:
                    grid[r_idx][c_idx].append((round(wtop, 1), wx0, w["text"]))

            parsed_rows = []
            for ri in range(n_rows):
                row_cells = []
                for ci in range(n_cols):
                    items = sorted(grid[ri][ci], key=lambda t: (t[0], t[1]))
                    row_cells.append(_clean_cell_text(" ".join(t[2] for t in items)))
                parsed_rows.append(row_cells)

            cleaned = _filter_to_core_table_rows(
                parsed_rows,
                preserve_columns=bool(prev_col_xs),
            )
            if cleaned and len(cleaned) >= 1 and len(cleaned[0]) >= 2:
                tables_out.append(
                    ExtractedTable(
                        page_num=page_num,
                        table_index=b_idx,
                        rows=cleaned,
                        source="TEXT-GRID",
                    )
                )
                best_col_xs = col_xs
                best_date_ci = date_ci

        selected = _select_main_data_tables(tables_out)
        if selected:
            return selected, best_col_xs, best_date_ci

    # 3) 폴백: pdfplumber 기본 extract_tables
    raw_tables = page.extract_tables(
        table_settings={
            "vertical_strategy": "lines",
            "horizontal_strategy": "lines",
            "snap_tolerance": 4,
            "join_tolerance": 4,
        }
    ) or []
    if not raw_tables:
        raw_tables = page.extract_tables(
            table_settings={
                "vertical_strategy": "text",
                "horizontal_strategy": "lines",
                "snap_tolerance": 4,
            }
        ) or []

    for t_idx, raw_tbl in enumerate(raw_tables, start=1):
        rows = [[_clean_cell_text(c) for c in (r or [])] for r in (raw_tbl or [])]
        cleaned = _filter_to_core_table_rows(rows, preserve_columns=True)
        if cleaned and len(cleaned) >= 2 and len(cleaned[0]) >= 2:
            tables_out.append(
                ExtractedTable(
                    page_num=page_num,
                    table_index=t_idx,
                    rows=cleaned,
                    source="TEXT",
                )
            )

    return _select_main_data_tables(tables_out), prev_col_xs, prev_date_col_idx


def _run_rapidocr_on_image(
    ocr_engine: Any,
    img_bgr: Any,
) -> List[Tuple[List[Tuple[float, float]], str]]:
    """RapidOCR 신형/구형 출력 포맷을 통합 파싱해 [(다각형 좌표, 텍스트), ...] 목록을 반환합니다."""
    if ocr_engine is None:
        return []
    try:
        raw_out = ocr_engine(img_bgr)
    except Exception:
        return []

    results: List[Tuple[List[Tuple[float, float]], str]] = []

    # 신형 rapidocr (RapidOCROutput 객체: boxes, txts, scores 속성 보유)
    boxes = getattr(raw_out, "boxes", None)
    txts = getattr(raw_out, "txts", None)
    if boxes is not None and txts is not None:
        for box, txt in zip(boxes, txts):
            clean_t = _clean_cell_text(txt)
            if not clean_t:
                continue
            pts = [(float(p[0]), float(p[1])) for p in box]
            results.append((pts, clean_t))
        return results

    # 구형 rapidocr_onnxruntime: (result_list, elapse) 형태
    seq = raw_out[0] if isinstance(raw_out, tuple) and len(raw_out) > 0 else raw_out
    if isinstance(seq, list):
        for item in seq:
            if not isinstance(item, (list, tuple)) or len(item) < 2:
                continue
            box, txt = item[0], item[1]
            clean_t = _clean_cell_text(txt)
            if not clean_t:
                continue
            pts = [(float(p[0]), float(p[1])) for p in box]
            results.append((pts, clean_t))

    return results


def _extract_tables_with_ocr(
    pdfium_doc: Any,
    page_index_zero: int,
    page_num: int,
    options: ConversionOptions,
    ocr_engine: Any,
    prev_ocr_col_ratios: Optional[List[float]] = None,
) -> Tuple[List[ExtractedTable], Optional[List[float]]]:
    """스캔/이미지 PDF 페이지를 고해상도로 렌더링하고 OpenCV 격자 탐지 + RapidOCR로 표를 복원합니다."""
    import cv2
    import numpy as np

    page = pdfium_doc[page_index_zero]
    bitmap = page.render(scale=options.dpi_scale)
    pil_img = bitmap.to_pil().convert("RGB")
    img_bgr = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
    h, w = img_bgr.shape[:2]

    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    bw = cv2.adaptiveThreshold(
        gray,
        255,
        cv2.ADAPTIVE_THRESH_MEAN_C,
        cv2.THRESH_BINARY_INV,
        15,
        10,
    )

    h_kernel_len = max(25, w // 25)
    v_kernel_len = max(25, h // 30)
    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (h_kernel_len, 1))
    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, v_kernel_len))

    h_mask = cv2.dilate(cv2.erode(bw, h_kernel, iterations=1), h_kernel, iterations=1)
    v_mask = cv2.dilate(cv2.erode(bw, v_kernel, iterations=1), v_kernel, iterations=1)
    grid_mask = cv2.bitwise_or(h_mask, v_mask)

    contours, _ = cv2.findContours(grid_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    table_rects: List[Tuple[int, int, int, int]] = []
    for cnt in contours:
        rx, ry, rw, rh = cv2.boundingRect(cnt)
        if rw >= int(w * 0.45) and rh >= 60:
            table_rects.append((rx, ry, rw, rh))

    # 격자 윤곽이 안 잡히면 페이지 전체를 대상으로 OCR 수행
    if not table_rects:
        table_rects = [(0, 0, w, h)]
    else:
        table_rects.sort(key=lambda r: r[1])

    tables_out: List[ExtractedTable] = []

    for t_idx, (rx, ry, rw, rh) in enumerate(table_rects, start=1):
        roi_bgr = img_bgr[ry : ry + rh, rx : rx + rw]
        roi_h = h_mask[ry : ry + rh, rx : rx + rw]
        roi_v = v_mask[ry : ry + rh, rx : rx + rw]

        # 가로선/세로선 좌표 투영으로 행·열 경계 산출
        h_proj = np.sum(roi_h > 0, axis=1)
        v_proj = np.sum(roi_v > 0, axis=0)

        row_ys_raw = [float(y) for y in np.where(h_proj >= rw * 0.35)[0]]
        col_xs_raw = [float(x) for x in np.where(v_proj >= rh * 0.25)[0]]

        row_bounds = _cluster_coords(row_ys_raw, tolerance=8.0)
        col_bounds = _cluster_coords([0.0, float(rw - 1)] + col_xs_raw, tolerance=10.0)

        if len(col_bounds) < 4 and prev_ocr_col_ratios and len(prev_ocr_col_ratios) >= 4:
            col_bounds = [ratio * float(rw) for ratio in prev_ocr_col_ratios]

        ocr_items = _run_rapidocr_on_image(ocr_engine, roi_bgr)
        if not ocr_items:
            continue

        # 행 경계가 충분하면 격자 기반 매핑, 아니면 OCR 박스 Y좌표 클러스터링
        rows_matrix: List[List[str]] = []
        if len(row_bounds) >= 3 and len(col_bounds) >= 4:
            n_rows = len(row_bounds) - 1
            n_cols = len(col_bounds) - 1
            grid: List[List[List[Tuple[float, float, str]]]] = [
                [[] for _ in range(n_cols)] for _ in range(n_rows)
            ]

            for pts, txt in ocr_items:
                xs = [p[0] for p in pts]
                ys = [p[1] for p in pts]
                cx = sum(xs) / len(xs)
                cy = sum(ys) / len(ys)

                r_idx = -1
                for ri in range(n_rows):
                    if row_bounds[ri] - 4.0 <= cy <= row_bounds[ri + 1] + 4.0:
                        r_idx = ri
                        break
                c_idx = -1
                for ci in range(n_cols):
                    if col_bounds[ci] - 4.0 <= cx <= col_bounds[ci + 1] + 4.0:
                        c_idx = ci
                        break
                if r_idx >= 0 and c_idx >= 0:
                    grid[r_idx][c_idx].append((round(min(ys), 1), min(xs), txt))

            for ri in range(n_rows):
                row_cells: List[str] = []
                for ci in range(n_cols):
                    items = sorted(grid[ri][ci], key=lambda t: (t[0], t[1]))
                    row_cells.append(_clean_cell_text(" ".join(t[2] for t in items)))
                rows_matrix.append(row_cells)
        else:
            # 선이 부족한 스캔 문서는 텍스트 Y좌표로 줄을 묶고 X좌표로 정렬
            items_with_center: List[Tuple[float, float, str]] = []
            for pts, txt in ocr_items:
                xs = [p[0] for p in pts]
                ys = [p[1] for p in pts]
                items_with_center.append((sum(ys) / len(ys), min(xs), txt))
            items_with_center.sort(key=lambda t: (t[0], t[1]))

            grouped_lines: List[List[Tuple[float, float, str]]] = []
            for item in items_with_center:
                if not grouped_lines or abs(item[0] - grouped_lines[-1][0][0]) > 14.0:
                    grouped_lines.append([item])
                else:
                    grouped_lines[-1].append(item)

            for line in grouped_lines:
                line.sort(key=lambda t: t[1])
                rows_matrix.append([_clean_cell_text(t[2]) for t in line])

        cleaned = _filter_to_core_table_rows(rows_matrix)
        if cleaned and len(cleaned) >= 2 and len(cleaned[0]) >= 4:
            if len(col_bounds) >= 4:
                prev_ocr_col_ratios = [x / max(float(rw), 1.0) for x in col_bounds]
            tables_out.append(
                ExtractedTable(
                    page_num=page_num,
                    table_index=t_idx,
                    rows=cleaned,
                    source="OCR",
                )
            )

    return _select_main_data_tables(tables_out), prev_ocr_col_ratios


# =============================================================================
# 6. 엑셀(.xlsx) 내보내기
# =============================================================================
def save_result_to_excel(
    result: ConversionResult,
    save_path: str,
    options: ConversionOptions,
    include_deleted: bool = False,
) -> None:
    """추출 결과를 서식이 정리된 통합 거래내역 엑셀(.xlsx) 파일로 저장합니다."""
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    default_ws = wb.active

    header_fill = PatternFill(start_color="EEF2FF", end_color="EEF2FF", fill_type="solid")
    header_font = Font(name="맑은 고딕", size=10, bold=True, color="1E293B")
    body_font = Font(name="맑은 고딕", size=10, color="1E293B")
    deleted_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    thin_side = Side(style="thin", color="CBD5E1")
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    def _write_sheet(
        ws: Any,
        rows: List[List[str]],
        deleted_flags: Optional[List[bool]] = None,
    ) -> None:
        if not rows:
            return
        max_cols = max(len(r) for r in rows)
        col_widths = [8] * max_cols
        has_header = _is_likely_table_header_row(rows[0])
        date_cols = set(_get_date_col_indices(rows[0])) if has_header else set()

        for r_idx, row in enumerate(rows, start=1):
            is_deleted = bool(deleted_flags and (r_idx - 1) < len(deleted_flags) and deleted_flags[r_idx - 1])
            is_hdr = (r_idx == 1 and has_header and not is_deleted)
            for c_idx in range(1, max_cols + 1):
                raw_txt = row[c_idx - 1] if (c_idx - 1) < len(row) else ""
                val, num_fmt = _parse_excel_cell_value(
                    raw_txt,
                    is_hdr,
                    options,
                    is_date_col=((c_idx - 1) in date_cols),
                )
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = cell_border

                if is_hdr:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                else:
                    cell.font = body_font
                    if is_deleted:
                        cell.fill = deleted_fill
                    if num_fmt == "yyyy-mm-dd":
                        cell.number_format = num_fmt
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif num_fmt:
                        cell.number_format = num_fmt
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

                display_txt = val.strftime("%Y-%m-%d") if isinstance(val, date) else str(raw_txt)
                display_len = sum(2 if ord(ch) > 127 else 1 for ch in display_txt)
                col_widths[c_idx - 1] = max(col_widths[c_idx - 1], min(display_len + 4, 48))

        for c_idx, width in enumerate(col_widths, start=1):
            ws.column_dimensions[get_column_letter(c_idx)].width = width

        if has_header:
            ws.freeze_panes = "A2"

    if include_deleted and result.preview_rows:
        default_ws.title = "거래내역_통합_가공전"
        _write_sheet(
            default_ws,
            [r for r, _ in result.preview_rows],
            deleted_flags=[d for _, d in result.preview_rows],
        )
    elif result.merged_rows:
        default_ws.title = "거래내역_통합"
        _write_sheet(default_ws, result.merged_rows)
    else:
        default_ws.title = "결과없음"

    wb.save(save_path)


# =============================================================================
# 7. 백그라운드 작업 스레드 (의존성 준비 / PDF 변환)
# =============================================================================
class BootstrapWorker(QThread):
    """앱 시작 시 필수 라이브러리를 백그라운드에서 점검합니다."""

    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, list)

    def __init__(self, dep_manager: DependencyManager) -> None:
        super().__init__()
        self.dep_manager = dep_manager

    def run(self) -> None:
        ok, warnings = self.dep_manager.ensure_core_packages(self.log_signal.emit)
        self.finished_signal.emit(ok, warnings)


class ConvertWorker(QThread):
    """PDF 파일에서 표 데이터를 추출하는 백그라운드 워커."""

    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int, int)
    password_requested = pyqtSignal(str, bool)  # (pdf_path, is_retry)
    finished_signal = pyqtSignal(object)  # ConversionResult
    error_signal = pyqtSignal(str)

    def __init__(
        self,
        pdf_path: str,
        options: ConversionOptions,
        dep_manager: DependencyManager,
        initial_password: Optional[str] = None,
    ) -> None:
        super().__init__()
        self.pdf_path = pdf_path
        self.options = options
        self.dep_manager = dep_manager
        self.password: Optional[str] = initial_password
        stem_parts = Path(pdf_path).stem.split()
        self._filename_pw: Optional[str] = stem_parts[-1] if len(stem_parts) >= 2 else None
        self._tried_filename_pw: bool = False
        self._pw_event = threading.Event()
        self._pw_cancelled = False

    def provide_password(self, pw: Optional[str]) -> None:
        """UI 스레드에서 입력받은 비밀번호를 전달받습니다."""
        if pw is None:
            self._pw_cancelled = True
        else:
            self.password = pw
            self._pw_cancelled = False
        self._pw_event.set()

    def _ask_password_sync(self, is_retry: bool = False) -> Optional[str]:
        """UI 스레드에 비밀번호 입력을 요청하고 응답을 대기합니다."""
        self._pw_event.clear()
        self.password_requested.emit(self.pdf_path, is_retry)
        self._pw_event.wait()
        if self._pw_cancelled:
            return None
        return self.password

    def _open_pdfplumber(self) -> Any:
        import pdfplumber

        attempt = 0
        while True:
            try:
                return pdfplumber.open(self.pdf_path, password=self.password or "")
            except Exception as exc:
                if _is_password_error(exc):
                    if not self.password and self._filename_pw and not self._tried_filename_pw:
                        self._tried_filename_pw = True
                        self.password = self._filename_pw
                        continue
                    self.password = None
                    pw = self._ask_password_sync(is_retry=(attempt > 0 or bool(self.password)))
                    if pw is None:
                        raise RuntimeError("비밀번호 입력이 취소되었습니다.") from exc
                    attempt += 1
                    continue
                raise

    def _open_pdfium(self) -> Any:
        import pypdfium2 as pdfium

        attempt = 0
        while True:
            try:
                if self.password:
                    return pdfium.PdfDocument(self.pdf_path, password=self.password)
                return pdfium.PdfDocument(self.pdf_path)
            except Exception as exc:
                if _is_password_error(exc):
                    if not self.password and self._filename_pw and not self._tried_filename_pw:
                        self._tried_filename_pw = True
                        self.password = self._filename_pw
                        continue
                    self.password = None
                    pw = self._ask_password_sync(is_retry=(attempt > 0 or bool(self.password)))
                    if pw is None:
                        raise RuntimeError("비밀번호 입력이 취소되었습니다.") from exc
                    attempt += 1
                    continue
                raise

    def run(self) -> None:
        result = ConversionResult(pdf_path=self.pdf_path)
        pdf_doc = None
        pdfium_doc = None
        try:
            self.log_signal.emit(f"📄 파일 열기: {Path(self.pdf_path).name}")
            pdf_doc = self._open_pdfplumber()
            total_pages = len(pdf_doc.pages)
            result.page_count = total_pages
            self.log_signal.emit(f"📑 총 {total_pages}페이지 분석 시작...")

            prev_col_xs: Optional[List[float]] = None
            prev_date_col_idx: Optional[int] = None
            prev_ocr_col_ratios: Optional[List[float]] = None

            for idx, page in enumerate(pdf_doc.pages):
                page_num = idx + 1
                self.progress_signal.emit(page_num, total_pages)

                page_tables: List[ExtractedTable] = []
                if not self.options.force_ocr:
                    page_tables, prev_col_xs, prev_date_col_idx = _extract_tables_with_pdfplumber(
                        page=page,
                        page_num=page_num,
                        options=self.options,
                        prev_col_xs=prev_col_xs,
                        prev_date_col_idx=prev_date_col_idx,
                    )

                if page_tables:
                    self.log_signal.emit(
                        f"  · {page_num}페이지: 텍스트 표 추출 완료 "
                        f"({len(page_tables[0].rows)}행 × {len(page_tables[0].rows[0])}열)"
                    )
                    result.tables.extend(page_tables)
                else:
                    # 텍스트 표가 없으면 OCR 모드로 전환
                    self.log_signal.emit(f"  · {page_num}페이지: 이미지/스캔 표 OCR 시도 중...")
                    ocr_engine = self.dep_manager.get_ocr_engine(self.log_signal.emit)
                    if ocr_engine is not None:
                        if pdfium_doc is None:
                            pdfium_doc = self._open_pdfium()
                        ocr_tables, prev_ocr_col_ratios = _extract_tables_with_ocr(
                            pdfium_doc=pdfium_doc,
                            page_index_zero=idx,
                            page_num=page_num,
                            options=self.options,
                            ocr_engine=ocr_engine,
                            prev_ocr_col_ratios=prev_ocr_col_ratios,
                        )
                        if ocr_tables:
                            result.used_ocr_pages.append(page_num)
                            result.tables.extend(ocr_tables)
                            self.log_signal.emit(
                                f"  · {page_num}페이지: OCR 표 추출 완료 "
                                f"({len(ocr_tables[0].rows)}행 × {len(ocr_tables[0].rows[0])}열)"
                            )
                        else:
                            self.log_signal.emit(f"  · {page_num}페이지: 추출 가능한 거래내역 표 없음")
                    else:
                        msg = f"{page_num}페이지: OCR 엔진 미설치로 스캔 이미지 인식 건너뜀"
                        result.warnings.append(msg)
                        self.log_signal.emit(f"  ⚠ {msg}")

            result.merged_rows = _merge_tables_into_single_grid(result.tables, self.options)
            all_preview: List[Tuple[List[str], bool]] = []
            seen_header = False
            for tbl in result.tables:
                src_rows = tbl.preview_rows if tbl.preview_rows else [(r, False) for r in tbl.rows]
                for r, is_del in src_rows:
                    if not is_del and _is_likely_table_header_row(r):
                        if seen_header:
                            all_preview.append((r, True))
                            continue
                        seen_header = True
                    all_preview.append((r, is_del))
            result.preview_rows = all_preview
            self.log_signal.emit(
                f"✅ 변환 완료: 총 {len(result.tables)}개 페이지 표 / "
                f"통합 {len(result.merged_rows)}행"
            )
            self.finished_signal.emit(result)
        except Exception as exc:
            self.log_signal.emit(f"❌ 오류 발생: {exc}")
            self.error_signal.emit(str(exc))
        finally:
            try:
                if pdf_doc is not None:
                    pdf_doc.close()
            except Exception:
                pass
            try:
                if pdfium_doc is not None:
                    pdfium_doc.close()
            except Exception:
                pass


# =============================================================================
# 8. 메인 GUI 윈도우 (라이트 테마 고정, 가로 열 헤더 숨김, 셀 선택 비활성화)
# =============================================================================
class DropZoneFrame(QFrame):
    """PDF 드래그 앤 드롭을 지원하는 상단 파일 영역."""

    file_dropped = pyqtSignal(str)

    def __init__(self) -> None:
        super().__init__()
        self.setAcceptDrops(True)
        self.setObjectName("DropZone")

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().lower().endswith(".pdf"):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:
        for url in event.mimeData().urls():
            local_path = url.toLocalFile()
            if local_path.lower().endswith(".pdf"):
                self.file_dropped.emit(local_path)
                event.acceptProposedAction()
                return
        event.ignore()


class MainWindow(QMainWindow):
    """꽁순이 OCellR - NH Out! 메인 윈도우."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.setWindowIcon(_create_app_icon())
        self.resize(1180, 780)
        self.setMinimumSize(920, 620)

        self.dep_manager = DependencyManager()
        self.current_pdf_path: Optional[str] = None
        self.session_passwords: Dict[str, str] = {}
        self.current_result: Optional[ConversionResult] = None
        self.bootstrap_worker: Optional[BootstrapWorker] = None
        self.convert_worker: Optional[ConvertWorker] = None
        self.is_ready = False

        self._build_ui()
        self._apply_styles()
        self._start_bootstrap()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(18, 16, 18, 16)
        root_layout.setSpacing(12)

        # 상단 헤더 영역
        header_frame = QFrame()
        header_frame.setObjectName("HeaderBanner")
        _banner_img = Path(__file__).resolve().parent / "BG.webp"
        if _banner_img.exists():
            _img_url = _banner_img.as_posix()
            header_frame.setStyleSheet(
                f"#HeaderBanner {{ background-image: url({_img_url}); background-repeat: no-repeat; background-position: center; background-color: #FFFFFF; border-radius: 10px; }}"
                " #AppTitle, #AppSubtitle { background-color: rgba(255, 255, 255, 0.28); padding: 2px 8px; border-radius: 6px; }"
                " #AppSubtitle { color: #0F172A; }"
            )
        header_row = QHBoxLayout(header_frame)
        header_row.setContentsMargins(14, 10, 14, 10)
        title_box = QVBoxLayout()
        title_box.setSpacing(2)

        self.lbl_title = QLabel(
            "꽁순이 OCellR <span style='font-size:11px; font-weight:normal; color:#334155;'>by tokis.tistory.com</span>"
        )
        self.lbl_title.setObjectName("AppTitle")
        self.lbl_subtitle = QLabel(
            "PDF 파일에서 표 바깥쪽 불필요한 영역을 제외하고 핵심 거래내역만 엑셀로 변환합니다."
        )
        self.lbl_subtitle.setObjectName("AppSubtitle")
        title_box.addWidget(self.lbl_title, alignment=Qt.AlignmentFlag.AlignLeft)
        title_box.addWidget(self.lbl_subtitle, alignment=Qt.AlignmentFlag.AlignLeft)
        from PyQt6.QtWidgets import QGraphicsDropShadowEffect
        for _lbl in (self.lbl_title, self.lbl_subtitle):
            _shadow = QGraphicsDropShadowEffect(_lbl)
            _shadow.setBlurRadius(6)
            _shadow.setOffset(0, 1)
            _shadow.setColor(QColor(255, 255, 255, 220))
            _lbl.setGraphicsEffect(_shadow)
        header_row.addLayout(title_box, stretch=1)

        self.lbl_status_badge = QLabel("초기화 중...")
        self.lbl_status_badge.setObjectName("StatusBadge")
        header_row.addWidget(
            self.lbl_status_badge,
            alignment=Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
        )
        root_layout.addWidget(header_frame)

        # 파일 선택 및 드래그 앤 드롭 카드
        self.drop_zone = DropZoneFrame()
        root_layout.addWidget(self.drop_zone)
        self.drop_zone.file_dropped.connect(self._on_file_selected)
        dz_layout = QHBoxLayout(self.drop_zone)
        dz_layout.setContentsMargins(16, 12, 16, 12)
        dz_layout.setSpacing(12)

        self.lbl_file_info = QLabel("여기에 PDF 파일을 끌어다 놓거나 [PDF 선택] 버튼을 눌러주세요.")
        self.lbl_file_info.setObjectName("FileInfoLabel")
        dz_layout.addWidget(self.lbl_file_info, stretch=1)

        self.btn_browse = QPushButton("PDF 선택")
        self.btn_browse.setObjectName("SecondaryButton")
        self.btn_browse.clicked.connect(self._browse_pdf)
        dz_layout.addWidget(self.btn_browse)

        self.btn_save = QPushButton("엑셀(.xlsx) 저장")
        self.btn_save.setObjectName("AccentButton")
        self.btn_save.setEnabled(False)
        self.btn_save.clicked.connect(self._save_excel)
        dz_layout.addWidget(self.btn_save)
        root_layout.addWidget(self.drop_zone)

        # 옵션 바
        opt_frame = QFrame()
        opt_frame.setObjectName("OptionBar")
        opt_layout = QHBoxLayout(opt_frame)
        opt_layout.setContentsMargins(14, 8, 14, 8)
        opt_layout.setSpacing(18)

        self.chk_dates = QCheckBox("날짜 형식 통일 (YYYY-MM-DD)")
        self.chk_dates.setChecked(True)
        self.chk_numbers = QCheckBox("금액 숫자 자동 변환")
        self.chk_numbers.setChecked(True)
        self.chk_force_ocr = QCheckBox("강제 OCR 모드 (스캔 PDF용)")
        self.chk_force_ocr.setChecked(False)
        self.chk_show_deleted = QCheckBox("가공 전 데이타 보기")
        self.chk_show_deleted.setChecked(False)
        self.chk_show_deleted.stateChanged.connect(self._on_show_deleted_toggled)

        for chk in (self.chk_dates, self.chk_numbers, self.chk_force_ocr):
            chk.stateChanged.connect(self._on_option_changed)

        opt_layout.addWidget(self.chk_dates)
        opt_layout.addWidget(self.chk_numbers)
        opt_layout.addWidget(self.chk_force_ocr)
        opt_layout.addWidget(self.chk_show_deleted)
        opt_layout.addStretch(1)

        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedWidth(200)
        self.progress_bar.setFixedHeight(16)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        opt_layout.addWidget(self.progress_bar)
        root_layout.addWidget(opt_frame)

        # 메인 스플리터 (상단: 미리보기 탭, 하단: 로그)
        splitter = QSplitter(Qt.Orientation.Vertical)
        splitter.setHandleWidth(6)

        self.preview_tabs = QTabWidget()
        self.preview_tabs.setObjectName("PreviewTabs")
        self._show_empty_preview("PDF 파일을 선택하면 변환된 거래내역 미리보기가 여기에 표시됩니다.")
        splitter.addWidget(self.preview_tabs)

        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setObjectName("LogView")
        self.log_view.setMaximumHeight(170)
        splitter.addWidget(self.log_view)
        splitter.setStretchFactor(0, 5)
        splitter.setStretchFactor(1, 1)

        root_layout.addWidget(splitter, stretch=1)

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow, QWidget {
                background-color: #F8FAFC;
                color: #1E293B;
                font-family: "맑은 고딕", "Malgun Gothic", sans-serif;
                font-size: 13px;
            }
            QLabel#AppTitle {
                font-size: 20px;
                font-weight: 700;
                color: #0F172A;
            }
            QLabel#AppSubtitle {
                font-size: 12px;
                color: #475569;
            }
            QLabel#StatusBadge {
                background-color: #E2E8F0;
                color: #334155;
                border-radius: 10px;
                padding: 5px 12px;
                font-weight: 600;
                font-size: 12px;
            }
            QFrame#DropZone {
                background-color: #FFFFFF;
                border: 2px dashed #CBD5E1;
                border-radius: 12px;
            }
            QLabel#FileInfoLabel {
                color: #334155;
                font-weight: 500;
                background: transparent;
            }
            QFrame#OptionBar {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 10px;
            }
            QCheckBox {
                color: #1E293B;
                spacing: 6px;
                background: transparent;
            }
            QPushButton {
                border-radius: 8px;
                padding: 8px 16px;
                font-weight: 600;
            }
            QPushButton#PrimaryButton {
                background-color: #2563EB;
                color: #FFFFFF;
                border: none;
            }
            QPushButton#PrimaryButton:hover {
                background-color: #1D4ED8;
            }
            QPushButton#PrimaryButton:disabled {
                background-color: #94A3B8;
                color: #F8FAFC;
            }
            QPushButton#AccentButton {
                background-color: #059669;
                color: #FFFFFF;
                border: none;
            }
            QPushButton#AccentButton:hover {
                background-color: #047857;
            }
            QPushButton#AccentButton:disabled {
                background-color: #94A3B8;
                color: #F8FAFC;
            }
            QPushButton#SecondaryButton {
                background-color: #F1F5F9;
                color: #1E293B;
                border: 1px solid #CBD5E1;
            }
            QPushButton#SecondaryButton:hover {
                background-color: #E2E8F0;
            }
            QTabWidget#PreviewTabs::pane {
                border: 1px solid #CBD5E1;
                border-radius: 8px;
                background-color: #FFFFFF;
            }
            QTabBar::tab {
                background-color: #E2E8F0;
                color: #334155;
                padding: 7px 14px;
                margin-right: 4px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                font-weight: 600;
            }
            QTabBar::tab:selected {
                background-color: #FFFFFF;
                color: #0F172A;
                border: 1px solid #CBD5E1;
                border-bottom: none;
            }
            QTableWidget {
                background-color: #FFFFFF;
                alternate-background-color: #F8FAFC;
                color: #1E293B;
                gridline-color: #E2E8F0;
                border: none;
            }
            QTableWidget::item {
                color: #1E293B;
                padding: 4px 8px;
            }
            QHeaderView::section {
                background-color: #F1F5F9;
                color: #475569;
                border: 1px solid #E2E8F0;
                padding: 4px 6px;
                font-weight: 600;
            }
            QTextEdit#LogView {
                background-color: #FFFFFF;
                color: #334155;
                border: 1px solid #CBD5E1;
                border-radius: 8px;
                padding: 8px;
                font-family: "Consolas", "맑은 고딕", monospace;
                font-size: 12px;
            }
            QProgressBar {
                background-color: #E2E8F0;
                border: none;
                border-radius: 8px;
                text-align: center;
                color: #1E293B;
                font-size: 11px;
                font-weight: 600;
            }
            QProgressBar::chunk {
                background-color: #2563EB;
                border-radius: 8px;
            }
            /* 모던 슬림 스크롤바 (세로/가로 공통) */
            QScrollBar:vertical {
                background: transparent;
                width: 10px;
                margin: 2px 2px 2px 0px;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical {
                background-color: #CBD5E1;
                min-height: 36px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #94A3B8;
            }
            QScrollBar::handle:vertical:pressed {
                background-color: #64748B;
            }
            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical {
                height: 0px;
                width: 0px;
                background: none;
                border: none;
            }
            QScrollBar::add-page:vertical,
            QScrollBar::sub-page:vertical {
                background: transparent;
            }
            QScrollBar:horizontal {
                background: transparent;
                height: 10px;
                margin: 0px 2px 2px 2px;
                border-radius: 5px;
            }
            QScrollBar::handle:horizontal {
                background-color: #CBD5E1;
                min-width: 36px;
                border-radius: 4px;
            }
            QScrollBar::handle:horizontal:hover {
                background-color: #94A3B8;
            }
            QScrollBar::handle:horizontal:pressed {
                background-color: #64748B;
            }
            QScrollBar::add-line:horizontal,
            QScrollBar::sub-line:horizontal {
                width: 0px;
                height: 0px;
                background: none;
                border: none;
            }
            QScrollBar::add-page:horizontal,
            QScrollBar::sub-page:horizontal {
                background: transparent;
            }
            QAbstractScrollArea::corner {
                background: transparent;
            }
            """
        )

    def _append_log(self, message: str) -> None:
        self.log_view.append(message)

    def _start_bootstrap(self) -> None:
        self.bootstrap_worker = BootstrapWorker(self.dep_manager)
        self.bootstrap_worker.log_signal.connect(self._append_log)
        self.bootstrap_worker.finished_signal.connect(self._on_bootstrap_finished)
        self.bootstrap_worker.start()

    def _on_bootstrap_finished(self, ok: bool, warnings: List[str]) -> None:
        self.is_ready = ok
        if ok:
            if warnings:
                self.lbl_status_badge.setText("준비 완료 (일부 안내)")
                self._append_log(f"▶ 준비 완료 (주의: {warnings[0]})")
            else:
                self.lbl_status_badge.setText("준비 완료")
                self._append_log("▶ 준비 완료! PDF 파일을 드래그하거나 선택해 주세요.")
            if self.current_pdf_path:
                self._start_conversion()
        else:
            self.lbl_status_badge.setText("초기화 실패")
            QMessageBox.critical(
                self,
                APP_TITLE,
                "필수 라이브러리 설치에 실패했습니다.\n로그 내용을 확인해 주세요.",
            )

    def _browse_pdf(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "변환할 PDF 파일 선택",
            "",
            "PDF 파일 (*.pdf)",
        )
        if path:
            self._on_file_selected(path)

    def _on_file_selected(self, path: str) -> None:
        self.current_pdf_path = path
        self.lbl_file_info.setText(f"선택된 파일: {Path(path).name}  ({path})")
        self._append_log(f"📂 파일 선택됨: {path}")
        if self.is_ready:
            self._start_conversion()

    def _set_options_enabled(self, enabled: bool) -> None:
        self.chk_dates.setEnabled(enabled)
        self.chk_numbers.setEnabled(enabled)
        self.chk_force_ocr.setEnabled(enabled)

    def _on_option_changed(self) -> None:
        if self.current_pdf_path and self.is_ready:
            self._start_conversion()

    def _collect_options(self) -> ConversionOptions:
        return ConversionOptions(
            merge_into_single_sheet=True,
            force_ocr=self.chk_force_ocr.isChecked(),
            normalize_dates=self.chk_dates.isChecked(),
            parse_numbers=self.chk_numbers.isChecked(),
        )

    def _start_conversion(self) -> None:
        if not self.current_pdf_path or not self.is_ready:
            return
        if self.convert_worker and self.convert_worker.isRunning():
            return

        self.btn_browse.setEnabled(False)
        self.btn_save.setEnabled(False)
        self._set_options_enabled(False)
        self.progress_bar.setValue(0)
        self.lbl_status_badge.setText("변환 중...")

        saved_pw = self.session_passwords.get(self.current_pdf_path)
        self.convert_worker = ConvertWorker(
            pdf_path=self.current_pdf_path,
            options=self._collect_options(),
            dep_manager=self.dep_manager,
            initial_password=saved_pw,
        )
        self.convert_worker.log_signal.connect(self._append_log)
        self.convert_worker.progress_signal.connect(self._on_progress)
        self.convert_worker.password_requested.connect(self._on_password_requested)
        self.convert_worker.finished_signal.connect(self._on_conversion_finished)
        self.convert_worker.error_signal.connect(self._on_conversion_error)
        self.convert_worker.start()

    def _on_progress(self, current: int, total: int) -> None:
        if total > 0:
            self.progress_bar.setValue(int(current * 100 / total))

    def _on_password_requested(self, pdf_path: str, is_retry: bool) -> None:
        prompt = (
            "비밀번호가 올바르지 않습니다. 다시 입력해 주세요:"
            if is_retry
            else f"'{Path(pdf_path).name}' 파일은 비밀번호로 보호되어 있습니다.\n비밀번호를 입력해 주세요:"
        )
        pw, ok = QInputDialog.getText(
            self,
            f"{APP_TITLE} - PDF 비밀번호 입력",
            prompt,
            QLineEdit.EchoMode.Password,
        )
        if ok:
            self.session_passwords[pdf_path] = pw
            if self.convert_worker:
                self.convert_worker.provide_password(pw)
        else:
            if self.convert_worker:
                self.convert_worker.provide_password(None)

    def _on_conversion_finished(self, result: ConversionResult) -> None:
        self.current_result = result
        self.btn_browse.setEnabled(True)
        self._set_options_enabled(True)
        self.progress_bar.setValue(100)
        self.lbl_status_badge.setText("변환 완료")

        self._render_preview(result)
        has_data = bool(result.merged_rows or result.tables)
        self.btn_save.setEnabled(has_data)

        if not has_data:
            QMessageBox.information(
                self,
                APP_TITLE,
                "추출된 거래내역 표가 없습니다.\n스캔 PDF라면 '강제 OCR 모드'를 켜고 다시 시도해 보세요.",
            )

    def _on_conversion_error(self, err_msg: str) -> None:
        self.btn_browse.setEnabled(True)
        self._set_options_enabled(True)
        self.lbl_status_badge.setText("오류 발생")
        QMessageBox.warning(self, APP_TITLE, f"변환 중 오류가 발생했습니다:\n{err_msg}")

    def _show_empty_preview(self, text: str) -> None:
        self.preview_tabs.clear()
        placeholder = QLabel(text)
        placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        placeholder.setStyleSheet("color: #64748B; font-size: 14px; background-color: #FFFFFF;")
        self.preview_tabs.addTab(placeholder, "미리보기")

    def _create_preview_table(
        self,
        rows: List[List[str]],
        deleted_flags: Optional[List[bool]] = None,
    ) -> QTableWidget:
        """가로 열 번호 헤더를 숨기고 셀 선택을 비활성화한 깔끔한 미리보기 표를 생성합니다."""
        table = QTableWidget()
        n_rows = len(rows)
        n_cols = max((len(r) for r in rows), default=0)
        table.setRowCount(n_rows)
        table.setColumnCount(n_cols)
        table.setAlternatingRowColors(True)

        # 요청사항 반영: 상단 가로 헤더('열 1', '열 2'...) 숨김 & 셀 선택/편집 비활성화
        table.horizontalHeader().setVisible(False)
        table.verticalHeader().setVisible(True)
        table.verticalHeader().setDefaultSectionSize(28)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        table.setFocusPolicy(Qt.FocusPolicy.NoFocus)

        has_header = bool(rows and _is_likely_table_header_row(rows[0]))
        header_bg = QColor("#EEF2FF")
        text_color = QColor("#1E293B")
        deleted_bg = QColor("#F8FAFC")
        deleted_text_color = QColor("#CBD5E1")
        header_font = QFont("맑은 고딕", 10, QFont.Weight.Bold)

        for r_idx, row in enumerate(rows):
            is_deleted = bool(deleted_flags and r_idx < len(deleted_flags) and deleted_flags[r_idx])
            is_hdr = (r_idx == 0 and has_header and not is_deleted)
            for c_idx in range(n_cols):
                val = row[c_idx] if c_idx < len(row) else ""
                item = QTableWidgetItem(val)
                item.setForeground(text_color)
                if is_deleted:
                    item.setForeground(deleted_text_color)
                    item.setBackground(deleted_bg)
                if is_hdr:
                    item.setBackground(header_bg)
                    item.setFont(header_font)
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                else:
                    if _NUMBER_RE.match(val):
                        item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    else:
                        item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                table.setItem(r_idx, c_idx, item)

        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        return table

    def _on_show_deleted_toggled(self) -> None:
        if self.current_result is not None:
            self._render_preview(self.current_result)

    def _render_preview(self, result: ConversionResult) -> None:
        self.preview_tabs.clear()
        if not result.merged_rows and not (self.chk_show_deleted.isChecked() and result.preview_rows):
            self._show_empty_preview("표 안쪽에서 추출된 거래내역 데이터가 없습니다.")
            return

        if self.chk_show_deleted.isChecked() and result.preview_rows:
            rows = [r for r, _ in result.preview_rows]
            deleted_flags = [d for _, d in result.preview_rows]
            merged_table = self._create_preview_table(rows, deleted_flags=deleted_flags)
            del_cnt = sum(1 for d in deleted_flags if d)
            tab_title = f"통합 거래내역 ({len(result.merged_rows)}행 / 삭제 {del_cnt}행 포함)"
        else:
            merged_table = self._create_preview_table(result.merged_rows)
            tab_title = f"통합 거래내역 ({len(result.merged_rows)}행)"

        self.preview_tabs.addTab(merged_table, tab_title)

    def _save_excel(self) -> None:
        if not self.current_result or not self.current_pdf_path:
            return

        src = Path(self.current_pdf_path)
        suffix = "_변환결과_가공전" if self.chk_show_deleted.isChecked() else "_변환결과"
        default_out = str(src.with_name(f"{src.stem}{suffix}.xlsx"))
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "엑셀 파일로 저장",
            default_out,
            "Excel 통합 문서 (*.xlsx)",
        )
        if not save_path:
            return
        if not save_path.lower().endswith(".xlsx"):
            save_path += ".xlsx"

        try:
            save_result_to_excel(
                result=self.current_result,
                save_path=save_path,
                options=self._collect_options(),
                include_deleted=self.chk_show_deleted.isChecked(),
            )
            self._append_log(f"💾 엑셀 저장 완료: {save_path}")
            QMessageBox.information(self, APP_TITLE, f"엑셀 파일이 저장되었습니다.\n{save_path}")
        except Exception as exc:
            self._append_log(f"❌ 엑셀 저장 실패: {exc}")
            QMessageBox.critical(self, APP_TITLE, f"엑셀 저장 중 오류가 발생했습니다:\n{exc}")


# =============================================================================
# 9. 애플리케이션 엔트리포인트 (앱 아이콘 및 라이트 팔레트 고정)
# =============================================================================
def _create_app_icon() -> QIcon:
    """외부 이미지 파일 없이 프로그램 창·작업표시줄용 꽁순이 아이콘을 생성합니다."""
    size = 64
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.GlobalColor.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

    # 둥근 배경 카드 (에메랄드-블루 톤)
    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QColor("#0F766E"))
    painter.drawRoundedRect(4, 4, 56, 56, 14, 14)

    # 꽁순이 토끼 귀
    painter.setBrush(QColor("#FFFFFF"))
    painter.drawEllipse(16, 10, 11, 24)
    painter.drawEllipse(37, 10, 11, 24)

    # 귀 안쪽 핑크
    painter.setBrush(QColor("#F9A8D4"))
    painter.drawEllipse(19, 14, 5, 15)
    painter.drawEllipse(40, 14, 5, 15)

    # 토끼 얼굴
    painter.setBrush(QColor("#FFFFFF"))
    painter.drawEllipse(14, 24, 36, 30)

    # 눈
    painter.setBrush(QColor("#0F172A"))
    painter.drawEllipse(23, 35, 4, 5)
    painter.drawEllipse(37, 35, 4, 5)

    # 코
    painter.setBrush(QColor("#F472B6"))
    painter.drawEllipse(30, 41, 4, 3)

    # 우측 하단 작은 엑셀 표 배지
    painter.setBrush(QColor("#10B981"))
    painter.setPen(QPen(QColor("#FFFFFF"), 2))
    painter.drawRoundedRect(40, 40, 20, 20, 5, 5)

    painter.setPen(QPen(QColor("#FFFFFF"), 1.5))
    painter.drawLine(45, 47, 55, 47)
    painter.drawLine(45, 52, 55, 52)
    painter.drawLine(50, 44, 50, 56)

    painter.end()
    return QIcon(pixmap)


def _apply_light_palette(app: QApplication) -> None:
    """Windows 다크 모드에서도 글자가 흰색으로 가려지지 않도록 라이트 팔레트를 고정합니다."""
    app.setStyle("Fusion")
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor("#F8FAFC"))
    palette.setColor(QPalette.ColorRole.WindowText, QColor("#1E293B"))
    palette.setColor(QPalette.ColorRole.Base, QColor("#FFFFFF"))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor("#F8FAFC"))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor("#FFFFFF"))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor("#1E293B"))
    palette.setColor(QPalette.ColorRole.Text, QColor("#1E293B"))
    palette.setColor(QPalette.ColorRole.Button, QColor("#F1F5F9"))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor("#1E293B"))
    palette.setColor(QPalette.ColorRole.BrightText, QColor("#DC2626"))
    palette.setColor(QPalette.ColorRole.Highlight, QColor("#2563EB"))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#FFFFFF"))
    app.setPalette(palette)


def main() -> None:
    # Windows 작업 표시줄에서도 Python 기본 아이콘 대신 앱 아이콘이 표시되도록 설정
    if sys.platform == "win32":
        try:
            import ctypes

            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "kkongsooni.ocellr.nhout.1"
            )
        except Exception:
            pass

    app = QApplication(sys.argv)
    _apply_light_palette(app)
    app.setWindowIcon(_create_app_icon())
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        input("\n오류가 발생했습니다. 종료하려면 Enter 키를 누르세요...")