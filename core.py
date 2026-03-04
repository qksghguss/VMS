
import os
import re
import pandas as pd
from datetime import datetime
from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional
import configparser

FIELDS = ["일자","LOT","작업자","매수","Defect_L1","Defect_L2","구분","L1 담당자","L2 담당자"]

def _norm_name(s: str) -> str:
    if s is None:
        return ""
    return re.sub(r"\s+", "", str(s)).strip()

def split_people(s):
    if s is None:
        return []
    s = str(s).strip()
    if not s:
        return []
    parts = [p.strip() for p in s.split("/") if p.strip()]
    seen = set()
    out = []
    for p in parts:
        key = _norm_name(p)
        if key and key not in seen:
            seen.add(key)
            out.append(p.strip())
    return out

def lot_to_count(x):
    if x is None:
        return 0
    s = str(x).strip()
    return 1 if s and s.lower() != "nan" else 0

def parse_date(x):
    if pd.isna(x):
        return None
    if isinstance(x, datetime):
        return x.date()
    try:
        if hasattr(x, "to_pydatetime"):
            return x.to_pydatetime().date()
    except Exception:
        pass
    s = str(x).strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d","%Y/%m/%d","%Y.%m.%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            continue
    try:
        dt = pd.to_datetime(s, errors="coerce")
        if pd.isna(dt):
            return None
        return dt.date()
    except Exception:
        return None


def to_period_label(d, granularity: str) -> str:
    if granularity == "weekly":
        iso_year, iso_week, _ = d.isocalendar()
        return f"{iso_year}-W{iso_week:02d}"
    if granularity == "monthly":
        return f"{d.year:04d}-{d.month:02d}"
    return d.isoformat()

def col_letter_to_index(letter: str) -> int:
    letter = letter.strip().upper()
    n = 0
    for ch in letter:
        if not ("A" <= ch <= "Z"):
            raise ValueError(letter)
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1

def select_columns(df: pd.DataFrame, mapping: Dict[str,str]) -> Dict[str,str]:
    out = {}
    for field, src in mapping.items():
        if not src:
            out[field] = None
            continue
        if src in df.columns:
            out[field] = src
        else:
            try:
                idx = col_letter_to_index(src)
                out[field] = df.columns[idx] if 0 <= idx < df.shape[1] else None
            except Exception:
                out[field] = None
    return out

@dataclass
class SourceSpec:
    use: bool
    name: str
    path: str
    sheet: str
    header_row: int = 1

@dataclass
class AppConfig:
    sources: List[SourceSpec]
    mapping: Dict[str,str]
    exclude: set
    start_date: Optional[datetime.date]
    end_date: Optional[datetime.date]
    granularity: str
    rank_metric: str
    round_decimals: int
    result_xlsx: str

def load_ini(path: str) -> AppConfig:
    cp = configparser.ConfigParser()
    cp.read(path, encoding="utf-8")

    # app
    round_decimals = int(cp.get("app","round_decimals", fallback="3"))
    rank_metric = cp.get("app","rank_metric", fallback="pages").lower()

    # period
    s = cp.get("period","start_date", fallback="").strip()
    e = cp.get("period","end_date", fallback="").strip()
    start_date = parse_date(s) if s else None
    end_date = parse_date(e) if e else None
    granularity = cp.get("period","granularity", fallback="monthly").lower()

    # output
    result_xlsx = cp.get("output","result_xlsx", fallback="result.xlsx")

    # sources
    sources = []
    for sec in cp.sections():
        if sec.lower().startswith("source_"):
            use = cp.get(sec,"use", fallback="0").strip() in ("1","y","Y","true","True")
            name = cp.get(sec,"name", fallback=sec)
            pathv = cp.get(sec,"path", fallback="").strip()
            sheet = cp.get(sec,"sheet", fallback="Sheet1").strip()
            header_row = int(cp.get(sec,"header_row", fallback="1"))
            sources.append(SourceSpec(use=use,name=name,path=pathv,sheet=sheet,header_row=header_row))
    sources.sort(key=lambda x: x.name)

    # mapping
    mapping = {f: cp.get("columns", f, fallback=f).strip() for f in FIELDS}

    # exclude
    names = cp.get("exclude","names", fallback="").strip()
    exclude = {_norm_name(x) for x in names.split(",") if _norm_name(x)}

    return AppConfig(
        sources=sources,
        mapping=mapping,
        exclude=exclude,
        start_date=start_date,
        end_date=end_date,
        granularity=granularity,
        rank_metric=rank_metric,
        round_decimals=round_decimals,
        result_xlsx=result_xlsx
    )

def aggregate(cfg: AppConfig) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, str]:
    """
    returns: expanded_df, daily_df, summary_df, detail_df, sort_col
    """
    records = []
    needed = FIELDS

    for s in cfg.sources:
        if not s.use:
            continue
        if not s.path or not os.path.exists(s.path):
            raise FileNotFoundError(f"Source file not found: {s.path}")
        df = pd.read_excel(s.path, sheet_name=s.sheet if s.sheet else 0, header=s.header_row-1, engine="openpyxl")
        colmap = select_columns(df, cfg.mapping)
        missing = [f for f in needed if colmap.get(f) is None]
        if missing:
            raise ValueError(f"[{s.name}] Missing mapped columns: {missing}")

        for _, row in df.iterrows():
            if str(row[colmap["구분"]]).strip() != "정상":
                continue
            d = parse_date(row[colmap["일자"]])
            if d is None:
                continue
            if cfg.start_date and d < cfg.start_date:
                continue
            if cfg.end_date and d > cfg.end_date:
                continue

            workers = split_people(row[colmap["작업자"]])
            if not workers:
                continue
            n = len(workers)

            lot_cnt = lot_to_count(row[colmap["LOT"]])
            pages = 0.0 if pd.isna(row[colmap["매수"]]) else float(row[colmap["매수"]])
            l1 = 0.0 if pd.isna(row[colmap["Defect_L1"]]) else float(row[colmap["Defect_L1"]])
            l2 = 0.0 if pd.isna(row[colmap["Defect_L2"]]) else float(row[colmap["Defect_L2"]])

            l1_owners = split_people(row[colmap["L1 담당자"]])
            l2_owners = split_people(row[colmap["L2 담당자"]])
            l1_den = len(l1_owners)
            l2_den = len(l2_owners)
            l1_owner_norm = {_norm_name(x) for x in l1_owners}
            l2_owner_norm = {_norm_name(x) for x in l2_owners}

            lot_share = lot_cnt / n
            pages_share = pages / n

            for w in workers:
                wn = _norm_name(w)
                records.append({
                    "일자": d,
                    "작업일보": s.name,
                    "작업자": w,
                    "LOT": lot_share,
                    "매수": pages_share,
                    "Defect_L1": (l1 / l1_den) if (l1_den and wn in l1_owner_norm) else 0.0,
                    "Defect_L2": (l2 / l2_den) if (l2_den and wn in l2_owner_norm) else 0.0,
                    "제외여부": wn in cfg.exclude,  # 분해는 하되 집계에서 제외
                })

    expanded = pd.DataFrame(records)
    if expanded.empty:
        raise ValueError("분해 결과가 비어있습니다. (구분=정상/기간/경로/시트/매핑 확인)")

    inc = expanded[expanded["제외여부"] == False].copy()

    inc["기간"] = inc["일자"].apply(lambda d: to_period_label(d, cfg.granularity))

    daily = (inc.groupby(["기간", "작업자"], as_index=False)[["LOT","매수","Defect_L1","Defect_L2"]].sum())
    detail = (inc.groupby(["작업자","작업일보"], as_index=False)[["LOT","매수","Defect_L1","Defect_L2"]].sum())
    summary = (inc.groupby(["작업자"], as_index=False)[["LOT","매수","Defect_L1","Defect_L2"]].sum())

    dec = cfg.round_decimals
    for df in (daily, detail, summary):
        df[["LOT","매수","Defect_L1","Defect_L2"]] = df[["LOT","매수","Defect_L1","Defect_L2"]].round(dec)

    metric_map = {"lot":"LOT","pages":"매수","l1":"Defect_L1","l2":"Defect_L2"}
    sort_col = metric_map.get(cfg.rank_metric, "매수")
    summary = summary.sort_values(sort_col, ascending=False).reset_index(drop=True)
    summary.insert(0, "순위", range(1, len(summary)+1))

    return expanded, daily, summary, detail, sort_col

def export_excel(path: str, daily: pd.DataFrame, summary: pd.DataFrame, detail: pd.DataFrame):
    # keep simple: pandas ExcelWriter
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        daily.to_excel(w, index=False, sheet_name="일일집계")
        summary.to_excel(w, index=False, sheet_name="최종집계_요약")
        detail.to_excel(w, index=False, sheet_name="최종집계_상세")
