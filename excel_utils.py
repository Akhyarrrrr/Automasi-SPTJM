from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, Optional, Tuple

import pandas as pd

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


@dataclass
class ExcelMeta:
    sheet_names: list[str]
    row_count: int
    col_count: int
    has_email_col: bool


def get_sheet_names(excel_bytes) -> list[str]:
    xls = pd.ExcelFile(excel_bytes, engine="openpyxl")
    return list(xls.sheet_names)


def read_excel_any(excel_bytes, sheet_name: str | None):
    """
    Load df + normalisasi nama kolom + paksa kolom identitas jadi string
    (biar NIP/Norek aman & Streamlit tidak spam Arrow warning).
    """
    df = pd.read_excel(excel_bytes, sheet_name=sheet_name, engine="openpyxl")
    if isinstance(df, dict):
        df = df[next(iter(df.keys()))]

    df.columns = [str(c).strip() for c in df.columns]

    # identitas = string (identifier, bukan numeric)
    for col in ["NIP", "Norek"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    return df


def normalize_email(s: str | None) -> Optional[str]:
    if s is None:
        return None
    s = str(s).strip()
    if not s:
        return None
    return s if _EMAIL_RE.match(s) else None


def build_email_map_from_df(df_map: pd.DataFrame) -> Dict[str, str]:
    """
    df_map harus punya kolom: NIP, Email (case-insensitive).
    """
    df_map = df_map.copy()
    df_map.columns = [str(c).strip() for c in df_map.columns]
    cols = {c.lower(): c for c in df_map.columns}

    if "nip" not in cols or "email" not in cols:
        raise ValueError("File mapping email wajib punya kolom: NIP dan Email")

    nip_col = cols["nip"]
    email_col = cols["email"]

    out: Dict[str, str] = {}
    for _, r in df_map.iterrows():
        nip = str(r.get(nip_col, "")).strip()
        em = normalize_email(r.get(email_col, None))
        if nip and em:
            out[nip] = em
    return out


def apply_email_map(df_main: pd.DataFrame, email_map: Dict[str, str]) -> pd.DataFrame:
    """
    Tambahkan/isi kolom Email pada df_main berdasarkan NIP.
    """
    df_main = df_main.copy()
    if "Email" not in df_main.columns and "email" not in df_main.columns:
        df_main["Email"] = ""

    # kalau kolomnya "email", seragamkan jadi "Email" biar konsisten
    if "email" in df_main.columns and "Email" not in df_main.columns:
        df_main["Email"] = df_main["email"]
        df_main.drop(columns=["email"], inplace=True)

    def _fill(r):
        current = str(r.get("Email", "")).strip()
        if current:
            return current
        nip = str(r.get("NIP", "")).strip()
        return email_map.get(nip, "")

    df_main["Email"] = df_main.apply(_fill, axis=1)
    return df_main


def detect_max_n(df: pd.DataFrame) -> int:
    no_prop_cols = [c for c in df.columns if re.fullmatch(r"NoProp\d+", str(c).strip())]
    if not no_prop_cols:
        raise ValueError(
            "Format Excel tidak sesuai: tidak ada kolom NoProp1/NoProp2/..."
        )

    nums = sorted(int(re.findall(r"\d+", c)[0]) for c in no_prop_cols)
    return max(nums)


def validate_required_columns(
    df: pd.DataFrame, require_email: bool
) -> Tuple[bool, str]:
    required_identity = ["NIP", "Nama", "Fakultas", "Norek"]
    missing = [c for c in required_identity if c not in df.columns]
    if missing:
        return False, f"Kolom wajib tidak ditemukan: {missing}"

    try:
        _ = detect_max_n(df)
    except Exception as e:
        return False, str(e)

    if require_email:
        if "Email" not in df.columns and "email" not in df.columns:
            return (
                False,
                "Mode email aktif tapi kolom Email tidak ada. Tambahkan Email atau upload mapping NIP→Email.",
            )
    return True, "OK"
