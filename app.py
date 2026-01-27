from __future__ import annotations

import io
import os
import random
import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from excel_utils import (
    get_sheet_names,
    read_excel_any,
    validate_required_columns,
    build_email_map_from_df,
    apply_email_map,
)
from generator import (
    iter_people_from_df,
    generate_pdf_bytes_for_person,
    generate_zip_pdf,
    filename_for_person,
    ResultRow,
)
from emailer import send_smtp

load_dotenv()

st.set_page_config(page_title="SPTJM Otomatis", page_icon="📄", layout="centered")

st.title("📄 SPTJM Otomatis (Excel → PDF → (Opsional) Email)")
st.caption(
    "Alur aman: Generate PDF dulu → cek sample → baru kirim email (tidak auto-kirim)."
)

# =========================
# Session State
# =========================
if "df_main" not in st.session_state:
    st.session_state.df_main = None
if "generated_pdf_items" not in st.session_state:
    st.session_state.generated_pdf_items = None  # list[(filename, bytes, person)]
if "generated_zip_bytes" not in st.session_state:
    st.session_state.generated_zip_bytes = None
if "generate_report" not in st.session_state:
    st.session_state.generate_report = None
if "email_report" not in st.session_state:
    st.session_state.email_report = None


# =========================
# Sidebar (pengaturan teknis)
# =========================
with st.sidebar:
    st.header("⚙️ Konfigurasi Sistem")

    soffice_default = os.getenv(
        "SOFFICE_PATH", r"C:\Program Files\LibreOffice\program\soffice.exe"
    )
    soffice_path = st.text_input(
        "Path LibreOffice (soffice)",
        value=soffice_default,
        help="Jika LibreOffice tidak terdeteksi, isi manual. Kalau sudah benar, biarkan.",
    )

    st.divider()
    st.subheader("📧 SMTP Status (.env)")

    smtp_host = os.getenv("SMTP_HOST", "")
    smtp_port = os.getenv("SMTP_PORT", "")
    smtp_user = os.getenv("SMTP_USER", "")

    if smtp_host and smtp_user:
        st.success("SMTP terkonfigurasi ✅")
        st.caption(f"Host: {smtp_host}:{smtp_port or '587'}")
        st.caption(f"User: {smtp_user}")
    else:
        st.warning("SMTP belum siap (hanya bisa Generate PDF, tidak bisa kirim email).")
        st.caption("Isi SMTP_HOST/SMTP_PORT/SMTP_USER/SMTP_PASS di .env")


# =========================
# Step 1: Upload Excel
# =========================
st.subheader("1) Upload Excel")
uploaded = st.file_uploader("Upload Excel utama (.xlsx)", type=["xlsx"])

if not uploaded:
    st.info("Silakan upload Excel terlebih dahulu.")
    st.stop()

excel_bytes = io.BytesIO(uploaded.getvalue())

try:
    sheet_names = get_sheet_names(excel_bytes)
except Exception as e:
    st.error("Gagal membaca Excel. Pastikan file .xlsx valid.")
    st.exception(e)
    st.stop()

sheet = st.selectbox("Pilih sheet", sheet_names, index=0)

c1, c2 = st.columns(2)
with c1:
    btn_load = st.button("🔎 Load & Preview", use_container_width=True)
with c2:
    if st.button("♻️ Reset Hasil (hapus PDF/ZIP/report)", use_container_width=True):
        st.session_state.generated_pdf_items = None
        st.session_state.generated_zip_bytes = None
        st.session_state.generate_report = None
        st.session_state.email_report = None
        st.success("Hasil di-reset.")
        st.stop()

if btn_load or st.session_state.df_main is None:
    try:
        excel_bytes.seek(0)
        df_main = read_excel_any(excel_bytes, sheet_name=sheet)
        st.session_state.df_main = df_main
    except Exception as e:
        st.error("Gagal load sheet.")
        st.exception(e)
        st.stop()

df_main: pd.DataFrame = st.session_state.df_main

st.success(
    f"Sheet loaded: **{sheet}** | Rows: **{len(df_main)}** | Cols: **{len(df_main.columns)}**"
)
with st.expander("Preview 10 baris pertama", expanded=False):
    st.dataframe(df_main.head(10).astype("string"), use_container_width=True)


# =========================
# Step 2: Generate PDF (NO EMAIL)
# =========================
st.divider()
st.subheader("2) Generate PDF (Tanpa Email dulu — untuk dicek)")

ok, msg = validate_required_columns(df_main, require_email=False)
if not ok:
    st.error(msg)
    st.stop()
else:
    st.caption("Validasi kolom: OK ✅")

sample_count = st.number_input(
    "Berapa sample PDF untuk dicek cepat?", min_value=1, max_value=20, value=5
)
shuffle_sample = st.checkbox("Acak sample (random)", value=False)

btn_generate = st.button("🚀 Generate ZIP PDF (Tanpa Email)", use_container_width=True)

if btn_generate:
    people = list(iter_people_from_df(df_main))
    total = len(people)

    if total == 0:
        st.warning("Tidak ada baris yang valid (mungkin semua NoProp kosong).")
        st.stop()

    progress = st.progress(0)
    status_box = st.empty()
    log_box = st.empty()

    pdf_items_full = []  # (filename, pdf_bytes, person)
    report_rows = []

    with st.status("Memulai generate PDF...", expanded=True) as st_status:
        st.write(f"Total penerima valid: **{total}**")

        for idx, (person, lampiran_rows) in enumerate(people, start=1):
            progress.progress(idx / total)
            status_box.info(
                f"[{idx}/{total}] Generating: {person.nama} | NIP: {person.nip}"
            )

            try:
                st_status.update(
                    label=f"Generating PDF: {person.nama}", state="running"
                )

                pdf_bytes = generate_pdf_bytes_for_person(
                    person,
                    lampiran_rows,
                    soffice_path=soffice_path.strip() or None,
                )
                fname = filename_for_person(person)
                pdf_items_full.append((fname, pdf_bytes, person))
                report_rows.append(
                    ResultRow(
                        person.nama, person.nip, person.email or "", "OK", "PDF dibuat"
                    )
                )

            except Exception as e:
                report_rows.append(
                    ResultRow(
                        person.nama, person.nip, person.email or "", "FAIL", str(e)
                    )
                )
                log_box.error(f"❌ ERROR {person.nama} ({person.nip}): {e}")

        st_status.update(label="Generate selesai ✅", state="complete")

    zip_bytes = generate_zip_pdf([(fn, b) for (fn, b, _) in pdf_items_full])

    st.session_state.generated_pdf_items = pdf_items_full
    st.session_state.generated_zip_bytes = zip_bytes
    st.session_state.generate_report = pd.DataFrame([r.__dict__ for r in report_rows])
    st.session_state.email_report = None

    st.success("ZIP PDF berhasil dibuat. Silakan download & cek sample di bawah.")


# Output ZIP + sample + report generate
if st.session_state.generated_zip_bytes:
    st.download_button(
        "⬇️ Download ZIP semua PDF",
        data=st.session_state.generated_zip_bytes,
        file_name="SPTJM_PDF.zip",
        mime="application/zip",
        use_container_width=True,
    )

    st.subheader("✅ Sample PDF untuk dicek cepat")
    items = st.session_state.generated_pdf_items or []
    if items:
        idxs = list(range(len(items)))
        if shuffle_sample:
            random.shuffle(idxs)
        idxs = idxs[: int(sample_count)]

        for i in idxs:
            fname, pdf_bytes, person = items[i]
            with st.expander(f"{person.nama} | {person.nip} — {fname}", expanded=False):
                st.download_button(
                    "⬇️ Download PDF ini",
                    data=pdf_bytes,
                    file_name=fname,
                    mime="application/pdf",
                    use_container_width=True,
                )

    st.subheader("📊 Report Generate")
    df_report = st.session_state.generate_report
    st.dataframe(df_report.astype("string"), use_container_width=True)
    st.download_button(
        "⬇️ Download Report Generate CSV",
        data=df_report.to_csv(index=False).encode("utf-8"),
        file_name="SPTJM_generate_report.csv",
        mime="text/csv",
        use_container_width=True,
    )


# =========================
# Step 3: Send Email (ONLY after Generate)
# =========================
st.divider()
st.subheader("3) Kirim Email (Opsional — setelah kamu cek sample PDF)")

if not st.session_state.generated_pdf_items:
    st.info("Generate PDF dulu (Step 2). Setelah itu, fitur kirim email akan muncul.")
    st.stop()

want_email = st.checkbox("Saya ingin kirim email sekarang", value=False)
if not want_email:
    st.stop()

# Email column / mapping
need_mapping = "Email" not in df_main.columns and "email" not in df_main.columns
if need_mapping:
    st.warning(
        "Kolom Email tidak ada di Excel utama. Upload file mapping (NIP, Email)."
    )
    map_file = st.file_uploader("Upload file mapping email (.xlsx)", type=["xlsx"])
    if not map_file:
        st.stop()

    try:
        df_map = pd.read_excel(io.BytesIO(map_file.getvalue()), engine="openpyxl")
        email_map = build_email_map_from_df(df_map)
        df_main2 = apply_email_map(df_main, email_map)
        st.session_state.df_main = df_main2
        df_main = df_main2
        st.success(f"Mapping email masuk: {len(email_map)} data.")
    except Exception as e:
        st.error("Gagal baca mapping email.")
        st.exception(e)
        st.stop()

# SMTP ready?
if not (os.getenv("SMTP_HOST") and os.getenv("SMTP_USER") and os.getenv("SMTP_PASS")):
    st.error(
        "SMTP belum siap di .env. Kamu bisa generate PDF, tapi tidak bisa kirim email."
    )
    st.stop()

dry_run = st.checkbox(
    "Dry-run (tidak benar-benar kirim email, hanya simulasi)", value=True
)
delay = st.number_input(
    "Delay antar email (detik) — anti limit",
    min_value=0.0,
    max_value=5.0,
    value=0.7,
    step=0.1,
)

subject_tpl = st.text_input("Subject email", value="SPTJM - {nama} ({nip})")
body_tpl = st.text_area(
    "Body email",
    value="Yth. Bapak/Ibu {nama},\n\nBerikut kami kirimkan file SPTJM (PDF).\n\nTerima kasih.\n",
    height=140,
)

confirm = st.checkbox(
    "✅ Saya sudah cek sample PDF dan yakin datanya sudah benar.", value=False
)

if not confirm:
    st.warning("Centang konfirmasi dulu agar tombol kirim email aktif.")
    st.stop()

btn_send = st.button(
    "📧 KIRIM EMAIL SEKARANG", type="primary", use_container_width=True
)

if btn_send:
    # Build nip->email from updated df_main
    df_email = df_main.copy()
    df_email.columns = [str(c).strip() for c in df_email.columns]
    email_col = (
        "Email"
        if "Email" in df_email.columns
        else ("email" if "email" in df_email.columns else None)
    )

    nip_to_email = {}
    if email_col:
        for _, r in df_email.iterrows():
            nip = str(r.get("NIP", "")).strip()
            em = str(r.get(email_col, "")).strip()
            if nip and em:
                nip_to_email[nip] = em

    items = st.session_state.generated_pdf_items
    total = len(items)

    progress = st.progress(0)
    status_box = st.empty()
    log_box = st.empty()

    report_rows = []

    with st.status("Mengirim email...", expanded=True) as st_status:
        for idx, (fname, pdf_bytes, person) in enumerate(items, start=1):
            progress.progress(idx / total)

            to_email = (person.email or "").strip()
            if not to_email:
                to_email = nip_to_email.get(person.nip, "").strip()

            status_box.info(f"[{idx}/{total}] {person.nama} → {to_email or '-'}")

            if not to_email:
                report_rows.append(
                    ResultRow(
                        person.nama, person.nip, "", "SKIP", "Email kosong/invalid"
                    )
                )
                continue

            try:
                st_status.update(label=f"Sending: {to_email}", state="running")
                subject = subject_tpl.format(nama=person.nama, nip=person.nip)
                body = body_tpl.format(nama=person.nama, nip=person.nip)

                if not dry_run:
                    send_smtp(
                        to_email=to_email,
                        subject=subject,
                        body=body,
                        attachment_bytes=pdf_bytes,
                        attachment_filename=fname,
                        sleep_seconds=float(delay),
                    )
                    report_rows.append(
                        ResultRow(
                            person.nama, person.nip, to_email, "OK", "Email terkirim"
                        )
                    )
                else:
                    report_rows.append(
                        ResultRow(
                            person.nama,
                            person.nip,
                            to_email,
                            "DRY-RUN",
                            "Simulasi kirim email",
                        )
                    )

            except Exception as e:
                report_rows.append(
                    ResultRow(person.nama, person.nip, to_email, "FAIL", str(e))
                )
                log_box.error(f"❌ FAIL {person.nama} ({to_email}): {e}")

        st_status.update(label="Selesai kirim email ✅", state="complete")

    st.session_state.email_report = pd.DataFrame([r.__dict__ for r in report_rows])
    st.success("Proses email selesai. Cek report di bawah.")

# Email report
if st.session_state.email_report is not None:
    st.subheader("📊 Report Email")
    df_er = st.session_state.email_report
    st.dataframe(df_er.astype("string"), use_container_width=True)
    st.download_button(
        "⬇️ Download Report Email CSV",
        data=df_er.to_csv(index=False).encode("utf-8"),
        file_name="SPTJM_email_report.csv",
        mime="text/csv",
        use_container_width=True,
    )
