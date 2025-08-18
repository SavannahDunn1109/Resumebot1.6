import streamlit as st
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.files.file import File
import io
import pandas as pd
import os
from PyPDF2 import PdfReader
from docx import Document
import re
from datetime import date

# ========== CONFIG ==========
SITE_URL = "https://eleven090.sharepoint.com/sites/Recruiting"
LIBRARY = "Shared Documents"
FOLDER = "Active Resumes"

# ========== AUTH ==========
@st.cache_resource
def connect_to_sharepoint():
    ctx_auth = AuthenticationContext(SITE_URL)
    if not ctx_auth.acquire_token_for_user(
        st.secrets["sharepoint"]["username"],
        st.secrets["sharepoint"]["password"]
    ):
        st.error("Authentication failed")
        return None
    return ClientContext(SITE_URL, ctx_auth)

# ========== FILE HELPERS ==========
def download_file(ctx, file_url):
    response = File.open_binary(ctx, file_url)
    return io.BytesIO(response.content)

def extract_text_from_pdf(file_bytes):
    text = ""
    reader = PdfReader(file_bytes)
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    return text

def extract_text_from_docx(file_bytes):
    doc = Document(file_bytes)
    return "\n".join([para.text for para in doc.paragraphs])

# ========== EXPERIENCE HELPERS ==========
MONTHS = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}

def _mk_date(y: int, m: int) -> date:
    m = min(max(1, m), 12)
    return date(int(y), int(m), 15)

def _parse_month(token: str):
    if not token:
        return None
    t = token.strip().lower()
    return MONTHS.get(t)

def _parse_year(token: str):
    if not token:
        return None
    m = re.match(r"(19|20)\d{2}$", token.strip())
    return int(m.group(0)) if m else None

def _present_to_date() -> date:
    today = date.today()
    return date(today.year, today.month, 15)

def _extract_date_ranges(text: str):
    t = text.replace("\u2013", "-").replace("\u2014", "-")
    ranges = []

    pat_month_year = re.compile(
        r"\b(?P<m1>[A-Za-z]{3,9})\s+(?P<y1>(?:19|20)\d{2})\s*[-to]+\s*(?P<m2>Present|Current|[A-Za-z]{3,9})\s*(?P<y2>(?:19|20)\d{2})?\b",
        flags=re.I
    )
    for m in pat_month_year.finditer(t):
        m1 = _parse_month(m.group("m1"))
        y1 = _parse_year(m.group("y1"))
        m2tok = m.group("m2")
        y2tok = m.group("y2")
        if m1 and y1:
            start = _mk_date(y1, m1)
            if m2tok and m2tok.lower() in ("present", "current"):
                end = _present_to_date()
            else:
                m2 = _parse_month(m2tok)
                y2 = _parse_year(y2tok) if y2tok else None
                if m2 and y2:
                    end = _mk_date(y2, m2)
                else:
                    continue
            if end > start:
                ranges.append((start, end))

    pat_year_year = re.compile(
        r"\b(?P<y1>(?:19|20)\d{2})\s*[-to]+\s*(?P<y2>Present|Current|(?:19|20)\d{2})\b",
        flags=re.I
    )
    for m in pat_year_year.finditer(t):
        y1 = _parse_year(m.group("y1"))
        y2tok = m.group("y2")
        if not y1:
            continue
        start = _mk_date(y1, 6)
        if y2tok.lower() in ("present", "current"):
            end = _present_to_date()
        else:
            y2 = _parse_year(y2tok)
            if not y2:
                continue
            end = _mk_date(y2, 6)
        if end > start:
            ranges.append((start, end))

    pat_mmyyyy = re.compile(
        r"\b(?P<m1>0?[1-9]|1[0-2])/(?P<y1>(?:19|20)\d{2})\s*[-to]+\s*(?P<m2>0?[1-9]|1[0-2])/(?P<y2>(?:19|20)\d{2}|Present|Current)\b",
        flags=re.I
    )
    for m in pat_mmyyyy.finditer(t):
        m1 = int(m.group("m1"))
        y1 = _parse_year(m.group("y1"))
        if not (y1 and 1 <= m1 <= 12):
            continue
        start = _mk_date(y1, m1)
        y2raw = m.group("y2")
        if y2raw.lower() in ("present", "current"):
            end = _present_to_date()
        else:
            m2 = int(m.group("m2"))
            y2 = _parse_year(y2raw)
            if not (y2 and 1 <= m2 <= 12):
                continue
            end = _mk_date(y2, m2)
        if end > start:
            ranges.append((start, end))

    if not ranges:
        return []
    ranges.sort(key=lambda r: r[0])
    merged = [ranges[0]]
    for s, e in ranges[1:]:
        last_s, last_e = merged[-1]
        if s <= last_e:
            merged[-1] = (last_s, max(last_e, e))
        else:
            merged.append((s, e))
    return merged

def _years_from_ranges(text: str) -> float:
    merged = _extract_date_ranges(text)
    total_months = 0
    for s, e in merged:
        diff = (e.year - s.year) * 12 + (e.month - s.month)
        total_months += max(0, diff)
    return round(total_months / 12.0, 1)

def _years_from_phrases(text: str) -> int:
    best = 0
    for m in re.finditer(r"\b([1-4]?\d)\s*\+?\s*[- ]?\s*(?:years?|yrs?)\b", text, flags=re.I):
        best = max(best, int(m.group(1)))
    return best

def estimate_years_experience(text: str):
    yrs_ranges = _years_from_ranges(text)
    yrs_phrases = _years_from_phrases(text)
    if yrs_ranges >= 0.5:
        return yrs_ranges, "ranges"
    return float(yrs_phrases), "phrases"

def classify_level(years: float, jr_max: int, mid_max: int) -> str:
    if years <= jr_max:
        return "Junior"
    elif years <= mid_max:
        return "Mid"
    else:
        return "Senior"

# ========== SCORING ==========
uploaded_req_file = st.file_uploader("📄 Upload Requirements (.txt)", type=["txt"])

KEYWORDS = []
if uploaded_req_file:
    req_lines = uploaded_req_file.read().decode("utf-8").splitlines()
    for line in req_lines:
        line = line.strip()
        if line and not any(line.startswith(prefix) for prefix in ("🧠", "💼", "🛡", "⚙️", "☁️", "👥", "🎯", "🧾", "🧩")):
            if not line.endswith(":"):
                KEYWORDS.append(line)
    st.success(f"✅ Loaded {len(KEYWORDS)} keywords from requirements file.")
else:
    st.warning("⚠️ Please upload a requirements .txt file to begin scoring.")
    st.stop()

# Scoring & filters UI
st.subheader("⚙️ Scoring & Filters")
exp_points_per_year = st.number_input("Points per year of experience", min_value=0, max_value=50, value=5, step=1)
jr_max = st.number_input("Max years for JUNIOR", min_value=0, max_value=10, value=2, step=1)
mid_max = st.number_input("Max years for MID", min_value=jr_max, max_value=25, value=6, step=1)
enforce_min = st.checkbox("Enforce minimum years of experience filter?", value=False)
min_years_required = st.number_input("Minimum years (hide resumes below this)", min_value=0, max_value=30, value=3, step=1)

def score_resume(text: str):
    kw_score = 0
    found_keywords = []
    lower_text = text.lower()
    for kw in KEYWORDS:
        if kw.lower() in lower_text:
            kw_score += 10
            found_keywords.append(kw)

    years, years_source = estimate_years_experience(text)
    exp_score = years * exp_points_per_year
    total = kw_score + exp_score

    return {
        "years": years,
        "years_source": years_source,
        "level": classify_level(years, jr_max, mid_max),
        "kw_score": kw_score,
        "exp_score": exp_score,
        "total": total,
        "keywords_found": ", ".join(found_keywords),
    }

# ========== MAIN STREAMLIT APP ==========
st.title("📄 Resume Scorer from SharePoint")

ctx = connect_to_sharepoint()
if ctx:
    folder_url = f"{LIBRARY}/{FOLDER}"
    folder = ctx.web.get_folder_by_server_relative_url(folder_url)
    files = folder.files
    ctx.load(files)
    ctx.execute_query()

    data = []

    for file in files:
        filename = file.properties["Name"]
        if filename.lower().endswith(".pdf") or filename.lower().endswith(".docx"):
            file_url = file.properties["ServerRelativeUrl"]
            file_bytes = download_file(ctx, file_url)

            if filename.lower().endswith(".pdf"):
                text = extract_text_from_pdf(file_bytes)
            else:
                text = extract_text_from_docx(file_bytes)

            result = score_resume(text)
            if enforce_min and result["years"] < float(min_years_required):
                continue

            data.append({
                "File Name": filename,
                "Est. Years": result["years"],
                "Level (Jr/Mid/Sr)": result["level"],
                "Experience Source": result["years_source"],
                "Keyword Score": result["kw_score"],
                "Experience Score": result["exp_score"],
                "Total Score": result["total"],
                "Keywords Found": result["keywords_found"],
            })

    df = pd.DataFrame(data)
    if not df.empty:
        df = df.sort_values(
            ["Level (Jr/Mid/Sr)", "Est. Years", "Total Score"],
            ascending=[True, False, False]
        ).reset_index(drop=True)
    st.dataframe(df)

    if not df.empty:
        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)
        st.download_button("📥 Download Excel Report", output, file_name="resume_scores.xlsx")

        # Upload Excel back to SharePoint
        if st.button("📤 Upload Excel to SharePoint"):
            upload_path = f"{LIBRARY}/resume_scores.xlsx"
            target_folder = ctx.web.get_folder_by_server_relative_url(LIBRARY)
            target_folder.upload_file("resume_scores.xlsx", output)
            ctx.execute_query()
            st.success("Excel uploaded to SharePoint!")

