
# app_improved_resume_scorer.py
# Streamlit app with better progress, timeouts, and safer fallbacks.
import io
import re
import time
from datetime import datetime
from typing import List, Dict, Any

import streamlit as st
import pandas as pd

# Optional dependencies used if available in the environment
try:
    from office365.runtime.client_request_exception import ClientRequestException
    from office365.sharepoint.client_context import ClientContext
    from office365.sharepoint.files.file import File
except Exception:
    ClientRequestException = Exception
    ClientContext = None
    File = None

try:
    import msal
except Exception:
    msal = None

try:
    from PyPDF2 import PdfReader
except Exception:
    PdfReader = None

try:
    from docx import Document as DocxDocument
except Exception:
    DocxDocument = None


# ------------------------------
# Utility: config
# ------------------------------
def get_secret(path: List[str], default=None):
    try:
        d = st.secrets
        for p in path:
            d = d[p]
        return d
    except Exception:
        return default


def get_config():
    """Load SharePoint/Azure configuration from st.secrets or sidebar inputs."""
    tenant_id = get_secret(["sharepoint", "tenant_id"])
    client_id = get_secret(["sharepoint", "client_id"])
    client_secret = get_secret(["sharepoint", "client_secret"])
    site_url = get_secret(["sharepoint", "site_url"])

    with st.sidebar.expander("🔧 Connection (fallback inputs)"):
        tenant_id = st.text_input("Tenant ID", tenant_id or "", type="default")
        client_id = st.text_input("Client (Application) ID", client_id or "", type="default")
        client_secret = st.text_input("Client Secret (Value)", client_secret or "", type="password")
        site_url = st.text_input("SharePoint Site URL", site_url or "https://<tenant>.sharepoint.com/sites/Recruiting")

    return tenant_id.strip(), client_id.strip(), client_secret.strip(), site_url.strip()


# ------------------------------
# Auth & SharePoint helpers
# ------------------------------
def acquire_sp_token(tenant_id: str, client_id: str, client_secret: str, site_url: str) -> str:
    if msal is None:
        raise RuntimeError("MSAL is not installed in this environment.")
    if not all([tenant_id, client_id, client_secret, site_url]):
        raise ValueError("Missing tenant_id, client_id, client_secret, or site_url.")

    # Scope: SharePoint resource-specific scope
    # For app-only, use the host of the site_url
    host = site_url.split("/")[2]
    scope = [f"https://{host}/.default"]

    st.write("🔐 Requesting app-only token…")
    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
    )
    token = app.acquire_token_for_client(scopes=scope)
    if "access_token" not in token:
        raise RuntimeError(f"MSAL error: {token}")
    st.write("🔑 Token acquired.")
    return token["access_token"]


def connect_with_azure_app(tenant_id: str, client_id: str, client_secret: str, site_url: str):
    if ClientContext is None:
        raise RuntimeError("office365-rest-python-client is not installed in this environment.")
    token = acquire_sp_token(tenant_id, client_id, client_secret, site_url)
    st.write("🌐 Pinging SharePoint site…")
    ctx = ClientContext(site_url).with_access_token(token)
    ctx.web.get().execute_query()
    st.success("✅ SharePoint reachable.")
    return ctx


def get_folder(ctx, server_relative_folder: str):
    if not server_relative_folder.startswith("/"):
        server_relative_folder = "/" + server_relative_folder
    folder = ctx.web.get_folder_by_server_relative_url(server_relative_folder)
    folder.expand(["Files"]).get().execute_query()
    return folder


def download_file(ctx, server_relative_url: str) -> io.BytesIO:
    response = File.open_binary(ctx, server_relative_url)
    bio = io.BytesIO()
    bio.write(response.content)
    bio.seek(0)
    return bio


# ------------------------------
# Extraction & Scoring
# ------------------------------
def extract_text_from_pdf(bio: io.BytesIO, timeout_seconds: int) -> str:
    if PdfReader is None:
        raise RuntimeError("PyPDF2 is not installed.")
    start = time.time()
    bio.seek(0)
    reader = PdfReader(bio)
    text_parts = []
    for i, page in enumerate(reader.pages):
        try:
            t = page.extract_text() or ""
        except Exception:
            t = ""
        text_parts.append(t)
        if time.time() - start > timeout_seconds:
            raise TimeoutError("PDF parse timeout")
    return "\n".join(text_parts).strip()


def extract_text_from_docx(bio: io.BytesIO) -> str:
    if DocxDocument is None:
        raise RuntimeError("python-docx is not installed.")
    bio.seek(0)
    doc = DocxDocument(bio)
    return "\n".join(p.text for p in doc.paragraphs)


YEARS_RE = re.compile(r"(\d{1,2})\s*\+?\s*(?:years?|yrs?)", re.I)


def estimate_years(text: str) -> float:
    years = 0.0
    for m in YEARS_RE.finditer(text):
        try:
            years = max(years, float(m.group(1)))
        except Exception:
            continue
    return years


def classify_level(years: float) -> str:
    if years >= 8:
        return "Senior"
    if years >= 4:
        return "Mid"
    return "Junior"


def simple_keyword_score(text: str, keywords: List[str]) -> (int, List[str]):
    found = []
    score = 0
    text_low = text.lower()
    for kw in keywords:
        k = kw.strip().lower()
        if not k:
            continue
        if k in text_low:
            score += 10
            found.append(kw)
    return score, found


def score_resume(text: str, keywords: List[str]) -> Dict[str, Any]:
    yrs = estimate_years(text)
    exp_score = int(round(yrs * 5))
    kw_score, found = simple_keyword_score(text, keywords)
    total = exp_score + kw_score
    return {
        "years": yrs,
        "years_source": "regex 'X years' matches",
        "level": classify_level(yrs),
        "exp_score": exp_score,
        "kw_score": kw_score,
        "total": total,
        "keywords_found": ", ".join(found),
    }


# ------------------------------
# UI
# ------------------------------
st.set_page_config(page_title="Resume Scorer (Improved)", layout="wide")
st.title("📄 Resume Scorer — Improved Responsiveness")

tenant_id, client_id, client_secret, site_url = get_config()

with st.sidebar:
    st.markdown("### ⚙️ Run Settings")
    folder_url = st.text_input("Server-relative folder", "/sites/Recruiting/Shared Documents/Active Resumes")
    max_files = st.slider("Max files to scan", 1, 500, 50, 1)
    per_file_timeout = st.slider("Per-file timeout (seconds)", 5, 120, 20, 5)
    enforce_min = st.checkbox("Enforce minimum years requirement", value=False)
    min_years_required = st.number_input("Minimum years (if enforced)", min_value=0.0, max_value=50.0, value=5.0, step=0.5)

st.subheader("Requirements (Optional)")
uploaded_req_file = st.file_uploader("Upload a plain-text requirements file (.txt) — one keyword/phrase per line", type=["txt"])

KEYWORDS: List[str] = []
if uploaded_req_file:
    try:
        req_lines = uploaded_req_file.read().decode("utf-8", errors="ignore").splitlines()
        for line in req_lines:
            line = line.strip()
            if line and not line.endswith(":"):
                KEYWORDS.append(line)
        st.success(f"✅ Loaded {len(KEYWORDS)} keywords from requirements file.")
    except Exception as e:
        st.warning(f"Could not parse requirements file: {e}")
        KEYWORDS = []
else:
    st.info("ℹ️ No requirements uploaded. Scoring will use experience only (0 keyword points).")

run = st.button("▶️ Connect & Scan")

if run:
    data: List[Dict[str, Any]] = []
    try:
        with st.status("Connecting to SharePoint…", expanded=True) as status_box:
            st.write("Obtaining token and testing site reachability.")
            ctx = connect_with_azure_app(tenant_id, client_id, client_secret, site_url)
            status_box.update(label="Connected.", state="complete")

        with st.status("Listing folder files…", expanded=True) as s2:
            folder = get_folder(ctx, folder_url)
            files = list(folder.files)[:max_files]
            s2.update(label=f"Found {len(files)} file(s) (capped to {max_files}).", state="complete")

        progress = st.progress(0)
        status = st.empty()

        for idx, f in enumerate(files, start=1):
            filename = f.properties.get("Name", "unknown")
            status.write(f"🔎 {idx}/{len(files)}: **{filename}**")
            start_ts = time.time()

            try:
                name_l = filename.lower()
                if not (name_l.endswith(".pdf") or name_l.endswith(".docx")):
                    st.caption(f"Skipping {filename} (not PDF/DOCX)")
                    progress.progress(min(idx / len(files), 1.0))
                    continue

                file_url = f.properties["ServerRelativeUrl"]
                bio = download_file(ctx, file_url)

                text = ""
                if name_l.endswith(".pdf"):
                    try:
                        text = extract_text_from_pdf(bio, per_file_timeout)
                    except TimeoutError:
                        st.warning(f"⏱️ Skipped {filename}: exceeded {per_file_timeout}s during PDF parse.")
                        progress.progress(min(idx / len(files), 1.0))
                        continue
                else:
                    try:
                        text = extract_text_from_docx(bio)
                    except Exception as e:
                        st.warning(f"⚠️ Skipped {filename}: DOCX parse error: {e}")
                        progress.progress(min(idx / len(files), 1.0))
                        continue

                if time.time() - start_ts > per_file_timeout:
                    st.warning(f"⏱️ Skipped {filename}: exceeded {per_file_timeout}s overall.")
                    progress.progress(min(idx / len(files), 1.0))
                    continue

                result = score_resume(text, KEYWORDS)

                if enforce_min and result["years"] < float(min_years_required):
                    st.caption(f"Filtered out {filename}: years {result['years']} < {min_years_required}")
                    progress.progress(min(idx / len(files), 1.0))
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

            except ClientRequestException as e:
                st.error(f"SharePoint error on {filename}: {e}")
            except Exception as e:
                st.error(f"Unexpected error on {filename}: {e}")

            progress.progress(min(idx / len(files), 1.0))

        status.write("✅ Done.")
        st.caption(f"Scanned {len(files)} file(s); {len(data)} included after filters.")

        if data:
            df = pd.DataFrame(data)
            df.sort_values(by=["Total Score", "Keyword Score", "Experience Score"], ascending=[False, False, False], inplace=True)
            st.subheader("Results")
            st.dataframe(df, use_container_width=True)

            # Export buttons
            csv_bytes = df.to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Download CSV", data=csv_bytes, file_name=f"resume_scores_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", mime="text/csv")

            try:
                bio_xlsx = BytesIO()
            except NameError:
                bio_xlsx = io.BytesIO()
            with pd.ExcelWriter(bio_xlsx, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Scores")
            bio_xlsx.seek(0)
            st.download_button("⬇️ Download Excel", data=bio_xlsx.getvalue(), file_name=f"resume_scores_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("No results to display. Try raising the file cap or disabling the minimum years filter.")

    except Exception as e:
        st.error(f"❌ Top-level failure: {e}")
        st.stop()
else:
    st.caption("Ready when you are—set your folder and click ▶️ Connect & Scan.")
