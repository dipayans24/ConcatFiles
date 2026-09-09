import re
import traceback
import tempfile
import streamlit as st
import pandas as pd, os

_INVALID_CHARS = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]')

FUNC_NAME = "custom_function"

# Only this body is ever shown to / edited by the user. The "def ...:" header
# and "return df" footer are fixed and hidden, and get spliced back in
# whenever the code is compiled — see build_function_source().
DEFAULT_BODY = '''df.drop_duplicates(subset=df.columns, inplace=True)
#df = df[~(df["Tags"].astype(str).str.contains("l1")) | (df["Tags"].astype(str).isna())]
#df = df[df["Status"] == "captured"]
#df["Phone Number"] = df["Phone Number"].astype(str).str.replace(r"\\D", "", regex=True)
#df["CreatedAt"] = pd.to_datetime(df["CreatedAt"], format="%Y-%m-%d %H:%M:%S", exact=True, dayfirst=True, yearfirst=False)
#df.sort_values(by="CreatedAt", ascending=False, inplace=True, ignore_index=True)'''


def build_function_source(body: str) -> str:
    """Wraps the user-entered body with the fixed, hidden def/return lines."""
    lines = body.splitlines() or [""]
    indented = "\n".join(("    " + line) if line.strip() else "" for line in lines)
    return f"def {FUNC_NAME}(df):\n{indented}\n    return df\n"


def extractValidText(text):
    return _INVALID_CHARS.sub('', str(text))


def save_upload(uploaded_file):
    tmp_dir = tempfile.mkdtemp()
    tmp_path = os.path.join(tmp_dir, uploaded_file.name)
    with open(tmp_path, "wb") as f:
        f.write(uploaded_file.read())
    return tmp_path


# ── exec() into a local namespace ──────────────────────────────
def runtimeFunction(source: str, fn_name: str):
    """Parse a full def block and return the callable. Raises SyntaxError /
    NameError / any exception the user's code raises at *definition* time."""
    ns = {"pd": pd, "re": re}
    exec(source, ns)
    if fn_name not in ns or not callable(ns[fn_name]):
        raise NameError(
            f"No function named '{fn_name}' was found. "
            f"Please define a function called {fn_name}(df) that returns a dataframe."
        )
    return ns[fn_name]


def load_units(file_paths):
    """Expand uploaded files into logical (name, dataframe) units.
    Each CSV is one unit; each sheet inside an xlsx is one unit."""
    units = []
    for f in file_paths:
        if f.endswith(".csv"):
            df = pd.read_csv(f, low_memory=False)
            name = extractValidText(os.path.basename(f).split(".")[0])
            units.append({"name": name, "df": df, "status": "pending"})
        elif f.endswith(".xlsx"):
            with pd.ExcelFile(f) as xf:
                for sheet in xf.sheet_names:
                    df = xf.parse(sheet)
                    name = extractValidText(sheet)
                    units.append({"name": name, "df": df, "status": "pending"})
    return units


def unique_sheet_name(name, used):
    base = (name or "Sheet")[:31]
    candidate = base
    i = 1
    while candidate in used:
        suffix = f"_{i}"
        candidate = base[: 31 - len(suffix)] + suffix
        i += 1
    used.add(candidate)
    return candidate


ROW_LIMIT_FOR_EXCEL = 1_000_000


def write_output(units, separateSheets):
    """Returns (tmp_path, extension). When merging (not separateSheets) and the
    concatenated data exceeds ROW_LIMIT_FOR_EXCEL rows, writes a CSV
    (comma-separated, no index) instead of an xlsx."""
    if separateSheets:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp.close()
        writer = pd.ExcelWriter(tmp.name, engine="openpyxl")
        used_names = set()
        for unit in units:
            sheet_name = unique_sheet_name(unit["name"], used_names)
            unit["df"].to_excel(writer, index=False, sheet_name=sheet_name)
        writer.close()
        return tmp.name, "xlsx"

    data = pd.DataFrame()
    for unit in units:
        df = unit["df"].copy()
        df["FileName"] = unit["name"]
        data = pd.concat([data, df], axis=0, ignore_index=True)

    if len(data) > ROW_LIMIT_FOR_EXCEL:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
        tmp.close()
        data.to_csv(tmp.name, sep=",", index=False)
        return tmp.name, "csv"
    else:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp.close()
        data.to_excel(tmp.name, index=False)
        return tmp.name, "xlsx"


# ── dialogs ─────────────────────────────────────────────────────
@st.dialog("Error!!")
def raiseError():
    st.text("Incorrect Folder Path.")


@st.dialog("Your code failed to compile")
def code_error_dialog(message):
    st.error(message)
    st.caption("Fix the code in the text box below, or skip it and proceed "
               "with the raw (unprocessed) data for every file.")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Edit code", width="stretch"):
            st.session_state.processing_active = False
            st.session_state.pending_error = None
            st.rerun()
    with col2:
        if st.button("Skip code entirely & continue", type="primary", width="stretch"):
            st.session_state.skip_all_code = True
            st.session_state.pending_error = None
            st.rerun()


@st.dialog("Error while running your code")
def file_error_dialog(unit_name, message):
    st.error(f"Failed on **{unit_name}**:")
    st.code(message)
    st.caption("You can edit your code and re-run, or skip this file's code "
               "and include it in the output unprocessed.")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Edit code", width="stretch", key="edit_file_err"):
            st.session_state.processing_active = False
            st.session_state.pending_error = None
            st.rerun()
    with col2:
        if st.button("Skip this file & continue", type="primary", width="stretch", key="skip_file_err"):
            st.session_state.skip_names.add(unit_name)
            st.session_state.pending_error = None
            st.rerun()


def run_pipeline(separateSheets):
    """Runs (or resumes) processing. Stops and shows a dialog on the first
    unresolved error; otherwise finishes and writes the output file."""
    units = st.session_state.units

    if st.session_state.use_code and not st.session_state.get("skip_all_code"):
        try:
            source = build_function_source(st.session_state.user_body)
            user_fn = runtimeFunction(source, FUNC_NAME)
        except Exception:
            st.session_state.pending_error = {"scope": "code", "message": traceback.format_exc()}
            return
    else:
        user_fn = None

    for unit in units:
        if unit["status"] == "done":
            continue

        if (not st.session_state.use_code
                or st.session_state.get("skip_all_code")
                or unit["name"] in st.session_state.skip_names):
            unit["status"] = "done"
            continue

        try:
            unit["df"] = user_fn(unit["df"])
            unit["status"] = "done"
        except Exception:
            st.session_state.pending_error = {
                "scope": "file",
                "name": unit["name"],
                "message": traceback.format_exc(),
            }
            return

    # All units resolved — write the file.
    st.session_state.output_path, st.session_state.output_ext = write_output(units, separateSheets)
    st.session_state.output_source_name = units[0]["name"] if units else "output"
    st.session_state.processing_active = False


def files_signature(files):
    """A fingerprint of the current uploader contents. Includes Streamlit's
    internal file_id (not just name/size) so that re-uploading a file with the
    same name still counts as a change."""
    return tuple(sorted((f.name, f.size, getattr(f, "file_id", None)) for f in (files or [])))


# ── UI ──────────────────────────────────────────────────────────
st.set_page_config("🧩 Files Merger", layout="wide")
st.header("🧩 Files Merger for CSV and Excel", divider=True, text_alignment="center")

for key, default in [
    ("user_body", DEFAULT_BODY),
    ("use_code", False),
    ("units", None),
    ("skip_names", set()),
    ("skip_all_code", False),
    ("pending_error", None),
    ("processing_active", False),
    ("output_path", None),
    ("output_ext", "xlsx"),
    ("output_source_name", None),
    ("last_upload_signature", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default

col1, col2 = st.columns(2, width="stretch")

folderPath = st.file_uploader("Upload Files.", type=["csv", "xlsx"], accept_multiple_files=True)

# If the uploaded fileset has changed since the last processed run (new file,
# removed file, or a same-name re-upload), the old output/download is stale —
# clear it so we don't keep showing results for files that are no longer here.
current_signature = files_signature(folderPath)
if current_signature != st.session_state.last_upload_signature:
    st.session_state.units = None
    st.session_state.skip_names = set()
    st.session_state.skip_all_code = False
    st.session_state.pending_error = None
    st.session_state.processing_active = False
    st.session_state.output_path = None
    st.session_state.output_ext = "xlsx"
    st.session_state.output_source_name = None
    st.session_state.last_upload_signature = current_signature

chkbox = st.checkbox("Create seperate sheets, instead of merging.")

st.session_state.use_code = st.checkbox(
    "Apply custom processing code to each file before merging", value=st.session_state.use_code
)

if st.session_state.use_code:
    st.caption(
        "This code runs on every uploaded file/sheet's dataframe (`df`) before concatenation. "
        "If it errors out, you'll be able to fix it or skip it."
    )
    with st.container(border=True):
        st.markdown(f"`def {FUNC_NAME}(df):`")
        st.session_state.user_body = st.text_area(
            "Function body",
            value=st.session_state.user_body,
            height=200,
            label_visibility="collapsed",
        )
        st.markdown("`return df`")

if folderPath and len(folderPath) > 0:
    btn = st.button("Process Files", type="primary")

    saved_paths = [save_upload(_) for _ in folderPath]

    if btn:
        # Fresh run — reset all pipeline state.
        st.session_state.units = load_units(saved_paths)
        st.session_state.skip_names = set()
        st.session_state.skip_all_code = False
        st.session_state.pending_error = None
        st.session_state.output_path = None
        st.session_state.processing_active = True

    if st.session_state.processing_active and st.session_state.units:
        with st.spinner("Processing", show_time=True):
            run_pipeline(chkbox)

    if st.session_state.pending_error:
        err = st.session_state.pending_error
        if err["scope"] == "code":
            code_error_dialog(err["message"])
        else:
            file_error_dialog(err["name"], err["message"])

    if st.session_state.output_path:
        filename = extractValidText(st.session_state.output_source_name or "output")[:10]
        ext = st.session_state.output_ext
        outputfileName = f"{filename}_merged.{ext}"
        mime = "text/csv" if ext == "csv" else (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        if ext == "csv":
            st.caption("Output exceeded 1,000,000 rows — exported as CSV instead of Excel.")

        with open(st.session_state.output_path, "rb") as f:
            st.download_button(
                "Download Files", f, file_name=outputfileName, mime=mime, on_click="ignore"
            )
else:
    if folderPath is not None and len(folderPath) == 0:
        pass
