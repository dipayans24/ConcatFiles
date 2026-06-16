import re
import tempfile
import streamlit as st
import pandas as pd, os

_INVALID_CHARS = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]')

def extractValidText(text):
    return _INVALID_CHARS.sub('', str(text))

def save_upload(uploaded_file):
    tmp_dir = tempfile.mkdtemp()
    tmp_path = os.path.join(tmp_dir, uploaded_file.name)
    with open(tmp_path, "wb") as f:
        f.write(uploaded_file.read())
    return tmp_path

def custom_function(df):
    #df = df[df["Amount"] <= 945]
    # df = df[~(df["Tags"].astype(str).str.contains("l1")) | (df["Tags"].astype(str).isna())]
    # df["Phone Number"] = df["Phone Number"].astype(str).str.replace(r"\D", "", regex=True)
    # df["CreatedAt"] = pd.to_datetime(df["CreatedAt"], format = "%Y-%m-%d %H:%M:%S", exact=True, dayfirst=True, yearfirst=False)
    # df.sort_values(by="CreatedAt", ascending=False, inplace=True, ignore_index=True)
    return df 


# ── exec() into a local namespace ──────────────────────────────
def runtimeFunction(source: str, fn_name : str):
    """Parse a full def block and return the callable."""
    ns = {}
    exec(source, ns)
    return ns[fn_name]


def concatFiles(folderPath, separateSheets):
    files = [f for f in folderPath  if f.endswith(".csv")]

    file_paths = [f for f in folderPath  if f.endswith(".xlsx")]

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.close()

    writer = pd.ExcelWriter(tmp.name, engine="openpyxl")

    data = pd.DataFrame()
    if len(files) > 0:
        for file in folderPath:
            if file.endswith(".csv"):
                files.append(file)
                df = pd.read_csv(file , low_memory=False)
                FileName = extractValidText(os.path.basename(file).split(".")[0])

                if not separateSheets:
                    df["FileName"] = FileName
                    data = pd.concat([data, df], axis="rows")
                else:
                    sheet_name = extractValidText(FileName)
                    try:
                        df = custom_function(df)
                        df.to_excel(writer, index=False, sheet_name=sheet_name)
                    except Exception as e:
                        df.to_excel(writer, index=False)
                
        if not separateSheets:
            data = custom_function(data)
            data.to_excel(writer, index=False)
    

    if len(file_paths) > 0:
        all_dfs = {}
        for file in file_paths:
            with pd.ExcelFile(file) as xf:
                for sheet in xf.sheet_names:
                    all_dfs.update({sheet:xf.parse(sheet)})  # reuses the open file handle

        if not separateSheets:
            df = pd.concat(all_dfs, axis=0, ignore_index=True)
            df = custom_function(df)

            data = pd.concat([data, df], axis=0, ignore_index=True)
            data.to_excel(writer, index=False)
        else:
            for sheetName, _dfs in all_dfs.items():   
                _dfs = custom_function(_dfs) 
                _dfs.to_excel(writer, index=False, sheet_name = sheetName)

    writer.close()

    return tmp.name, files[0]

      
@st.dialog("Error!!")
def raiseError():
    st.text("Incorrect Folder Path.")

st.header("Files Merger for CSV and Excel.")

col1, col2 = st.columns(2, width="stretch")

folderPath = st.file_uploader("Upload Files.", type=["csv", "xlsx"], accept_multiple_files=True)  

chkbox = st.checkbox("Create seperate sheets, instead of merging.")
if len(folderPath)>0:
    btn = st.button("Process Files", type="primary")

    folderPath = [save_upload(_) for _ in folderPath]

    if btn:
        if len(folderPath) > 0:
            if not chkbox:
                with st.spinner("Processing", show_time=True) as spinner:
                    tmp_path, outputfileName = concatFiles(folderPath, chkbox)

            elif chkbox:
                with st.spinner("Processing", show_time=True) as spinner:
                    tmp_path, outputfileName = concatFiles(folderPath, chkbox) 
            
            filename = os.path.basename(outputfileName).rsplit(".", maxsplit=1)[0]
            outputfileName  = filename[:10]+"_merged.xlsx"

            with open(tmp_path, "rb") as f:
                st.download_button("Download Files", f,  file_name=outputfileName,)

            os.unlink(tmp_path) 
        else:
            raiseError()
