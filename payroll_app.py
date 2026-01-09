import os
import subprocess
import streamlit as st
from pathlib import Path

st.set_page_config(page_title="Payroll Builder", layout="wide")

st.title("💼 Payroll & Attendance Builder")
st.caption("Upload files (CSV/XLS/XLSX) → choose output folder → generate payroll + debug outputs")

# ---------- Helpers ----------
def save_upload(uploaded_file, out_dir: str) -> str:
    out_path = os.path.join(out_dir, uploaded_file.name)
    with open(out_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return out_path

def ensure_dir(path: str, create_if_missing: bool) -> bool:
    p = Path(path).expanduser()
    if p.exists() and p.is_dir():
        return True
    if create_if_missing:
        p.mkdir(parents=True, exist_ok=True)
        return True
    return False

# ---------- UI ----------
with st.form("inputs"):
    col1, col2 = st.columns(2)

    allowed_types = ["csv", "xls", "xlsx"]

    with col1:
        toptracker = st.file_uploader("TopTracker file", type=allowed_types)
        resource = st.file_uploader("Resource Availability file", type=allowed_types)

    with col2:
        leave = st.file_uploader("Leave View file", type=allowed_types)
        attendance = st.file_uploader("Biometric monthinout file", type=allowed_types)

    st.divider()

    default_out = os.path.expanduser("~/Documents/Salary Calculation/PPS/Output")
    output_dir = st.text_input("Output directory (local path on this machine)", value=default_out)

    c1, c2, c3 = st.columns([1, 1, 2])
    with c1:
        create_folder = st.checkbox("Create folder if missing", value=True)
    with c2:
        save_uploads_into_output = st.checkbox("Save uploaded files into output folder", value=True)

    run_btn = st.form_submit_button("🚀 Generate Payroll")

# ---------- Run ----------
if run_btn:
    missing = [name for name, f in {
        "TopTracker": toptracker,
        "Resource Availability": resource,
        "Leave View": leave,
        "Biometric": attendance,
    }.items() if f is None]

    if missing:
        st.error("Missing uploads: " + ", ".join(missing))
        st.stop()

    if not ensure_dir(output_dir, create_folder):
        st.error("Output folder does not exist. Enable 'Create folder if missing' or provide an existing folder.")
        st.stop()

    output_dir = str(Path(output_dir).expanduser())

    # where to store uploads before running script
    uploads_dir = output_dir if save_uploads_into_output else str(Path(output_dir) / "_uploads")
    os.makedirs(uploads_dir, exist_ok=True)

    tt_path = save_upload(toptracker, uploads_dir)
    res_path = save_upload(resource, uploads_dir)
    lv_path = save_upload(leave, uploads_dir)
    att_path = save_upload(attendance, uploads_dir)

    script_path = os.path.abspath("build_master_workbook_v6.py")
    if not os.path.exists(script_path):
        st.error(f"Payroll script not found: {script_path}")
        st.stop()

    cmd = [
        "python", script_path,
        "--toptracker", tt_path,
        "--resource", res_path,
        "--leave", lv_path,
        "--attendance", att_path,
        "--output_dir", output_dir,
    ]

    with st.spinner("Running payroll engine..."):
        proc = subprocess.run(cmd, capture_output=True, text=True)

    if proc.returncode != 0:
        st.error("❌ Failed")
        st.code(proc.stderr)
        st.stop()

    st.success("✅ Payroll generated successfully!")

    # Downloads
    out_csv = os.path.join(output_dir, "monthly_payroll_summary.csv")

    c1, c2, c3 = st.columns(3)
    with c1:
        if os.path.exists(out_csv):
            st.download_button(
                "⬇️ Download monthly_payroll_summary.csv",
                data=open(out_csv, "rb"),
                file_name="monthly_payroll_summary.csv",
                mime="text/csv"
            )

    with c2:
        st.write("Output folder:")
        st.code(output_dir)

    with c3:
        st.write("Quick links (files created)")
        created = [
            "monthly_payroll_summary.csv",
            "_debug_working_calendar.csv",
            "_debug_toptracker_monthly.csv",
            "_debug_leave_monthly.csv",
            "_debug_biometric_daily.csv",
            "_debug_biometric_monthly.csv",
            "_debug_name_mismatch.csv",
            "_name_aliases.csv",
        ]
        existing = [f for f in created if os.path.exists(os.path.join(output_dir, f))]
        st.write("\n".join(existing) if existing else "No files found")

    with st.expander("Logs"):
        st.text(proc.stdout)
