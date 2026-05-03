import streamlit as st
import pandas as pd
import os
import glob
import base64
from math import isnan

st.set_page_config(page_title="GHSS Adina Results 2025", layout="centered")

st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
/* ── Global Reset ── */
html, body, [class*="css"] {
    font-family: 'Poppins', sans-serif !important;
}

/* ── Animated Gradient Background ── */
[data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #0f0c29, #302b63, #24243e);
    background-size: 400% 400%;
    animation: gradientBG 12s ease infinite;
    min-height: 100vh;
}
@keyframes gradientBG {
    0%   { background-position: 0% 50%; }
    50%  { background-position: 100% 50%; }
    100% { background-position: 0% 50%; }
}

/* ── Hide top Streamlit bar ── */
[data-testid="stHeader"] { background: transparent !important; }
[data-testid="stToolbar"] { display: none !important; }

/* ── Title ── */
h1 {
    font-size: 2.4rem !important;
    font-weight: 800 !important;
    background: linear-gradient(90deg, #a78bfa, #60a5fa, #34d399);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    text-align: center;
    margin-bottom: 4px !important;
    animation: fadeInDown 0.8s ease;
}
@keyframes fadeInDown {
    from { opacity: 0; transform: translateY(-20px); }
    to   { opacity: 1; transform: translateY(0); }
}

p, label, div {
    color: #e2e8f0 !important;
}

/* ── Glass Card for main block ── */
[data-testid="stMain"] .block-container {
    background: rgba(255,255,255,0.06);
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
    border: 1px solid rgba(255,255,255,0.12);
    border-radius: 20px;
    padding: 2rem 2.5rem !important;
    margin-top: 2rem;
    animation: fadeIn 0.6s ease;
    max-width: 860px;
}
@keyframes fadeIn {
    from { opacity: 0; transform: translateY(10px); }
    to   { opacity: 1; transform: translateY(0); }
}

/* ── Selectbox & Input ── */
[data-testid="stSelectbox"] > div, [data-testid="stNumberInput"] > div {
    background: rgba(255,255,255,0.08) !important;
    border-radius: 10px !important;
    border: 1px solid rgba(255,255,255,0.15) !important;
    color: white !important;
}
[data-testid="stSelectbox"] label, [data-testid="stNumberInput"] label {
    color: #a78bfa !important;
    font-weight: 600 !important;
    font-size: 0.85rem !important;
}

/* ── Buttons ── */
[data-testid="stButton"] > button, [data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #7c3aed, #4f46e5) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 0.6rem 1.5rem !important;
    font-weight: 600 !important;
    font-size: 1rem !important;
    transition: transform 0.2s, box-shadow 0.2s !important;
    box-shadow: 0 4px 20px rgba(124,58,237,0.4) !important;
}
[data-testid="stButton"] > button:hover, [data-testid="stDownloadButton"] > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 6px 28px rgba(124,58,237,0.6) !important;
}

/* ── Subheaders ── */
h2, h3 {
    color: #a78bfa !important;
    font-weight: 700 !important;
    border-bottom: 1px solid rgba(167,139,250,0.3);
    padding-bottom: 4px;
    margin-top: 1.2rem !important;
}

/* ── Table ── */
[data-testid="stTable"] {
    border-radius: 12px;
    overflow: hidden;
    border: 1px solid rgba(255,255,255,0.1);
}
[data-testid="stTable"] th {
    background: rgba(124,58,237,0.5) !important;
    color: white !important;
    text-align: center !important;
    font-weight: 700 !important;
}
[data-testid="stTable"] td {
    background: rgba(255,255,255,0.04) !important;
    color: #e2e8f0 !important;
    text-align: center !important;
}

/* ── Metrics ── */
[data-testid="metric-container"] {
    background: rgba(255,255,255,0.07) !important;
    border: 1px solid rgba(255,255,255,0.12) !important;
    border-radius: 12px !important;
    padding: 1rem !important;
    text-align: center;
}
[data-testid="stMetricLabel"] { color: #94a3b8 !important; font-size: 0.8rem !important; }
[data-testid="stMetricValue"] { color: white !important; font-size: 1.4rem !important; font-weight: 700 !important; }

/* ── Success / Error banners ── */
[data-testid="stAlert"] {
    border-radius: 10px !important;
    font-weight: 600 !important;
}

/* ── Caption ── */
[data-testid="stCaptionContainer"] p {
    color: #94a3b8 !important;
    text-align: center;
}
</style>
""", unsafe_allow_html=True)

DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')

@st.cache_data
def get_data_files():
    if not os.path.exists(DATA_DIR):
        return {}
    files = glob.glob(os.path.join(DATA_DIR, '*.xlsx'))
    options = {}
    for f in files:
        basename = os.path.basename(f).replace('.xlsx', '')
        parts = basename.split(' ')
        if len(parts) >= 2:
            cls = parts[0]
            sec = ' '.join(parts[1:])
            if cls not in options:
                options[cls] = []
            options[cls].append(sec)
    for cls in options:
        options[cls].sort()
    return options

def load_result(cls, sec, roll_no):
    filename = f"{cls} {sec}.xlsx"
    filepath = os.path.join(DATA_DIR, filename)

    if not os.path.exists(filepath):
        return None, f"Data file not found for the selected class and section."

    try:
        df = pd.read_excel(filepath, header=1)
        df.columns = df.columns.str.strip()

        roll_col = None
        for col in ['Roll Number', 'Exam Roll Number', 'Exam Roll No', 'Roll No']:
            if col in df.columns:
                roll_col = col
                break
                
        if not roll_col:
             return None, "Could not identify the Roll Number column in this file."

        student = df[df[roll_col] == roll_no]
        if student.empty:
             return None, "Result not found for the given Roll No."

        student_dict = student.iloc[0].to_dict()
        
        if roll_col != 'Roll Number':
            student_dict['Roll Number'] = student_dict.pop(roll_col)
            
        for k in list(student_dict.keys()):
            if k.lower() == 'obtained marks' and k != 'Obtained Marks':
                student_dict['Obtained Marks'] = student_dict.pop(k)

        clean_dict = {}
        for k, v in student_dict.items():
            if pd.isna(v):
                clean_dict[k] = ""
            elif hasattr(v, 'strftime'):
                clean_dict[k] = v.strftime('%d/%m/%Y')
            elif isinstance(v, (int, float)) and v == int(v):
                clean_dict[k] = int(v)
            else:
                clean_dict[k] = v

        remarks_val = str(clean_dict.get('Remarks', '')).strip()
        remarks_lower = remarks_val.lower()
        position_keywords = ['first', 'second', 'third', '1st', '2nd', '3rd', '4th', '5th', '6th', '7th', '8th', '9th', '10th']
        
        if any(remarks_lower == pw or remarks_lower.startswith(pw + ' ') for pw in position_keywords):
            if not clean_dict.get('Position'):
                clean_dict['Position'] = remarks_val
            clean_dict['Remarks'] = 'Pass'
        elif remarks_lower == 'pass' or remarks_lower == 'promoted':
            clean_dict['Remarks'] = 'Pass'

        return clean_dict, None
    except Exception as e:
        return None, f"An error occurred while reading the file: {str(e)}"

st.title("GHSS Adina Annual Results 2025")
st.write("Select your class, section, and enter your roll number to view your result.")

class_options = get_data_files()

if not class_options:
    st.error("No data files found in the 'data' folder.")
else:
    col1, col2, col3 = st.columns(3)
    
    with col1:
        selected_class = st.selectbox("Class", options=list(class_options.keys()))
    
    with col2:
        sections = class_options.get(selected_class, [])
        selected_section = st.selectbox("Section", options=sections)
        
    with col3:
        roll_no = st.number_input("Roll Number", min_value=1, step=1)
        
    if st.button("Check Result", type="primary"):
        with st.spinner("Fetching result..."):
            result, error = load_result(selected_class, selected_section, roll_no)
            
            if error:
                st.error(error)
            elif result:
                st.success("Result found!")
                
                # Student Details Card
                st.subheader("Student Details")
                
                row1_col1, row1_col2, row1_col3 = st.columns(3)
                with row1_col1:
                    st.write(f"**Name:** {result.get('Name', '-')}")
                with row1_col2:
                    st.write(f"**Father Name:** {result.get('Father Name', '-')}")
                with row1_col3:
                    st.write(f"**Roll No:** {result.get('Roll Number', '-')}")
                
                row2_col1, row2_col2 = st.columns(2)
                with row2_col1:
                    st.write(f"**Admission No:** {result.get('Admission Number', '-')}")
                with row2_col2:
                    if result.get('Date of Birth'):
                        st.write(f"**Date of Birth:** {result.get('Date of Birth')}")
                
                # Subjects Table
                st.subheader("Marks Summary")
                non_subject_keys = ['Roll Number', 'Admission Number', 'Name', 'Father Name', 'Date of Birth', 'Obtained Marks', 'Per%', 'Failed Subjects', 'Below 25', 'Remarks', 'Position', 'S.No', 'Unnamed: 0']
                
                subjects_data = []
                for k, v in result.items():
                    if k not in non_subject_keys and v != "":
                        subjects_data.append({"Subject": k, "Marks": v})
                
                if subjects_data:
                    # Center the table by putting it in a middle column to make it less wide
                    _, mid_col, _ = st.columns([1, 2, 1])
                    with mid_col:
                        # Hide the pandas index when displaying the table
                        st.table(pd.DataFrame(subjects_data).set_index('Subject'))
                
                # Result Metrics
                st.subheader("Overall Result")
                m1, m2, m3, m4 = st.columns(4)
                
                is_pass = str(result.get("Remarks", "")).lower() == "pass"

                if is_pass:
                    st.success("Congratulations! You have passed. 🎓")
                else:
                    st.error("Keep trying! Success is just around the corner. 📚")

                with m1:
                    st.metric("Total Marks", result.get("Obtained Marks", "-"))
                with m2:
                    pct = result.get("Per%")
                    if pct:
                        if isinstance(pct, str) and '%' in pct:
                            pct_str = pct
                        else:
                            try:
                                pct_str = f"{float(pct) * 100:.2f}%"
                            except ValueError:
                                pct_str = str(pct)
                    else:
                        pct_str = "-"
                    st.metric("Percentage", pct_str)
                with m3:
                    st.metric("Remarks", result.get("Remarks", "-"))
                with m4:
                    st.metric("Position", result.get("Position", "-"))

                # Generate Subject Rows for Print
                subject_rows = ""
                for s in subjects_data:
                    subject_rows += f"<tr><td>{s['Subject']}</td><td>{s['Marks']}</td></tr>"

                # Build complete standalone DMC HTML page (no auto-print script)
                popup_html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>DMC - {result.get('Name', '')}</title>
<style>
  @page {{ size: A4; margin: 1cm; }}
  body {{ font-family: 'Times New Roman', serif; font-size: 13px; margin: 0; padding: 20px; color: black; background: white; }}
  h1 {{ font-size: 22px; margin: 4px 0; }}
  h2 {{ font-size: 18px; margin: 4px 0; }}
  h3 {{ font-size: 15px; margin: 4px 0; }}
  .header {{ text-align: center; border-bottom: 2px solid black; padding-bottom: 10px; margin-bottom: 15px; }}
  .info-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 6px; margin-bottom: 15px; }}
  .info-grid p {{ margin: 3px 0; }}
  table {{ width: 60%; border-collapse: collapse; margin: 15px auto; }}
  th, td {{ border: 1px solid black; padding: 6px 12px; text-align: center; }}
  th {{ background-color: #f0f0f0; }}
  .summary {{ border: 1px solid black; padding: 10px; margin-top: 15px; display: grid; grid-template-columns: 1fr 1fr; gap: 6px; }}
  .summary p {{ margin: 3px 0; }}
  .footer {{ margin-top: 40px; display: flex; justify-content: space-between; text-align: center; }}
  .footer div {{ flex: 1; }}
  .print-btn {{ display: block; margin: 20px auto; padding: 10px 30px; background: #4f46e5; color: white; border: none; border-radius: 8px; font-size: 15px; cursor: pointer; }}
  @media print {{ .print-btn {{ display: none; }} }}
</style>
</head>
<body>
  <button class="print-btn" onclick="window.print()">Print DMC</button>
  <div class="header">
    <h1>GHSS Adina</h1>
    <h2>Annual Examination Result 2025</h2>
    <h3>Detailed Marks Certificate</h3>
  </div>
  <div class="info-grid">
    <p><strong>Name:</strong> {result.get('Name', '-')}</p>
    <p><strong>Father Name:</strong> {result.get('Father Name', '-')}</p>
    <p><strong>Roll Number:</strong> {result.get('Roll Number', '-')}</p>
    <p><strong>Class:</strong> {selected_class} {selected_section}</p>
    <p><strong>Admission No:</strong> {result.get('Admission Number', '-')}</p>
    <p><strong>Date of Birth:</strong> {result.get('Date of Birth', '-')}</p>
  </div>
  <table>
    <thead><tr><th>Subject</th><th>Marks Obtained</th></tr></thead>
    <tbody>{subject_rows}</tbody>
  </table>
  <div class="summary">
    <p><strong>Total Marks:</strong> {result.get('Obtained Marks', '-')}</p>
    <p><strong>Percentage:</strong> {pct_str}</p>
    <p><strong>Remarks:</strong> {result.get('Remarks', '-')}</p>
    <p><strong>Position:</strong> {result.get('Position', '-')}</p>
  </div>
  <div class="footer">
    <div><p>_______________________</p><p>Principal Signature</p></div>
    <div><p>_______________________</p><p>Controller of Exam</p></div>
  </div>
</body>
</html>"""

                # Store in session state so the print button can access it
                st.session_state['dmc_html'] = popup_html

                # Download DMC as HTML file - most reliable approach in Streamlit
                student_name = result.get('Name', 'student').replace(' ', '_')
                st.download_button(
                    label="🖨️ Download & Print DMC",
                    data=popup_html,
                    file_name=f"DMC_{student_name}.html",
                    mime="text/html",
                    type="primary",
                    help="Download the DMC as an HTML file. Open it in your browser and press Ctrl+P to print."
                )
                st.caption("After downloading, open the file in your browser and press **Ctrl+P** to print.")



