import streamlit as st
import pandas as pd
import os
import glob
from math import isnan

st.set_page_config(page_title="School Results Portal", layout="centered")

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
                st.write(f"**Name:** {result.get('Name', '-')}")
                st.write(f"**Father Name:** {result.get('Father Name', '-')}")
                st.write(f"**Roll No:** {result.get('Roll Number', '-')}")
                st.write(f"**Admission No:** {result.get('Admission Number', '-')}")
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

                # Printable DMC Section
                st.markdown("""
                    <style>
                    @media print {
                        @page {
                            size: A4;
                            margin: 1cm;
                        }
                        /* Hide everything and collapse its space */
                        body * {
                            visibility: hidden;
                            height: 0 !important;
                            margin: 0 !important;
                            padding: 0 !important;
                        }
                        /* Show only the DMC card and its contents */
                        .print-only, .print-only * {
                            visibility: visible;
                            height: auto !important;
                        }
                        .print-only {
                            display: block !important;
                            position: absolute;
                            left: 0;
                            top: 0;
                            width: 100%;
                            border: 2px solid black;
                            padding: 20px !important;
                            margin: 0 !important;
                        }
                        /* Reset some basic print styles */
                        body { background: white !important; }
                        .dmc-card { font-family: 'Times New Roman', serif; color: black !important; background: white !important; }
                        .dmc-header { text-align: center; border-bottom: 2px solid black; margin-bottom: 20px; }
                        .dmc-header h1, .dmc-header h2, .dmc-header h3 { margin: 5px 0 !important; }
                        .dmc-table { width: 100%; border-collapse: collapse; margin-top: 20px; }
                        .dmc-table th, .dmc-table td { border: 1px solid black; padding: 8px !important; text-align: left; color: black !important; }
                        .dmc-footer { margin-top: 50px; display: flex; justify-content: space-between; }
                    }
                    .print-only { display: none; }
                    </style>
                """, unsafe_allow_html=True)


                # Generate Subject Rows for Print
                subject_rows = ""
                for s in subjects_data:
                    subject_rows += f"<tr><td>{s['Subject']}</td><td>{s['Marks']}</td></tr>"

                dmc_html = f"""
                <div class="print-only dmc-card">
                    <div class="dmc-header">
                        <h1>GHSS Adina</h1>
                        <h2>Annual Examination Result 2025</h2>
                        <h3>Detailed Marks Certificate</h3>
                    </div>
                    <div style="display: flex; justify-content: space-between;">
                        <div>
                            <p><strong>Name:</strong> {result.get('Name', '-')}</p>
                            <p><strong>Father Name:</strong> {result.get('Father Name', '-')}</p>
                            <p><strong>Roll Number:</strong> {result.get('Roll Number', '-')}</p>
                        </div>
                        <div>
                            <p><strong>Class:</strong> {selected_class} {selected_section}</p>
                            <p><strong>Admission No:</strong> {result.get('Admission Number', '-')}</p>
                            <p><strong>Date of Birth:</strong> {result.get('Date of Birth', '-')}</p>
                        </div>
                    </div>
                    <table class="dmc-table">
                        <thead>
                            <tr><th>Subject</th><th>Marks Obtained</th></tr>
                        </thead>
                        <tbody>
                            {subject_rows}
                        </tbody>
                    </table>
                    <div style="margin-top: 20px; border: 1px solid black; padding: 10px;">
                        <p><strong>Total Marks:</strong> {result.get('Obtained Marks', '-')}</p>
                        <p><strong>Percentage:</strong> {pct_str}</p>
                        <p><strong>Remarks:</strong> {result.get('Remarks', '-')}</p>
                        <p><strong>Position:</strong> {result.get('Position', '-')}</p>
                    </div>
                    <div class="dmc-footer">
                        <div style="text-align: center;"><p>_______________________</p><p>Principal Signature</p></div>
                        <div style="text-align: center;"><p>_______________________</p><p>Controller of Exam</p></div>
                    </div>
                </div>
                """
                st.markdown(dmc_html, unsafe_allow_html=True)
                
                st.components.v1.html("""
                    <button onclick="window.parent.print()" style="
                        background-color: #4f46e5;
                        border: none;
                        color: white;
                        padding: 12px 24px;
                        text-align: center;
                        text-decoration: none;
                        display: inline-block;
                        font-size: 16px;
                        margin: 4px 2px;
                        cursor: pointer;
                        border-radius: 8px;
                        width: 100%;
                        font-family: 'Inter', sans-serif;
                        font-weight: 600;
                    ">Print DMC Certificate</button>
                """, height=70)




