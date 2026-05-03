from flask import Flask, render_template, request, jsonify
import pandas as pd
import os
import glob
from math import isnan

app = Flask(__name__)
DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')

def get_data_files():
    files = glob.glob(os.path.join(DATA_DIR, '*.xlsx'))
    # Extract class and section from filename (e.g., "6th A.xlsx" -> Class: "6th", Section: "A")
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
            
    # Sort sections for consistency
    for cls in options:
        options[cls].sort()
    return options

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/options')
def options():
    return jsonify(get_data_files())

@app.route('/api/result', methods=['POST'])
def get_result():
    data = request.json
    cls = data.get('class')
    sec = data.get('section')
    roll_no = data.get('rollNo')

    if not all([cls, sec, roll_no]):
        return jsonify({'error': 'Please provide Class, Section, and Roll No.'}), 400

    filename = f"{cls} {sec}.xlsx"
    filepath = os.path.join(DATA_DIR, filename)

    if not os.path.exists(filepath):
        return jsonify({'error': 'Data file not found for the selected class and section.'}), 404

    try:
        roll_no = int(roll_no)
        # Read excel, header is at row 1 (0-indexed)
        df = pd.read_excel(filepath, header=1)
        
        # Clean up column names by stripping whitespace
        df.columns = df.columns.str.strip()

        # Dynamically find the Roll Number column
        roll_col = None
        for col in ['Roll Number', 'Exam Roll Number', 'Exam Roll No', 'Roll No']:
            if col in df.columns:
                roll_col = col
                break
                
        if not roll_col:
             return jsonify({'error': 'Could not identify the Roll Number column in this file.'}), 500

        # Find the student
        student = df[df[roll_col] == roll_no]
        if student.empty:
             return jsonify({'error': 'Result not found for the given Roll No.'}), 404

        # Convert to dictionary and handle NaNs/dates
        student_dict = student.iloc[0].to_dict()
        
        # Normalize Roll Number key for frontend
        if roll_col != 'Roll Number':
            student_dict['Roll Number'] = student_dict.pop(roll_col)
            
        # Normalize Obtained Marks key (handle case issues like 'obtained Marks')
        for k in list(student_dict.keys()):
            if k.lower() == 'obtained marks' and k != 'Obtained Marks':
                student_dict['Obtained Marks'] = student_dict.pop(k)
        
        # Format dates and handle NaNs
        clean_dict = {}
        for k, v in student_dict.items():
            if pd.isna(v):
                clean_dict[k] = ""
            elif hasattr(v, 'strftime'):
                clean_dict[k] = v.strftime('%d/%m/%Y')
            else:
                clean_dict[k] = v

        # Fix Remarks and Position logic for files where Position is in Remarks
        remarks_val = str(clean_dict.get('Remarks', '')).strip()
        remarks_lower = remarks_val.lower()
        
        position_keywords = ['first', 'second', 'third', '1st', '2nd', '3rd', '4th', '5th', '6th', '7th', '8th', '9th', '10th']
        
        if any(remarks_lower == pw or remarks_lower.startswith(pw + ' ') for pw in position_keywords):
            # The remark is actually a position!
            if not clean_dict.get('Position'):
                clean_dict['Position'] = remarks_val
            clean_dict['Remarks'] = 'Pass'
        elif remarks_lower == 'pass' or remarks_lower == 'promoted':
            clean_dict['Remarks'] = 'Pass'

        return jsonify({'success': True, 'data': clean_dict})

    except ValueError:
         return jsonify({'error': 'Invalid Roll No format.'}), 400
    except Exception as e:
        return jsonify({'error': f"An error occurred while reading the file: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
