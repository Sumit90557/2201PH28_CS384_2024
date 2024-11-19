import streamlit as st
import pandas as pd
import base64

# Set up custom CSS for styling
st.markdown("""
    <style>
    .main-title {
        font-size: 36px;
        color: #333399;
        font-weight: bold;
        text-align: center;
        margin-bottom: 20px;
    }
    .section-title {
        font-size: 24px;
        color: #003366;
        font-weight: bold;
        margin-top: 20px;
        border-bottom: 2px solid #333399;
        padding-bottom: 5px;
    }
    .dataframe, .summary-table {
        margin-top: 10px;
        background-color: #f5f5f5;
        border: 1px solid #e1e1e1;
        border-radius: 5px;
        padding: 10px;
    }
    .stButton > button {
        background-color: #4CAF50;
        color: white;
        border: none;
        padding: 10px 20px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        font-size: 16px;
        margin: 4px 2px;
        border-radius: 4px;
        cursor: pointer;
        transition-duration: 0.4s;
    }
    .stButton > button:hover {
        background-color: #45a049;
    }
    </style>
""", unsafe_allow_html=True)

# Display title and instructions
st.markdown("<div class='main-title'>Grade Processing App</div>", unsafe_allow_html=True)
st.write("Upload an Excel file to process grades and generate a summary with detailed grading statistics.")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

def calculate_total_marks(df, marks_columns, max_marks, weightage):
    def calculate(row):
        total = 0
        for i, col in enumerate(marks_columns, start=2):
            if pd.notnull(row[col]) and pd.notnull(max_marks[i]) and pd.notnull(weightage[i]):
                scaled_mark = (row[col] / max_marks[i]) * weightage[i]
                total += scaled_mark
        return total
    return df.apply(calculate, axis=1)

def assign_grades(df, total_students, schema):
    grade_boundaries = {}
    boundary = 0
    for grade, percentage in schema.items():
        count_for_grade = (percentage / 100) * total_students
        grade_boundaries[grade] = (boundary, boundary + count_for_grade)
        boundary += count_for_grade

    df_sorted = df.sort_values(by='total scaled/100', ascending=False).reset_index(drop=True)
    grade_labels = []
    for i, row in df_sorted.iterrows():
        for grade, (lower, upper) in grade_boundaries.items():
            if lower <= i < upper:
                grade_labels.append(grade)
                break
    df_sorted['Grade'] = grade_labels
    return df_sorted

def generate_summary(df, schema, all_grades):
    total_students = len(df)
    summary_data = []
    for grade in all_grades:
        percentage = schema.get(grade, 0)
        counts = (percentage / 100) * total_students
        rounded_counts = round(counts)
        verified_counts = df['Grade'].value_counts().get(grade, 0)
        summary_data.append([grade, percentage, counts, rounded_counts, verified_counts])
    summary_df = pd.DataFrame(summary_data, columns=['Grade', 'Old IAPC Reco', 'Counts', 'Round', 'Count Verified'])
    summary_df.loc[-1] = ['Total Students', total_students, '', '', '']
    summary_df.index = summary_df.index + 1
    summary_df = summary_df.sort_index()
    return summary_df

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=None)
    column_names = df.iloc[0].values
    df.columns = column_names
    max_marks = df.iloc[1].values
    weightage = df.iloc[2].values
    marks_columns = df.columns[2:]

    # Calculate total marks
    df.loc[3:, 'total scaled/100'] = calculate_total_marks(df.loc[3:], marks_columns, max_marks, weightage)

    # Assign grades
    default_schema = {'AA': 5, 'AB': 15, 'BB': 25, 'BC': 30, 'CC': 15, 'CD': 5, 'DD': 5}
    all_grades = ['AA', 'AB', 'BB', 'BC', 'CC', 'CD', 'DD', 'F', 'I', 'PP', 'NP']
    df_with_grades = assign_grades(df.loc[3:], len(df.loc[3:]), default_schema)

    # Generate summary table
    summary_df = generate_summary(df_with_grades, default_schema, all_grades)

    # Display the processed data and summary
    st.markdown("<div class='section-title'>Processed Data with Grades</div>", unsafe_allow_html=True)
    st.dataframe(df_with_grades.style.set_table_attributes("class='dataframe'"))

    st.markdown("<div class='section-title'>Grade Summary Table</div>", unsafe_allow_html=True)
    st.dataframe(summary_df.style.set_table_attributes("class='summary-table'"))

    # Download processed file
    df_with_summary = pd.concat([df_with_grades, summary_df], axis=1)
    sorted_df = df_with_summary.sort_values(by='Roll')

    @st.cache_data
    def convert_df_to_excel(df):
        return df.to_excel(index=False, engine='openpyxl')

    processed_file = convert_df_to_excel(df_with_summary)
    sorted_file = convert_df_to_excel(sorted_df)

    st.download_button("Download Processed Data", processed_file, "output.xlsx")
