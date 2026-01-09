"""
Student Marks & Performance Analysis (Final Version)

Features:
- Subject-wise analysis
- Student-wise analysis
- User-selected visualizations
- Grade classification
- At-risk student detection
- Automatic insights
- Excel report generation
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

# -------------------
# CONFIGURATION
# -------------------
INPUT_CSV = "students_marks.csv"
OUTPUT_DIR = "analysis_output"
EXCEL_OUT = os.path.join(OUTPUT_DIR, "summary.xlsx")

os.makedirs(OUTPUT_DIR, exist_ok=True)

# -------------------
# LOAD DATA
# -------------------
def load_data(path=INPUT_CSV):
    return pd.read_csv(path, encoding="utf-8")

# -------------------
# IDENTIFY SUBJECT COLUMNS
# -------------------
def get_subject_columns(df):
    non_subjects = {'StudentID', 'Name', 'Class', 'Section'}
    return [c for c in df.columns if c not in non_subjects]

# -------------------
# DATA CLEANING
# -------------------
def clean_data(df, subjects):
    df = df.copy()

    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].astype(str).str.strip()

    for s in subjects:
        df[s] = pd.to_numeric(df[s], errors='coerce')
        df[s] = df[s].fillna(df[s].mean())

    return df

# -------------------
# STUDENT METRICS
# -------------------
def student_metrics(df, subjects):
    df = df.copy()
    df['Total'] = df[subjects].sum(axis=1)
    df['Average'] = df[subjects].mean(axis=1)
    df['Rank'] = df['Total'].rank(method='min', ascending=False).astype(int)
    return df

# -------------------
# SUBJECT METRICS
# -------------------
def subject_metrics(df, subjects):
    stats = {}
    for s in subjects:
        arr = df[s]
        stats[s] = {
            'Mean': arr.mean(),
            'Std': arr.std(ddof=0),
            'Min': arr.min(),
            'Max': arr.max(),
            'Skew': arr.skew()
        }
    return pd.DataFrame(stats).T

# -------------------
# GRADE ASSIGNMENT
# -------------------
def assign_grade(avg):
    if avg >= 90:
        return "A"
    elif avg >= 75:
        return "B"
    elif avg >= 60:
        return "C"
    else:
        return "D"

# -------------------
# STUDENT-WISE ANALYSIS
# -------------------
def student_consistency(df, subjects):
    df['Consistency'] = df[subjects].std(axis=1)
    return df

def strongest_weakest_subject(df, subjects):
    df['Strongest_Subject'] = df[subjects].idxmax(axis=1)
    df['Weakest_Subject'] = df[subjects].idxmin(axis=1)
    return df

def get_top_bottom_students(df, top_n=5):
    top = df.sort_values(by='Average', ascending=False).head(top_n)
    bottom = df.sort_values(by='Average').head(top_n)
    return top, bottom

# -------------------
# USER CHOICE FOR VISUALIZATION
# -------------------
def ask_plot_choice():
    print("\nChoose visualization:")
    print("1. Histogram (Marks distribution)")
    print("2. Bar chart (Average per subject)")
    print("3. Box plot (Outliers & spread)")
    print("4. All plots")
    return input("Enter choice (1–4): ").strip()

# -------------------
# PLOTTING FUNCTIONS
# -------------------
def plot_histograms(df, subjects):
    for s in subjects:
        plt.hist(df[s], bins=10, edgecolor='black')
        plt.title(f"Marks Distribution - {s}")
        plt.xlabel("Marks")
        plt.ylabel("Students")
        plt.tight_layout()
        plt.savefig(os.path.join(OUTPUT_DIR, f"hist_{s}.png"))
        plt.show()

def plot_average_bar(subj_df):
    subj_df['Mean'].plot(kind='bar', title="Average Marks by Subject")
    plt.ylabel("Average Marks")
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "avg_by_subject.png"))
    plt.show()

def plot_boxplot(df, subjects):
    df[subjects].boxplot()
    plt.title("Marks Distribution (Box Plot)")
    plt.ylabel("Marks")
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "boxplot.png"))
    plt.show()

# -------------------
# AUTOMATIC INSIGHTS
# -------------------
def generate_insights(subj_df, students_df):
    print("\n--- AUTOMATIC INSIGHTS ---")
    print("Hardest Subject       :", subj_df['Mean'].idxmin())
    print("Easiest Subject       :", subj_df['Mean'].idxmax())
    print("Most Consistent Subj  :", subj_df['Std'].idxmin())
    print("Students at Risk (<50):", len(students_df[students_df['Average'] < 50]))

# -------------------
# EXPORT TO EXCEL
# -------------------
def export_to_excel(students, subjects, top, bottom):
    with pd.ExcelWriter(EXCEL_OUT, engine="openpyxl") as writer:
        students.to_excel(writer, sheet_name="All_Students", index=False)
        subjects.to_excel(writer, sheet_name="Subject_Stats")
        top.to_excel(writer, sheet_name="Top_Students", index=False)
        bottom.to_excel(writer, sheet_name="At_Risk_Students", index=False)

# -------------------
# MAIN PIPELINE
# -------------------
def run_pipeline():
    df = load_data()
    subjects = get_subject_columns(df)
    print("Detected subjects:", subjects)

    df = clean_data(df, subjects)

    students = student_metrics(df, subjects)
    students['Grade'] = students['Average'].apply(assign_grade)

    students = student_consistency(students, subjects)
    students = strongest_weakest_subject(students, subjects)

    subject_stats = subject_metrics(df, subjects)

    top_students, bottom_students = get_top_bottom_students(students)

    choice = ask_plot_choice()
    if choice == "1":
        plot_histograms(df, subjects)
    elif choice == "2":
        plot_average_bar(subject_stats)
    elif choice == "3":
        plot_boxplot(df, subjects)
    else:
        plot_histograms(df, subjects)
        plot_average_bar(subject_stats)
        plot_boxplot(df, subjects)

    generate_insights(subject_stats, students)

    export_to_excel(students, subject_stats, top_students, bottom_students)

    print("\n✅ Analysis Completed Successfully")
    print("📁 Output saved in:", os.path.abspath(OUTPUT_DIR))

# -------------------
# RUN
# -------------------
run_pipeline()
