import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="Student Progress Report", layout="centered")

st.title("📊 Student Progress Report Generator")

# -----------------------------
# Student Info
# -----------------------------
st.header("👩‍🎓 Student Information")

name = st.text_input("Student Name")
student_id = st.text_input("Student ID")
course = st.text_input("Course / Subject")
semester = st.selectbox("Semester", ["Sem 1", "Sem 2", "Sem 3", "Sem 4"])

# -----------------------------
# Marks Input
# -----------------------------
st.header("📝 Assessment Scores")

data = {
    "Assessment": ["Quiz", "Assignment", "Midterm", "Final"],
    "Score": [0, 0, 0, 0]
}

df = pd.DataFrame(data)

edited_df = st.data_editor(
    df,
    num_rows="fixed",
    use_container_width=True
)

# -----------------------------
# Calculation
# -----------------------------
average = edited_df["Score"].mean()

def grade(avg):
    if avg >= 80:
        return "A"
    elif avg >= 70:
        return "B"
    elif avg >= 60:
        return "C"
    elif avg >= 50:
        return "D"
    else:
        return "F"

final_grade = grade(average)

# -----------------------------
# Results
# -----------------------------
st.header("📈 Performance Summary")

st.metric("Average Score", f"{average:.2f}")
st.metric("Final Grade", final_grade)

# Chart
fig, ax = plt.subplots()
ax.bar(edited_df["Assessment"], edited_df["Score"])
ax.set_ylim(0, 100)
ax.set_ylabel("Score")
ax.set_title("Assessment Breakdown")

st.pyplot(fig)

# -----------------------------
# Report Text
# -----------------------------
st.header("📄 Generated Report")

report = f"""
Student Name: {name}
Student ID: {student_id}
Course: {course}
Semester: {semester}

Assessment Scores:
{edited_df.to_string(index=False)}

Average Score: {average:.2f}
Final Grade: {final_grade}

Remarks:
{"Good progress. Keep it up!" if average >= 70 else "Needs improvement. More practice recommended."}
"""

st.text_area("Progress Report", report, height=300)
