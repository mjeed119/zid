import sys, json
from collections import defaultdict

try:
    import pandas as pd
except ImportError:
    print("يجب تثبيت pandas: pip install pandas openpyxl")
    sys.exit(1)

COLUMN_MAP = {
    "Class": "Class",
    "StudentID": "StudentID",
    "StudentName": "StudentName",
    "Homework": "Homework",
    "HomeworkOutOf": "HomeworkOutOf",
    "Research": "Research",
    "ResearchOutOf": "ResearchOutOf",
    "ClassActivity": "ClassActivity",
    "ClassActivityOutOf": "ClassActivityOutOf",
    "Absence": "Absence",
    "AbsenceOutOf": "AbsenceOutOf",
    "Theoretical": "Theoretical",
    "TheoreticalOutOf": "TheoreticalOutOf",
    "Practical": "Practical",
    "PracticalOutOf": "PracticalOutOf",
}

NUMERIC_COLS = [
    "Homework","HomeworkOutOf","Research","ResearchOutOf","ClassActivity","ClassActivityOutOf",
    "Absence","AbsenceOutOf","Theoretical","TheoreticalOutOf","Practical","PracticalOutOf"
]

def main(xlsx_path, output_json_path):
    df = pd.read_excel(xlsx_path, dtype=str).fillna("")
    for col in NUMERIC_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    classes = defaultdict(list)
    for _, row in df.iterrows():
        klass = str(row.get(COLUMN_MAP["Class"], "")).strip()
        sid = str(row.get(COLUMN_MAP["StudentID"], "")).strip()
        name = str(row.get(COLUMN_MAP["StudentName"], "")).strip()

        grades = {
            "Homework": float(row.get(COLUMN_MAP["Homework"], 0) or 0),
            "Research": float(row.get(COLUMN_MAP["Research"], 0) or 0),
            "ClassActivity": float(row.get(COLUMN_MAP["ClassActivity"], 0) or 0),
            "Absence": float(row.get(COLUMN_MAP["Absence"], 0) or 0),
            "Theoretical": float(row.get(COLUMN_MAP["Theoretical"], 0) or 0),
            "Practical": float(row.get(COLUMN_MAP["Practical"], 0) or 0),
        }
        maxvals = {
            "Homework": float(row.get(COLUMN_MAP["HomeworkOutOf"], 0) or 0),
            "Research": float(row.get(COLUMN_MAP["ResearchOutOf"], 0) or 0),
            "ClassActivity": float(row.get(COLUMN_MAP["ClassActivityOutOf"], 0) or 0),
            "Absence": float(row.get(COLUMN_MAP["AbsenceOutOf"], 0) or 0),
            "Theoretical": float(row.get(COLUMN_MAP["TheoreticalOutOf"], 0) or 0),
            "Practical": float(row.get(COLUMN_MAP["PracticalOutOf"], 0) or 0),
        }
        if klass and sid and name:
            classes[klass].append({
                "id": sid, "name": name,
                "grades": grades, "max": maxvals
            })

    data = {"classes": {k: v for k, v in classes.items()}}
    with open(output_json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"تم إنشاء الملف: {output_json_path}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("طريقة الاستخدام:\\npython3 xlsx_to_json.py grades.xlsx data/students.json")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
