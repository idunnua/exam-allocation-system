# parser.py
import pandas as pd
from docx import Document

#load the document
doc_path = r"TIMTEC Roster_Schedule Second semester 2024_2025.docx"
doc = Document(doc_path)

#match invigilators to their SP numbers
sp_numbers = {
    "Dr. A.J. Olanipekun": "SP1432",
    "Mr. T.K. Adebowale": "SP2701",
    "Dr. S.A Oshadare": "SP2178",
    "Dr. O.T. Ojo": "SP1929",
    "Dr. D.O. Adams": "SP2286",
    "Mr John Ebenezer": "SP2768",
    "Dr. M.B. Adekola": "SP2469",
    "Dr (Mrs) A. A. Akintunde": "SP1138",
    "Dr Mrs. T.O. Kehinde": "SP1598",
    "Dr. O.E. Eteng": "SP2480",
    "Dr. Adetoun. A. Adekitan": "SP2184",
    "Dr. C.P. Njoku": "SP2172",
    "Dr. A.A. Adeyanju": "SP2284",
    "Mr. P.A.S. Soremi": "SP2460",
    "Prof. S.A. Olurode": "SP1308",
    "Dr. Mrs. K.O. Ogunjinmi": "SP2821",
    "Dr. P.O. Omotainse": "SP2465",
    "Dr. O.O. Fawibe": "SP2091",
    "Dr O.A. Makinde": "SP2476",
    "Dr. I. A. Kukoyi": "SP2288"
}

#find the relevant table in the document-with more than 5 rows
table = None
for tbl in doc.tables:
    if len(tbl.rows) > 5:
        table = tbl
        break

#define the static info
days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
weeks = ["Week 1", "Week 2"]
times = ["9AM–2PM", "2PM–7PM"]

#extract the data from the table
data = []
for row in table.rows[1:]:  # skip the header row
    venue = row.cells[0].text.strip()
    if not venue:
        continue
    slots = [cell.text.strip() for cell in row.cells[1:5]]
    for week_index, week in enumerate(weeks):
        for time_index, time in enumerate(times):
            invigilator = slots[week_index * 2 + time_index]
            if not invigilator:
                continue
            for day in days:
                data.append({
                    "Day": day,
                    "Week": week,
                    "Time": time,
                    "Venue": venue,
                    "Invigilator": invigilator,
                    "SP Number": sp_numbers.get(invigilator, "N/A")
                })

#create a DataFrame
df_flat = pd.DataFrame(data)
df_grouped = df_flat.sort_values(by=["Venue", "Day", "Time"])

#save the DataFrame to Excel and CSV files
df_flat.to_excel("timetable_flat.xlsx", index=False)
df_flat.to_csv("timetable_flat.csv", index=False)
df_grouped.to_excel("timetable_grouped.xlsx", index=False)
df_grouped.to_csv("timetable_grouped.csv", index=False)