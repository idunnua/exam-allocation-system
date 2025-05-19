# Automated Exam Slot & Staff Allocation System (AESSAS)

This project processes a university's raw exam timetable, expands overlapping venue data, and assigns invigilators based on academic rank and constraints.

## Features
- Converts raw, unstructured timetables into structured formats.
- Handles multi-venue conflicts with intelligent filtering.
- Assigns staff based on fairness and usage constraints.
- Outputs Excel/CSV files ready for administrative use.

## Technologies Used
- Python 3.11+
- Pandas 
- Regex 
- Excel (via openpyxl)
- python-docx

## Folder Structure
exam-allocation-system/
│
├── README.md
├── requirements.txt
├── main.py
├── modules/
│   ├── parser.py
│   ├── venue_splitter.py
│   ├── assign_staff.py
│
├── data/
│   ├── 2024_2025 SECOND SEMESTER WRITTEN EXAM TT_FINAL2.xlsx
│   ├── generatedSlotsIDUNNU.xlsx
│   ├── TIMTEC Roster_Schedule Second semester 2024_2025.docx
│
├── output/
│   ├── Cleaned_Timetable.xlsx
│   ├── Expanded_Timetable.xlsx
│   ├── Assigned_Staff_Slots.xlsx

## How To Run
1. Clone the repo
git clone https://github.com/idunnua/exam-allocation-system
cd exam-allocation-system
2. Install requirements
	pip install -r requirements.txt
3. Put your data in /data
4. Run the full automation:
python main.py

## Sample Outputs
- Cleaned_Timetable.xlsx → Basic time slot mapping
- Expanded_Timetable.xlsx → No venue conflicts
- Assigned_Staff_Slots.xlsx → Fair and complete SP assignment

## Author
Idunnuoluwa Adebambo
https://www.linkedin.com/in/idunnua
