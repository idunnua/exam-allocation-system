import pandas as pd
import re

# Load Excel file
df_raw = pd.read_excel("2024_2025 SECOND SEMESTER WRITTEN EXAM TT_FINAL2.xlsx", header=None)

# Venue mapping dictionary
venue_map = {
    "COLBIOS AUD": 1, "A105": 2, "A110": 3, "A201": 4, "A215": 5, "ANE1": 6, "ANE2": 7,
    "CAD": 8, "CLOTHING LAB": 9, "COLFHEC1": 10, "COLFHEC2": 11, "CPL": 12, "ENG-AUD": 13,
    "HEALTH-CENTRE": 14, "JAO1": 15, "JAO2": 16, "JAO3": 17, "MP01": 18, "MP02": 19,
    "MPL": 20, "MP-BIO LAB": 21, "MP-CHEM LAB": 22, "NTDR-LAB": 23, "RC101": 24,
    "RC102": 25, "RC201": 26, "RC202": 27, "RC203": 28, "R205": 29, "R206": 30,
    "VET 200": 31, "VET 300": 32, "VET 400": 33, "VET 500": 34, "VET 600": 35,
    "VET AUD": 36, "VET LAB": 37, "VET PATH LAB": 38, "VET ANAT LAB": 39,
    "250-SEATER COMPUTER LAB": 40, "500-SEATER COMPUTER LAB": 41, "CVE UPS": 42,
    "CVE DOWN": 43, "COLFHEC LAB": 44, "BAM-400LR": 45, "ETS-400LR": 46,
    "YAKUB_MAMHOOD_HALL(HA)": 47, "VET_PHY_LAB": 48, "VET_PARA_LAB": 49,
    "YAKUB_MAHMOOD_HALL(EXAM)": 50, "VET_PHARM_LAB": 51, "COLVET/VTH": 52,
    "ACCT-400LR": 53, "ECO-400LR": 54, "BFN-400LR": 55, "COLENDS-LR1": 56,
    "COLENDS-LR2": 57, "CENTS-AUD": 58, "PATTERN LAB": 59, "TEXTILE LAB": 60,
    "1000 CAP LT HALL": 61, "TETFUND PHS LAB": 62, "TETFUND CHM LAB": 63,
    "TETFUND BIO LAB": 64, "COLFHEC3": 65, "MP03/04": 66, "VET MICRO LAB": 67,
    "NEW HALL": 68, "AGRIC LAB 1": 69, "AGRIC LAB 2": 70, "GLR I": 71, "GLR II": 72,
    "GLR III": 73, "ABE LR": 74, "MCE LR": 75, "PISAD AUD": 76, "ELBOG I": 77,
    "ELBOG II": 78, "ACAD C1-C3": 79, "ACAD C4-C6": 80, "ACAD A3": 81, "ACAD A5": 82,
    "ACAD B3": 83, "ACAD B5": 84, "CVE STUDIO": 85, "AUD II": 86, "AUD I": 87,
    "PPCP LR": 88, "PBST LR": 89, "SSLM LR": 90, "CPT LR": 91, "HRT LR": 92,
    "AUD III": 93, "VET FARM": 94, "VET THEATRE": 95, "VET CLINIC": 96, "VET THERIO LAB": 97, "GEO_ROOM1": 98, "GEO_ROOM2": 99
}

# Get exam date rows
dates = [(i, df_raw.iloc[i, 0]) for i in range(len(df_raw)) if isinstance(df_raw.iloc[i, 0], str) and "2025" in df_raw.iloc[i, 0]]
records = []

for day_index, (start_row, _) in enumerate(dates):
    exam_day = day_index + 1
    next_row = dates[day_index + 1][0] if day_index + 1 < len(dates) else len(df_raw)
    daily_data = df_raw.iloc[start_row:next_row].reset_index(drop=True)

    for session_index, (course_col, venue_col, start, end, label) in enumerate([
        (1, 2, "9:00am", "12:00 noon", "Morning (9am - 12noon)"),
        (3, 4, "2:00pm", "5:00pm", "Afternoon(2pm - 5pm)")
    ]):
        day_paper_id = f"{exam_day}--{session_index + 1}"
        seen_venues = set()

        for _, row in daily_data.iterrows():
            course_raw = row[course_col] if course_col < len(row) else None
            venue_raw = row[venue_col] if venue_col < len(row) else None

            if pd.isna(course_raw) or pd.isna(venue_raw):
                continue

            course_list = [c.strip() for c in str(course_raw).split(",")]
            venue_list = [v.strip().upper() for v in str(venue_raw).split(",")]

            # Check if any venue has already been used in this session
            if any(v in seen_venues for v in venue_list):
                continue  # Skip this record entirely

            for course in course_list:
                course_code = re.sub(r"\s*\[.*?\]", "", course).strip()

                for venue in venue_list:
                    seen_venues.add(venue)
                    hall_id = venue_map.get(venue, "")
                    hall_name = venue if hall_id else ""

                    records.append({
                        "Day_Paper_ID": day_paper_id,
                        "CourseCode": course_code,
                        "CourseStart": start,
                        "CourseEnd": end,
                        "CoursePeriod": label,
                        "ExamDayId": exam_day,
                        "IsActive": 1,
                        "Semester": 2,
                        "Session": "2024/2025",
                        "Hall": hall_name,
                        "HallID": hall_id
                    })

# Final export
df = pd.DataFrame(records)
df.to_excel("Course_Schedule_With_Venues_Deduplicated.xlsx", index=False)
df.to_csv("Course_Schedule_With_Venues_Deduplicated.csv", index=False)

print("Venues split, deduplicated, and exported successfully.")