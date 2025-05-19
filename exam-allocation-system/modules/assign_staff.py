import pandas as pd
from collections import defaultdict
import random

file_path = "generatedSlotsIDUNNU.xlsx"
slots_df = pd.read_excel(file_path, sheet_name=0)  # assuming 'Slots' is the first sheet
profs_df = pd.read_excel(file_path, sheet_name="AllPRofs")
non_profs_df = pd.read_excel(file_path, sheet_name="Non Profs")

profs = profs_df['StaffID'].dropna().astype(str).tolist()
non_profs = non_profs_df['StaffID'].dropna().astype(str).tolist()

random.shuffle(profs)
random.shuffle(non_profs)

prof_day_count = defaultdict(set)     # SP -> set of days worked
nonprof_day_count = defaultdict(set)  # SP -> set of days worked
used_today = defaultdict(set)         # day -> SPs used on that day

# Organize slots by day-period to better manage one-appearance-per-day rule
slots_df['StaffID'] = None
grouped = slots_df.groupby(['day', 'Period'])
assigned_rows = []

#assign Profs (max 5 days each)
for (day, period), group in grouped:
    day = int(day)
    for idx in group.index:
        assigned = False
        for sp in profs:
            if day not in prof_day_count[sp] and len(prof_day_count[sp]) < 5 and sp not in used_today[day]:
                slots_df.at[idx, 'StaffID'] = sp
                prof_day_count[sp].add(day)
                used_today[day].add(sp)
                assigned = True
                break
        if not assigned:
            continue  

#assign Non-Profs (at least 5 days)
remaining = slots_df[slots_df['StaffID'].isna()]

for (day, period), group in remaining.groupby(['day', 'Period']):
    day = int(day)
    for idx in group.index:
        for sp in non_profs:
            if sp not in used_today[day]:
                slots_df.at[idx, 'StaffID'] = sp
                nonprof_day_count[sp].add(day)
                used_today[day].add(sp)
                break

# Re-check for unassigned slots
unassigned = slots_df[slots_df['StaffID'].isna()]
if not unassigned.empty:
    print(f"Warning: {len(unassigned)} unassigned slots. Trying to reassign...")

    for idx in unassigned.index:
        day = int(slots_df.at[idx, 'day'])
        for sp in non_profs:
            if sp not in used_today[day]:
                slots_df.at[idx, 'StaffID'] = sp
                nonprof_day_count[sp].add(day)
                used_today[day].add(sp)
                break

still_unassigned = slots_df[slots_df['StaffID'].isna()]
if not still_unassigned.empty:
    print(f"{len(still_unassigned)} slots still unassigned. Please review manually.")
else:
    print("All slots assigned.")

slots_df.to_excel("Assigned_Staff_Slots_Full.xlsx", index=False)
slots_df.to_csv("Assigned_Staff_Slots_Full.csv", index=False)

print("StaffID assignment complete and saved!")


