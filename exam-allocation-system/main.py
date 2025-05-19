# main.py
import subprocess

print("Starting Smart Invigilation Scheduler...")
subprocess.run(["python", "parser.py"])
subprocess.run(["python", "venue_splitter.py"])
subprocess.run(["python", "assign_staff.py"])
print("All modules executed successfully!")
