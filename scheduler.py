"""
Optional: Schedule reports to run automatically
Install: pip install schedule
"""
import schedule
import time
from report_generator import ExcelReportGenerator
from config import Config

def job():
    print("Running scheduled report...")
    generator = ExcelReportGenerator()
    generator.run(recipient_email="manager@company.com")

# Schedule daily at 9 AM
schedule.every().day.at(Config.SCHEDULE_TIME).do(job)

print(f"⏰ Scheduler started. Reports will run daily at {Config.SCHEDULE_TIME}")
print("Press Ctrl+C to stop")

while True:
    schedule.run_pending()
    time.sleep(60)
