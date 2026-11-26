import schedule
import time
import subprocess
import os
from datetime import datetime

# Log file path
LOG_FILE = "scheduler_logs.txt"

def log_message(message):
    """Write logs with timestamp"""
    with open(LOG_FILE, "a") as f:
        f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")
    print(message)

def run_program(program_path, working_dir):
    """Run python program from a specific directory"""
    try:
        log_message(f"Starting {program_path} in {working_dir}")
        result = subprocess.run(
            ["python", program_path],
            cwd=working_dir,
            capture_output=True,
            text=True
        )
        if result.returncode == 0:
            log_message(f"Successfully finished {program_path}")
            log_message(result.stdout)
        else:
            log_message(f"Error running {program_path}")
            log_message(result.stderr)
    except Exception as e:
        log_message(f"Exception: {str(e)}")

# === Program paths and directories ===
# Keys are added in the order in which the corresponding report automation is initiated.
programs = {
    "report_Pamith": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\risk_based_pricing_rate_approval_report\report-Pamith"
    },
    "report_Nimna_V03": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\risk_based_pricing_rate_approval_report\report-Nimna_V03"
    },
    "Annual_Cr_review": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\Annual_Cr_review"
    },
    "report_Nimna_RBP_5pm": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\risk_based_pricing_rate_approval_report\report-Nimna_RBP_5pm"
    },
    "Decision_Pending": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\Decision_Pending"
    },
    "Decision_Pending_4pm": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\Decision Pending 4pm"
    },
    "RBP_full_report": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\risk_based_pricing_rate_approval_report\report_Pamith_V03"
    },
    "Manual_report": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\risk_based_pricing_rate_approval_report\Manual_value_report"
    },
    "ex_director_report": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\risk_based_pricing_rate_approval_report\report_Pamith_V03_Exec_dirt"
    },
    "Document_Verification_report": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\Document_Verification"
    },
    "Document_Verification_monthly_report": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\Document_Verification_monthly"
    },
    "evening_status_report": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\Evening Status"
    },
    "morning_status_report": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\morning_status"
    },
    "Security_Auth_performance_report": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\Security_Auth_performance"
    },
    "CED_Performance_report": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\CED_Performance"
    },
    "LOS_Performance_report": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\los_performance_validation"
    },
    "Fraud_Incident_Register_Report": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\Fraud_Incident_Register_Report"
    },
    "Approved_facility_details_report": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\Approved_Facility_Details_Report_V2"
    },
    "CED_Performance_monthly": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\CED_Performance_monthly"
    },
    "Vehicle_Loan_Initiation_Report": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\Vehicle_Loan_Initiation_Report"
    },
    "3W_Auto_Finance_Report": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\3W_Auto_Finance_Report"
    },
    "RBP_5_10_PM_Report": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\risk_based_pricing_rate_approval_report\report_Pamith_V03_Copy"
    },
    "Initiated_Facility_Details_Report": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\Initiated_Facility_Details_Report"
    },
    "Auto_Finance_AI_Credit_Decision_Report_Monitoring": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\Auto_Finance_AI_Credit_Decision_Report_Monitoring"
    },
    "LTV_Exception_Daily_Report": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\LTV_Exception_Report_daily"
    },
    "3W_AI_Credit_Decision_Report_Monitoring": {
        "path": "main_orchestrator.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\3W_AI_Credit_Decision_Report_Monitoring"
    },
    "morning_status_before_after_4pm" : {
        "path": "main_orchestrator.py",
        "dir" : r"C:\Users\sachintha.fernando\Documents\Report_Automation\morning_status_before_after_4pm"
    }, 
    "CRIB_Downlaoding_Report" : {
        "path": "main_orchestrator.py",
        "dir" : r"C:\Users\sachintha.fernando\Documents\Report_Automation\CRIB_Downloading_Report"
    },
    "Cumulative_Exposure_Report" : {
        "path": "main_orchestrator.py",
        "dir" : r"C:\Users\sachintha.fernando\Documents\Report_Automation\Cumulative_Exposure_Report"
    },
    "Credit_Approval_Limits_Cumulative_Report": {
        "path": "main_orchestrator.py",
        "dir" : r"C:\Users\sachintha.fernando\Documents\Report_Automation\Credit_Approval_Limits_Cumulative_Report"
    },
    "CRIB_Requests_&_Initiation_Report" : {
        "path": "main_orchestrator.py",
        "dir" : r"C:\Users\sachintha.fernando\Documents\Report_Automation\CRIB_Requests_&_Initiation_Report"
    },
    "vehicle_manual_creator" : {
        "path": "auto-manual2.py",
        "dir": r"C:\Users\sachintha.fernando\Documents\Report_Automation\vehicle_manual_creator"
    },
    "Annual_Cr_Self_Unavailable_Summary_Report" :{
        "path": "main_orchestrator.py",
        "dir" : r"C:\Users\sachintha.fernando\Documents\Report_Automation\Annual_Cr_Self_Unavailable_List"
    },
    "TIL_Customer_Conversion_Report":{
        "path": "main_orchestrator.py",
        "dir" : r"C:\Users\sachintha.fernando\Documents\Report_Automation\TIL_Customer_Conversion"
    },
    "Special_Report":{
        "path": "main_orchestrator.py",
        "dir" : r"C:\Users\sachintha.fernando\Documents\Report_Automation\Special_Report_Mr_Dave"
    }
}

# === Scheduling functions (skip Sundays) === 
# Scheduling functions are added in the order in which the corresponding report automation is initiated.
# Pamith RBR Detailed 5pm
def schedule_report_Pamith():
    if datetime.today().weekday() != 6:  # 6 = Sunday
        run_program(programs["report_Pamith"]["path"], programs["report_Pamith"]["dir"])

# Nimna RBP 8.30am
def schedule_report_Nimna_V03_morning():
    if datetime.today().weekday() != 6:
        run_program(programs["report_Nimna_V03"]["path"], programs["report_Nimna_V03"]["dir"])

# Nimna RBP 5pm
def schedule_report_Nimna_V02_evening():
    if datetime.today().weekday() != 6:
        run_program(programs["report_Nimna_RBP_5pm"]["path"], programs["report_Nimna_RBP_5pm"]["dir"])

# Annual Credit Review 10.00am
def schedule_Annual_Cr_review():
    if datetime.today().weekday() != 6:
        run_program(programs["Annual_Cr_review"]["path"], programs["Annual_Cr_review"]["dir"])

def schedule_Decision_Pending():
    if datetime.today().weekday() != 6:  # Changed from 7 to 6 (Sunday)
        run_program(programs["Decision_Pending"]["path"], programs["Decision_Pending"]["dir"])

def schedule_Decision_Pending_4pm():
    if datetime.today().weekday() != 6:  # Changed from 7 to 6 (Sunday)
        run_program(programs["Decision_Pending_4pm"]["path"], programs["Decision_Pending_4pm"]["dir"])

def schedule_RBP_full_report():
    if datetime.today().weekday() <= 6: # Run every day including Sunday
        run_program(programs["RBP_full_report"]["path"], programs["RBP_full_report"]["dir"])

def schedule_Manual_report():
    if datetime.today().weekday() != 6:
        run_program(programs["Manual_report"]["path"], programs["Manual_report"]["dir"])

def schedule_ex_director_report():
    if datetime.today().weekday() <= 6:  # Run every day including Sunday
        run_program(programs["ex_director_report"]["path"], programs["ex_director_report"]["dir"])

def schedule_Document_Verification_report():
    if datetime.today().weekday() != 6:  # Skip Sundays
        run_program(programs["Document_Verification_report"]["path"], programs["Document_Verification_report"]["dir"])

def schedule_Document_Verification_monthly_report():
    if datetime.today().day == 1:  # Run on the 1st of each month
        run_program(
            programs["Document_Verification_monthly_report"]["path"], 
            programs["Document_Verification_monthly_report"]["dir"]
        )

def schedule_evening_status_report():
    if datetime.today().weekday() != 6:
        run_program(programs["evening_status_report"]["path"], programs["evening_status_report"]["dir"])

def schedule_morning_status_report():
    if datetime.today().weekday() != 6:
        run_program(programs["morning_status_report"]["path"], programs["morning_status_report"]["dir"])

def schedule_Security_Auth_performance_report():
    if datetime.today().weekday() != 6:
        run_program(programs["Security_Auth_performance_report"]["path"], programs["Security_Auth_performance_report"]["dir"])

def schedule_CED_performance_report():
    if datetime.today().weekday() != 6:
        run_program(programs["CED_Performance_report"]["path"], programs["CED_Performance_report"]["dir"])

def schedule_LOS_Performance_report():
    if datetime.today().weekday() != 6:  # Skip Sundays
        current_hour = datetime.now().hour
        # Run only between 8:55 AM and 6:55 PM
        if 8 <= current_hour <= 18:
            run_program(programs["LOS_Performance_report"]["path"], programs["LOS_Performance_report"]["dir"])

def schedule_Fraud_Incident_Register_Report():
    if datetime.today().weekday() != 6:
        run_program(programs["Fraud_Incident_Register_Report"]["path"], programs["Fraud_Incident_Register_Report"]["dir"])


def schedule_Approved_facility_details_report():
    if datetime.today().weekday() != 0:
        run_program(programs["Approved_facility_details_report"]["path"], programs["Approved_facility_details_report"]["dir"])

# def schedule_CED_Performance_monthly():
#     if datetime.today().day == 1:
#         run_program(
#             programs["CED_Performance_monthly"]["path"], programs["CED_Performance_monthly"]["dir"])

# def schedule_CED_Performance_monthly():
#     """Run on 1st of every month at 3:00 AM"""
#     if datetime.today().day == 1:  # Only run on 1st of month
#         run_program(
#             programs["CED_Performance_monthly"]["path"], 
#             programs["CED_Performance_monthly"]["dir"]
#         )
        
def schedule_Vehicle_Loan_Initiation_Report():
    current_datetime = datetime.now()
    if current_datetime.weekday() != 6:
        current_hour = current_datetime.hour
        run_program(programs["Vehicle_Loan_Initiation_Report"]["path"], programs["Vehicle_Loan_Initiation_Report"]["dir"])  

def schedule_3W_Auto_Finance_Report():
    current_datetime = datetime.now()
    if current_datetime.weekday() != 6:
        current_hour = current_datetime.hour
        run_program(programs["3W_Auto_Finance_Report"]["path"], programs["3W_Auto_Finance_Report"]["dir"])
        
def schedule_RBP_5_10_PM_Report():
    current_datetime = datetime.now()
    if current_datetime.weekday() != 6:
        current_hour = current_datetime.hour
        run_program(programs["RBP_5_10_PM_Report"]["path"], programs["RBP_5_10_PM_Report"]["dir"])
        
def schedule_Initiated_Facility_Details_Report():
    current_datetime = datetime.now()
    if current_datetime.weekday() != 6:
        run_program(programs["Initiated_Facility_Details_Report"]["path"], programs["Initiated_Facility_Details_Report"]["dir"])

def schedule_Auto_Finance_Credit_Decision_AI_Report_Monitoring():
    current_datetime = datetime.now()
    if current_datetime.weekday() <= 6:
        run_program(programs["Auto_Finance_AI_Credit_Decision_Report_Monitoring"]["path"], programs["Auto_Finance_AI_Credit_Decision_Report_Monitoring"]["dir"])

def schedule_LTV_Exception_Daily_Report():
    current_datetime = datetime.now()
    if current_datetime.weekday() != 6:
        run_program(programs["LTV_Exception_Daily_Report"]["path"], programs["LTV_Exception_Daily_Report"]["dir"])
        
def schedule_3W_Credit_Decision_AI_Report_Monitoring():
    current_datetime = datetime.now()
    if current_datetime.weekday() != 6:
        run_program(programs["3W_AI_Credit_Decision_Report_Monitoring"]["path"], programs["3W_AI_Credit_Decision_Report_Monitoring"]["dir"])

def schedule_morning_status_before_after_4pm_report():
    current_datetime = datetime.now()
    if current_datetime.weekday() != 6:
        run_program(programs["morning_status_before_after_4pm"]["path"], programs["morning_status_before_after_4pm"]["dir"])
        
def schedule_crib_downloading_report():
    if datetime.today().day == 1: 
        run_program(programs["CRIB_Downlaoding_Report"]["path"], programs["CRIB_Downlaoding_Report"]["dir"])

def schedule_cumulative_exposure_report():
    current_datetime = datetime.now()
    if current_datetime.weekday() != 6:
        run_program(programs["Cumulative_Exposure_Report"]["path"], programs["Cumulative_Exposure_Report"]["dir"])

def schedule_credit_approval_limits_cumulative_report():
    current_datetime = datetime.now()
    if current_datetime.weekday() <=6:
        run_program(programs["Credit_Approval_Limits_Cumulative_Report"]["path"], programs["Credit_Approval_Limits_Cumulative_Report"]["dir"])
    

def schedule_CRIB_Requests_Initiation_Report():
    current_datetime = datetime.now()
    if current_datetime.weekday() <= 6:
        run_program(programs["CRIB_Requests_&_Initiation_Report"]["path"], programs["CRIB_Requests_&_Initiation_Report"]["dir"])

def schedule_vehicle_manual_creator():
    current_datetime = datetime.now()
    if current_datetime.weekday() > 0:
        run_program(programs["vehicle_manual_creator"]["path"], programs["vehicle_manual_creator"]["dir"])

def schedule_annual_cr_self_unavailable_summary_report():
    current_datetime = datetime.now()
    if current_datetime.weekday() <= 6:
        run_program(programs["Annual_Cr_Self_Unavailable_Summary_Report"]["path"], programs["Annual_Cr_Self_Unavailable_Summary_Report"]["dir"])

def schedule_TIL_Customer_Conversion_Report():
    current_datetime = datetime.now()
    if current_datetime.weekday() <= 6:
        run_program(programs["TIL_Customer_Conversion_Report"]["path"], programs["TIL_Customer_Conversion_Report"]["dir"])

def schedule_Special_Report():
    current_datetime = datetime.now()
    if current_datetime.weekday() <= 6:
        run_program(programs["Special_Report"]["path"], programs["Special_Report"]["dir"])

# === Daily schedules ===
# Uncomment if needed: schedule.every().day.at("07:58").do(schedule_Decision_Pending)
# schedule.every().day.at("03:00").do(schedule_CED_Performance_monthly)
schedule.every().day.at("01:00").do(schedule_TIL_Customer_Conversion_Report)
schedule.every().day.at("03:00").do(schedule_Special_Report)
schedule.every().day.at("03:30").do(schedule_vehicle_manual_creator)
schedule.every().day.at("04:00").do(schedule_cumulative_exposure_report)
schedule.every().day.at("04:30").do(schedule_credit_approval_limits_cumulative_report)
schedule.every().day.at("05:00").do(schedule_RBP_full_report)
schedule.every().day.at("05:15").do(schedule_ex_director_report)
schedule.every().day.at("05:30").do(schedule_annual_cr_self_unavailable_summary_report)
schedule.every().day.at("05:45").do(schedule_report_Nimna_V03_morning)
schedule.every().day.at("06:00").do(schedule_Approved_facility_details_report)
schedule.every().day.at("06:15").do(schedule_crib_downloading_report)
schedule.every().day.at("06:30").do(schedule_Auto_Finance_Credit_Decision_AI_Report_Monitoring)
schedule.every().day.at("06:40").do(schedule_3W_Credit_Decision_AI_Report_Monitoring)
schedule.every().day.at("06:50").do(schedule_morning_status_report)
schedule.every().day.at("07:10").do(schedule_Document_Verification_report)       
schedule.every().day.at("07:15").do(schedule_Document_Verification_monthly_report) 
schedule.every().day.at("07:25").do(schedule_CED_performance_report)
schedule.every().day.at("07:30").do(schedule_morning_status_before_after_4pm_report)
schedule.every().day.at("08:00").do(schedule_Initiated_Facility_Details_Report)
schedule.every().day.at("08:30").do(schedule_CRIB_Requests_Initiation_Report)
schedule.every().day.at("09:00").do(schedule_Security_Auth_performance_report)
schedule.every().day.at("09:15").do(schedule_Manual_report)
schedule.every().day.at("10:00").do(schedule_Annual_Cr_review)
schedule.every().day.at("10:15").do(schedule_Vehicle_Loan_Initiation_Report)
schedule.every().day.at("10:30").do(schedule_3W_Auto_Finance_Report)
schedule.every().day.at("14:00").do(schedule_Fraud_Incident_Register_Report)
schedule.every().day.at("16:30").do(schedule_evening_status_report)
# schedule.every().day.at("16:30").do(schedule_Decision_Pending_4pm) # Uncomment if needed: 
# schedule.every().day.at("17:10").do(schedule_RBP_5_10_PM_Report) # Served for testing:  

# Schedule LOS Performance Report every hour at :55 from 8:55 AM to 6:55 PM

def hourly_LOS_schedule():
    """Wrapper function to schedule LOS Performance Report at :55 every hour"""
    if datetime.today().weekday() != 6:  # Skip Sundays
        current_hour = datetime.now().hour
        # Only schedule if within 8 AM to 6 PM range
        if 8 <= current_hour <= 18:
            schedule_LOS_Performance_report()

# Schedule to run every hour at :55
schedule.every().hour.at(":55").do(hourly_LOS_schedule)

# Schedule LTV Exception Daily Report every two hours at :30 from 9:30 AM to 9:30 PM
def hourly_LTV_Exception_schedule():
    """Wrapper function to schedule LTV Exception Daily Report at :30 every two hours"""
    if datetime.today().weekday() != 6:  # Skip Sundays
        current_hour = datetime.now().hour
        current_minute = datetime.now().minute
        
        # Check if current time is within 9:30 AM to 9:30 PM
        # and only run at :30 minute mark for specific hours
        if (9 <= current_hour <= 21) and current_minute == 30:
            # Only run at 9:30, 11:30, 13:30, 15:30, 17:30, 19:30, 21:30
            if current_hour % 2 == 1 or current_hour == 9:  # Odd hours (9,11,13,15,17,19,21)
                schedule_LTV_Exception_Daily_Report()

# Schedule to run every hour at :30, but will only execute on odd hours
schedule.every().hour.at(":30").do(hourly_LTV_Exception_schedule)

log_message("Scheduler started. Waiting for tasks...")

# === Run loop ===
while True:
    schedule.run_pending()
    time.sleep(30)