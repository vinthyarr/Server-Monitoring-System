import smtplib
import wmi
import psutil
import csv
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

DISK_SPACE_THRESHOLD = 70
TEMPERATURE_THRESHOLD = 20

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_FROM = "senderitc.testproject@gmail.com"
EMAIL_PASSWORD = "hvzylrtakxbotghl"
EMAIL_TO = "receiveritc.testproject@gmail.com"
EMAIL_SUBJECT = "Server Monitoring Alert"

LOG_FILE = f"server_monitoring_{datetime.now().strftime('%Y-%m-%d')}.log"

# Configure logging to save information to the log file
logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format="%(asctime)s - %(levelname)s: %(message)s")

def send_email_alert(temp_alert, disk_alert, antivirus_alert):
    msg = MIMEMultipart()
    msg["Subject"] = EMAIL_SUBJECT
    msg["From"] = EMAIL_FROM
    msg["To"] = EMAIL_TO
    
    # Create a formatted string for the email body
    email_body = "IT Routine Activity Log                                            "
    
    # Create the table header
    table_header = "Date:" + datetime.now().strftime('%Y-%m-%d')
    table_header = "Daily Routine Activities\n\n"
    table_header += "Activity details                                                       Status                       Remarks                            Checked by\n"

    # Create a formatted string for temperature alerts
    temperature_text = ""
    for row in temp_alert:
        attribute_name = row[0]
        status = row[2]
        remarks = row[1]
        
        # Add extra spaces to align the columns
        attribute_name = " " * 70 + attribute_name.ljust(40)
        status = " " * 50 + status.ljust(10)
        remarks = " " * 50 + remarks.ljust(10)
        
        temperature_text += f"{attribute_name}{status}{remarks}\n"

    # Create a formatted string for disk space alerts
    disk_space_text = ""
    for row in disk_alert:
        attribute_name = row[0]
        status = row[2]
        remarks = row[1]
        
        # Add extra spaces to align the columns
        attribute_name = " " * 70 + attribute_name.ljust(40)
        status = " " * 50 + status.ljust(10)
        remarks = " " * 50 + remarks.ljust(10)
        
        disk_space_text += f"{attribute_name}{status}{remarks}\n"

    # Create a formatted string for antivirus alerts
    antivirus_text = ""
    for row in antivirus_alert:
        attribute_name = row[0]
        status = row[2]
        remarks = row[1]

        # Add extra spaces to align the columns
        attribute_name = " " * 70 + attribute_name.ljust(40)
        status = " " * 50 + status.ljust(10)
        remarks = " " * 50 + remarks.ljust(10)

        antivirus_text += f"{attribute_name}{status}{remarks}\n"

    # Combine all the parts of the email body
    email_body += table_header + temperature_text

    # Insert an empty line between temperature and disk space details
    email_body += "\n\n"

    email_body += "Disk Space Details\n\n" + disk_space_text

    # Insert an empty line between disk space and antivirus details
    email_body += "\n\n"

    email_body += "Antivirus Update Details\n\n" + antivirus_text

    text = MIMEText(email_body, "plain")
    msg.attach(text)

    # Create and attach the CSV file
    csv_file = f"server_alerts_{datetime.now().strftime('%Y-%m-%d')}.csv"
    with open(csv_file, 'w', newline='') as csvfile:
        csv_writer = csv.writer(csvfile)
        csv_writer.writerow([" " * 50 + "ITC TIRUVOTTIYUR" + " " * 50])
        csv_writer.writerow([" " * 50 + "IT Routine Activity Log" + " " * 50])
        csv_writer.writerow([" " * 100 + "Date: " + datetime.now().strftime('%Y-%m-%d')])
        csv_writer.writerow([])  # Empty row for spacing
        csv_writer.writerow([" " * 50 + "Daily Morning Routine Activities" + " " * 50])
        csv_writer.writerow(["Temperature Details"])
        csv_writer.writerow([" " * 10 + "Server details" + " " * 50, " " * 10 + "Status" + " " * 50, " " * 10 + "Remarks" + " " * 50, "Checked by"])
        
        for row in temp_alert:
            csv_writer.writerow(row)

        # Insert an empty row between temperature and disk space details
        csv_writer.writerow([])

        csv_writer.writerow(["Disk Space Details"])

        for row in disk_alert:
            csv_writer.writerow(row)

        # Insert an empty row between disk space and antivirus details
        csv_writer.writerow([])

        csv_writer.writerow(["Antivirus Update Details"])

        for row in antivirus_alert:
            csv_writer.writerow(row)

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(open(csv_file, 'rb').read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{csv_file}"')
    msg.attach(part)

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_FROM, EMAIL_PASSWORD)
        server.sendmail(EMAIL_FROM, [EMAIL_TO], msg.as_string())

def get_cpu_temperature(ip_address):
    try:
        w = wmi.WMI(namespace="root/OpenHardwareMonitor")
        temperature_info = w.Sensor()
        
        temp_alert = []

        for sensor in temperature_info:
            if sensor.SensorType == u'Temperature' and 'CPU' in sensor.Name:
                cpu_temperature = sensor.Value
                attribute_name = f"({sensor.Name})"
                
                if cpu_temperature >= TEMPERATURE_THRESHOLD:
                    status = "High Temperature"
                else:
                    status = "Ok"
                remarks = f"{cpu_temperature} °C"

                temp_alert.append([attribute_name, status, remarks])
        

        if not temp_alert:
            temp_alert.append(["CPU Temperature", "N/A", "Ok"])

        return temp_alert

    except Exception as e:
        return [["CPU Temperature", "N/A", str(e)]]

def monitor_disk_space(ip_address):
    disk_partitions = psutil.disk_partitions(all=True)
    disk_alert = []

    for partition in disk_partitions:
        partition_mount_point = partition.mountpoint
        partition_usage = psutil.disk_usage(partition_mount_point)
        
        free_space = partition_usage.free / (1024 ** 3)  # Convert to GB
        f"\n "
        f"\nDisk Space Details"
        attribute = f"Disk Space ({partition_mount_point})"
        
        if free_space < DISK_SPACE_THRESHOLD:
            statuss = "Low Disk Space"
        else:
            statuss = "Ok"
        remark = f"{free_space:.2f} GB"
        disk_alert.append([attribute, statuss, remark])

    if not disk_alert:
        disk_alert.append(["Disk Space", "N/A", "Ok"])

    return disk_alert

def is_antivirus_updating():
    antivirus_processes = ["MsMpEng.exe", "mcshield.exe", "ccSvcHst.exe"]
    
    for process in psutil.process_iter(attrs=['pid', 'name']):
        process_info = process.info
        process_name = process_info['name']
        
        if process_name in antivirus_processes:
            status = "Ok"
    
    return "Ok"

if __name__ == "__main__":
    ip_address = "localhost"  
    alerts = []
    
    temperature_alert = get_cpu_temperature(ip_address)
    disk_space_alerts = monitor_disk_space(ip_address)
    antivirus_status = is_antivirus_updating()

    alerts.extend(temperature_alert)
    alerts.extend(disk_space_alerts)
    alerts.append(["Antivirus Update", antivirus_status, ""])

    # Log the alerts to the log file
    for alert in alerts:
        logging.info(f"{alert[0]}: {alert[1]} ({alert[2]})")

    # Send the email with the alerts and the CSV attachment
    send_email_alert(temperature_alert, disk_space_alerts, [["Antivirus Update", antivirus_status, ""]])
        
