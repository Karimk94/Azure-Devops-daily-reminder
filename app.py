import requests
from requests_negotiate_sspi import HttpNegotiateAuth
import smtplib
from email.message import EmailMessage
from datetime import datetime
import logging
import os
import sys
from urllib.parse import quote
import win32timezone 

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

# Configuration
ADO_URL = "https://rtadevops.rta.global"
COLLECTION = "ITSD"
API_VERSION = "6.0"
SMTP_SERVER = "rtauedcsmtp.RTA.Global"
FROM_EMAIL = "devops@rta.ae"
DATE_FIELD = "Custom.TargetCompletionDate"
LOG_FILE = os.path.join(application_path, "task_monitor.log")

# Setup Logging
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

def send_summary_email(to_email, tasks):
    count = len(tasks)
    subject = f"Action Required: {count} Overdue Task{'s' if count > 1 else ''}"
    
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = FROM_EMAIL
    msg['To'] = to_email
    
    # Generate Task Rows matches ADO list style
    task_rows = ""
    for task in tasks:
        task_rows += f"""
        <tr>
            <td valign="top" style="padding: 12px 0; border-bottom: 1px solid #EAEAEA; width: 60px;">
                <a href="{task['link']}" style="font-family: 'Segoe UI', sans-serif; font-size: 14px; font-weight: 600; color: #0078D4; text-decoration: none;">
                    {task['id']}
                </a>
            </td>
            <td valign="top" style="padding: 12px 16px; border-bottom: 1px solid #EAEAEA;">
                <div style="margin-bottom: 4px;">
                    <a href="{task['link']}" style="font-family: 'Segoe UI', sans-serif; font-size: 14px; font-weight: 600; color: #212121; text-decoration: none;">
                        {task['title']}
                    </a>
                </div>
                <div style="font-family: 'Segoe UI', sans-serif; font-size: 12px; color: #666666;">
                    {task['project']}
                </div>
            </td>
            <td valign="top" align="right" style="padding: 12px 0; border-bottom: 1px solid #EAEAEA; white-space: nowrap;">
                <span style="font-family: 'Segoe UI', sans-serif; font-size: 12px; font-weight: 600; color: #D93025; background-color: #FCE8E6; padding: 4px 8px; border-radius: 4px;">
                    {task['due_date']}
                </span>
            </td>
        </tr>
        """

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, Arial, sans-serif; margin: 0; padding: 0; }}
            a {{ text-decoration: none; }}
            a:hover {{ text-decoration: underline; }}
        </style>
    </head>
    <body style="margin: 0; padding: 0; background-color: #F8F8F8;">
        <!-- Full Width Header -->
        <table width="100%" border="0" cellspacing="0" cellpadding="0" style="width:100%; background-color: white; border-bottom: 1px solid #EAEAEA;">
            <tr>
                <td align="center">
                    <table width="640" border="0" cellspacing="0" cellpadding="0" style="width: 640px; max-width: 640px;">
                        <tr>
                            <td style="padding: 16px 0;">
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                        <!-- Logo Section -->
                                        <td valign="middle">
                                            <!-- Embedded CID Logo -->
                                            <img src="cid:logo_image" alt="RTA" height="55" style="display: block; border: 0;">
                                        </td>
                                        <!-- Right Title -->
                                        <td align="right" valign="middle" style="font-family: 'Segoe UI', sans-serif; font-size: 14px; color: #212121;">
                                            <span style="color: #666666;">ITSD</span> <span style="color: #0078D4; font-weight: bold;">Task Tracker</span>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>

        <!-- Main Body -->
        <table width="100%" border="0" cellspacing="0" cellpadding="0" style="width:100%; background-color: #F8F8F8;">
            <tr>
                <td align="center" style="padding: 24px 0;">
                    <!-- Content Card -->
                    <table width="640" border="0" cellspacing="0" cellpadding="0" style="width: 640px; max-width: 640px; background-color: white; border: 1px solid #EAEAEA;">
                        <tr>
                            <td style="padding: 32px;">
                                <!-- Alert Title -->
                                <div style="font-family: 'Segoe UI', sans-serif; font-size: 24px; font-weight: 300; color: #0078D4; margin-bottom: 24px;">
                                    Action Required: <span style="font-weight: 600;">{count} Overdue Task{'s' if count > 1 else ''}</span>
                                </div>
                                
                                <div style="font-family: 'Segoe UI', sans-serif; font-size: 14px; color: #333333; line-height: 20px; margin-bottom: 24px;">
                                    The following work items assigned to you have exceeded their target completion date.
                                </div>

                                <!-- Tasks Table -->
                                <table width="100%" border="0" cellspacing="0" cellpadding="0" style="width: 100%; border-collapse: collapse;">
                                    <thead>
                                        <tr>
                                            <th align="left" style="font-family: 'Segoe UI', sans-serif; font-size: 12px; color: #666666; font-weight: 600; border-bottom: 2px solid #F1F3F4; padding-bottom: 8px;">ID</th>
                                            <th align="left" style="font-family: 'Segoe UI', sans-serif; font-size: 12px; color: #666666; font-weight: 600; border-bottom: 2px solid #F1F3F4; padding: 0 16px 8px 16px;">DETAILS</th>
                                            <th align="right" style="font-family: 'Segoe UI', sans-serif; font-size: 12px; color: #666666; font-weight: 600; border-bottom: 2px solid #F1F3F4; padding-bottom: 8px;">DUE DATE</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {task_rows}
                                    </tbody>
                                </table>

                                <!-- Arabic Section -->
                                <div style="margin-top: 32px; padding-top: 20px; border-top: 1px solid #EAEAEA; text-align: center; direction: rtl;">
                                    <p style="font-family: 'Segoe UI', Tahoma, Arial, sans-serif; font-size: 14px; color: #333333; margin: 4px 0;">يرجى الملاحظةأن المهام المذكورة أعلاه قد تجاوزت تاريخ الاستحقاق.</p>
                                    <p style="font-family: 'Segoe UI', Tahoma, Arial, sans-serif; font-size: 14px; color: #333333; margin: 4px 0;">يرجى اتخاذ الإجراء اللازم.</p>
                                </div>
                            </td>
                        </tr>
                    </table>

                    <!-- Footer -->
                    <table width="640" border="0" cellspacing="0" cellpadding="0" style="width: 640px; max-width: 640px; margin-top: 16px;">
                        <tr>
                            <td align="center" style="font-family: 'Segoe UI', sans-serif; font-size: 12px; color: #70757A; line-height: 18px;">
                                &copy; {datetime.now().year} RTA. All rights reserved.<br>
                                This is an automated notification. Please do not reply.
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </body>
    </html>
    """
    
    msg.set_content("This email requires an HTML-enabled client to view properly.", subtype='plain')
    msg.add_alternative(html_content, subtype='html')

    # Embed Logo
    logo_path = os.path.join(application_path, "logo.png")
    if os.path.exists(logo_path):
        with open(logo_path, 'rb') as f:
            file_data = f.read()
            msg.get_payload()[1].add_related(file_data, maintype='image', subtype='png', cid='logo_image')
    else:
        logging.warning(f"Logo not found at {logo_path}")

    try:
        with smtplib.SMTP(SMTP_SERVER, 25) as smtp:
            smtp.send_message(msg)
            logging.info(f"SENT: To: {to_email} | Count: {count}")
    except Exception as e:
        logging.error(f"FAILED: To: {to_email} | Error: {e}")

def get_email_from_identity_url(url, auth):
    try:
        response = requests.get(url, auth=auth, verify=False)
        if response.status_code == 200:
            data = response.json()
            
            properties = data.get('Properties') or data.get('properties') or {}
            
            def extract_val(v):
                if isinstance(v, dict):
                    return v.get('$value')
                return v

            if 'Mail' in properties:
                return extract_val(properties['Mail'])
            
            if 'Email' in properties:
                return extract_val(properties['Email'])

            if 'emailAddress' in data: 
                return data['emailAddress']
    except Exception as e:
        logging.error(f"Debug: Failed to fetch identity from {url}: {e}")
    return None

def test_email_delivery(target_email):
    
    today = datetime.now().date()
    email_cache = {}
    tasks_by_email = {}
    
    wiql_url = f"{ADO_URL}/{COLLECTION}/_apis/wit/wiql?api-version={API_VERSION}"
    query_payload = {
        "query": "SELECT [System.Id] FROM WorkItems WHERE [System.State] <> 'Closed'"
    }
    
    auth = HttpNegotiateAuth()
    
    try:
        response = requests.post(wiql_url, json=query_payload, auth=auth, verify=False)
        response.raise_for_status()
        work_items = response.json().get("workItems", [])
    except Exception as e:
        logging.error(f"Error querying WIQL: {e}")
        return

    for item_ref in work_items:
        item_id = item_ref['id']
        item_url = f"{ADO_URL}/{COLLECTION}/_apis/wit/workitems/{item_id}?api-version={API_VERSION}"
        
        try:
            item_resp = requests.get(item_url, auth=auth, verify=False)
            if item_resp.status_code != 200: continue
            
            item = item_resp.json()
            fields = item.get('fields', {})
            
            # Check Date (Strict Overdue Check)
            target_date_str = fields.get(DATE_FIELD)
            if not target_date_str: continue
            
            try:
                due_dt = datetime.fromisoformat(target_date_str.replace('Z', '+00:00'))
                due_date = due_dt.date()
            except ValueError: continue

            if due_date >= today: continue # Not overdue

            assigned_to = fields.get('System.AssignedTo')
            if not assigned_to: continue
            
            to_email = None
            unique_identifier = None

            if isinstance(assigned_to, dict):
                unique_identifier = assigned_to.get('uniqueName')
                identity_url = assigned_to.get('url')
                
                if unique_identifier and '@' in unique_identifier and '.' in unique_identifier:
                    to_email = unique_identifier
                elif identity_url:
                    if unique_identifier in email_cache:
                        to_email = email_cache[unique_identifier]
                    else:
                        fetched_email = get_email_from_identity_url(identity_url, auth)
                        if fetched_email:
                            to_email = fetched_email
                            email_cache[unique_identifier] = to_email

            elif isinstance(assigned_to, str):
                if '<' in assigned_to:
                    to_email = assigned_to.split('<')[1].strip('>')
                elif '@' in assigned_to:
                    to_email = assigned_to

            if not to_email or '@' not in to_email: continue

            # Collect Task
            title = fields.get('System.Title', 'No Title')
            project_name = fields.get('System.TeamProject', 'ITSD Task Tracker')
            formatted_date = due_date.strftime("%d/%m/%Y")
            link = f"{ADO_URL}/{COLLECTION}/{quote(project_name)}/_workitems/edit/{item_id}/"
            
            if to_email not in tasks_by_email:
                tasks_by_email[to_email] = []
            
            tasks_by_email[to_email].append({
                'id': item_id,
                'title': title,
                'project': project_name,
                'due_date': formatted_date,
                'link': link
            })
            
        except Exception:
            continue

    # Logic to select the best candidate
    candidate_tasks = []
    found_multi = False

    for email, tasks in tasks_by_email.items():
        if len(tasks) > 1:
            candidate_tasks = tasks
            found_multi = True
            print(f"Selected candidate: {email} with {len(tasks)} tasks.")
            break
            
    if not found_multi and tasks_by_email:
        # Fallback to any user
        first_email = list(tasks_by_email.keys())[0]
        candidate_tasks = tasks_by_email[first_email]
        print(f"No multi-task candidate found. Falling back to: {first_email}")

    if candidate_tasks:
        send_summary_email(target_email, candidate_tasks)
        print("Test email sent successfully.")
    else:
        print("No overdue tasks found in the system to test.")

def monitor_tasks():
    today = datetime.now().date()
    email_cache = {}
    tasks_by_email = {}
    
    wiql_url = f"{ADO_URL}/{COLLECTION}/_apis/wit/wiql?api-version={API_VERSION}"
    query_payload = {
        "query": "SELECT [System.Id] FROM WorkItems WHERE [System.State] <> 'Closed'"
    }
    
    auth = HttpNegotiateAuth()
    
    try:
        response = requests.post(wiql_url, json=query_payload, auth=auth, verify=False)
        response.raise_for_status()
        work_items = response.json().get("workItems", [])
    except Exception as e:
        logging.error(f"Error querying WIQL: {e}")
        return

    for item_ref in work_items:
        item_id = item_ref['id']
        item_url = f"{ADO_URL}/{COLLECTION}/_apis/wit/workitems/{item_id}?api-version={API_VERSION}"
        
        try:
            item_resp = requests.get(item_url, auth=auth, verify=False)
            if item_resp.status_code != 200: 
                continue
                
            item = item_resp.json()
            fields = item.get('fields', {})
            
            target_date_str = fields.get(DATE_FIELD)
            if not target_date_str:
                continue
                
            try:
                due_dt = datetime.fromisoformat(target_date_str.replace('Z', '+00:00'))
                due_date = due_dt.date()
            except ValueError:
                continue

            if due_date >= today:
                print(f"Skipping ID {item_id}: Not overdue (Due: {due_date})")
                continue
                
            assigned_to = fields.get('System.AssignedTo')
            if not assigned_to:
                continue
            
            to_email = None
            unique_identifier = None

            if isinstance(assigned_to, dict):
                unique_identifier = assigned_to.get('uniqueName')
                identity_url = assigned_to.get('url')
                
                if unique_identifier and '@' in unique_identifier and '.' in unique_identifier:
                    to_email = unique_identifier
                elif identity_url:
                    if unique_identifier in email_cache:
                        to_email = email_cache[unique_identifier]
                    else:
                        fetched_email = get_email_from_identity_url(identity_url, auth)
                        if fetched_email:
                            to_email = fetched_email
                            print(f"Resolved email for {unique_identifier}: {to_email}")
                            if unique_identifier:
                                email_cache[unique_identifier] = to_email

            elif isinstance(assigned_to, str):
                if '<' in assigned_to and '>' in assigned_to:
                    start = assigned_to.find('<') + 1
                    end = assigned_to.find('>', start)
                    to_email = assigned_to[start:end]
                elif '@' in assigned_to:
                    to_email = assigned_to

            if not to_email or '@' not in to_email:
                print(f"Could not resolve email for item {item_id}. Assignee: {unique_identifier or assigned_to}")
                continue

            # Instead of sending immediately, collect the task
            title = fields.get('System.Title', 'No Title')
            project_name = fields.get('System.TeamProject', 'ITSD Task Tracker')
            formatted_date = due_date.strftime("%d/%m/%Y")
            link = f"{ADO_URL}/{COLLECTION}/{quote(project_name)}/_workitems/edit/{item_id}/"
            
            if to_email not in tasks_by_email:
                tasks_by_email[to_email] = []
            
            tasks_by_email[to_email].append({
                'id': item_id,
                'title': title,
                'project': project_name,
                'due_date': formatted_date,
                'link': link
            })
            
        except Exception as e:
            print(f"Error processing item {item_id}: {e}")

    for email, tasks in tasks_by_email.items():
        send_summary_email(email, tasks)

if __name__ == "__main__":
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    
    # monitor_tasks()
    
    # Uncomment the line below to test email delivery to your personal address
    test_email_delivery("okool_karim.abdulwahed@rta.ae")