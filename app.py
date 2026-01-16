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

def send_overdue_email(to_email, item_id, title, due_date_str, project_name):
    link = f"{ADO_URL}/{COLLECTION}/{quote(project_name)}/_workitems/edit/{item_id}/"
    
    msg = EmailMessage()
    msg['Subject'] = f"Action Required: Overdue Task {item_id} - {title}"
    msg['From'] = FROM_EMAIL
    msg['To'] = to_email
    
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
    </head>
    <body style="font-family: 'Google Sans', Roboto, Helvetica, Arial, sans-serif; background-color: #f0f2f5; margin: 0; padding: 0; color: #202124;">
        <!-- Wrapper Table for Centering in Outlook -->
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#f0f2f5" style="width: 100%; background-color: #f0f2f5;">
            <tr>
                <td align="center" style="padding: 40px 0;">
                    
                    <!-- Main Container Table (Fixed Width 600px) -->
                    <table width="600" border="0" cellpadding="0" cellspacing="0" bgcolor="#ffffff" style="width: 600px; max-width: 600px; background-color: #ffffff; border-radius: 8px; font-family: sans-serif; box-shadow: 0 1px 2px 0 rgba(60,64,67,0.3);">
                        
                        <!-- Header -->
                        <tr>
                            <td align="center" style="padding: 24px 24px 0 24px;">
                                <div style="font-size: 24px; color: #5f6368; font-weight: 400;">ITSD Task Tracker</div>
                            </td>
                        </tr>

                        <!-- Content -->
                        <tr>
                            <td style="padding: 24px;">
                                <div style="color: #d93025; font-size: 48px; text-align: center; margin-bottom: 16px; line-height: 1;">&#9888;</div>
                                <h1 style="font-size: 22px; font-weight: 400; margin: 0 0 16px 0; color: #202124; text-align: center;">Action Required: Task Overdue</h1>
                                <p style="font-size: 16px; line-height: 1.5; margin-bottom: 24px; color: #3c4043; text-align: center;">The following work item has exceeded its target completion date.</p>
                                
                                <!-- Card (Nested Table for Border/Padding support in Outlook) -->
                                <table width="100%" border="0" cellpadding="16" cellspacing="0" bgcolor="#f8f9fa" style="background-color: #f8f9fa; border: 1px solid #dadce0; border-radius: 8px; text-align: left;">
                                    <tr>
                                        <td>
                                            <div style="margin-bottom: 12px;">
                                                <span style="font-size: 12px; font-weight: 500; letter-spacing: 0.2px; color: #5f6368; text-transform: uppercase; display: block; margin-bottom: 4px;">Task ID</span>
                                                <div style="font-size: 16px; color: #202124;">{item_id}</div>
                                            </div>

                                            <div style="margin-bottom: 12px;">
                                                <span style="font-size: 12px; font-weight: 500; letter-spacing: 0.2px; color: #5f6368; text-transform: uppercase; display: block; margin-bottom: 4px;">Title</span>
                                                <div style="font-size: 16px; color: #202124;">{title}</div>
                                            </div>

                                            <div style="margin-bottom: 12px;">
                                                <span style="font-size: 12px; font-weight: 500; letter-spacing: 0.2px; color: #5f6368; text-transform: uppercase; display: block; margin-bottom: 4px;">Project</span>
                                                <div style="font-size: 16px; color: #202124;">{project_name}</div>
                                            </div>

                                            <div style="margin-bottom: 0;">
                                                <span style="font-size: 12px; font-weight: 500; letter-spacing: 0.2px; color: #5f6368; text-transform: uppercase; display: block; margin-bottom: 4px;">Due Date</span>
                                                <div style="font-size: 16px; color: #d93025; font-weight: 500;">{due_date_str}</div>
                                            </div>
                                        </td>
                                    </tr>
                                </table>

                                <!-- Button -->
                                <div style="text-align: center; margin-top: 24px; margin-bottom: 24px;">
                                    <a href="{link}" style="background-color: #1a73e8; color: #ffffff; padding: 12px 24px; font-size: 14px; font-weight: 500; text-decoration: none; border-radius: 4px; display: inline-block;">View Work Item</a>
                                </div>

                                <!-- Arabic Section -->
                                <div style="border-top: 1px solid #dadce0; margin-top: 24px; padding-top: 24px; text-align: center; direction: rtl;">
                                    <p style="font-family: 'Segoe UI', Tahoma, Arial, sans-serif; font-size: 16px; color: #3c4043; margin: 5px 0;">يرجى ملاحظة أن المهمة أعلاه قد تجاوزت تاريخ الاستحقاق.</p>
                                    <p style="font-family: 'Segoe UI', Tahoma, Arial, sans-serif; font-size: 16px; color: #3c4043; margin: 5px 0;">يرجى اتخاذ الإجراء اللازم.</p>
                                </div>
                            </td>
                        </tr>

                        <!-- Footer -->
                        <tr>
                            <td style="background-color: #f8f9fa; padding: 24px; text-align: center; font-size: 12px; color: #70757a; border-top: 1px solid #dadce0;">
                                You received this email because you are assigned to this task.<br>
                                &copy; {datetime.now().year} RTA. All rights reserved.
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

    try:
        with smtplib.SMTP(SMTP_SERVER, 25) as smtp:
            smtp.send_message(msg)
            print(f"Overdue reminder sent for ID {item_id} to {to_email}")
            logging.info(f"SENT: ID {item_id} | Title: {title} | To: {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")
        logging.error(f"FAILED: ID {item_id} | To: {to_email} | Error: {e}")

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
        print(f"Debug: Failed to fetch identity from {url}: {e}")
    return None

def test_email_delivery(target_email):
    print(f"--- Running Email Test ---")
    print(f"Target: {target_email}")
    
    # Dummy Data for testing
    test_id = 99999
    test_title = "Test Task: Verify Email Layout and Delivery"
    test_date = datetime.now().strftime("%d/%m/%Y")
    test_project = "ITSD Test Project"
    
    send_overdue_email(target_email, test_id, test_title, test_date, test_project)
    print("Test email sent.")

def monitor_tasks():
    today = datetime.now().date()
    email_cache = {} # Cache to avoid repeated API calls for same user
    
    wiql_url = f"{ADO_URL}/{COLLECTION}/_apis/wit/wiql?api-version={API_VERSION}"
    query_payload = {
        "query": "SELECT [System.Id] FROM WorkItems WHERE [System.State] <> 'Closed'"
    }
    
    auth = HttpNegotiateAuth()
    
    try:
        response = requests.post(wiql_url, json=query_payload, auth=auth, verify=False) # verify=False if internal certs issues
        response.raise_for_status()
        work_items = response.json().get("workItems", [])
    except Exception as e:
        print(f"Error querying WIQL: {e}")
        return

    print(f"Checking {len(work_items)} items...")

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

            # Check if overdue (Due date must be strictly less than today)
            if due_date >= today:
                print(f"Skipping ID {item_id}: Not overdue (Due: {due_date})")
                continue
                
            # Get Assignee and Extract Email
            assigned_to = fields.get('System.AssignedTo')
            if not assigned_to:
                continue
            
            to_email = None
            unique_identifier = None

            if isinstance(assigned_to, dict):
                # If API returns an object
                unique_identifier = assigned_to.get('uniqueName')
                identity_url = assigned_to.get('url')
                
                # If uniqueName looks like an email, use it.
                if unique_identifier and '@' in unique_identifier and '.' in unique_identifier:
                    to_email = unique_identifier
                # Otherwise, fetch from Identity API using the URL provided in the object
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
                # If API returns a string: "Display Name <email@domain.com>"
                if '<' in assigned_to and '>' in assigned_to:
                    start = assigned_to.find('<') + 1
                    end = assigned_to.find('>', start)
                    to_email = assigned_to[start:end]
                elif '@' in assigned_to:
                    to_email = assigned_to

            # Validate extracted email
            if not to_email or '@' not in to_email:
                print(f"Could not resolve email for item {item_id}. Assignee: {unique_identifier or assigned_to}")
                continue

            # Prepare data for email
            title = fields.get('System.Title', 'No Title')
            project_name = fields.get('System.TeamProject', 'ITSD Task Tracker') # Default if missing
            formatted_date = due_date.strftime("%d/%m/%Y")
            
            # Call new function signature
            send_overdue_email(to_email, item_id, title, formatted_date, project_name)
            
        except Exception as e:
            print(f"Error processing item {item_id}: {e}")

if __name__ == "__main__":
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    
    monitor_tasks()
    # debug_list_fields()

    # Uncomment the line below to test email delivery to your personal address
    # test_email_delivery("your.personal.email@example.com")