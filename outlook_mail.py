import pythoncom
import win32com.client
import os

def send_outlook_mail(to_list, cc_list=None, subject="", body="", signature=True, attachments=None):
    """
    Send an email using Outlook desktop application.

    Parameters:
    - to_list (list): List of recipient email addresses.
    - cc_list (list, optional): List of CC email addresses. Default is None.
    - subject (str): Email subject.
    - body (str): Email body (supports HTML).
    - signature (bool): If True, appends the default Outlook signature.
    - attachments (list, optional): List of file names (files must be in the same folder as the script).

    Returns:
    - str: Success or failure message.
    """
    try:
        # Initialize COM library
        pythoncom.CoInitialize()

        # Get the current script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))

        # Initialize Outlook application
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = Mail Item

        # Set email recipients
        mail.To = ";".join(to_list)
        if cc_list:
            mail.CC = ";".join(cc_list)

        # Set email subject and body
        mail.Subject = subject
        if signature:
            # Get default Outlook signature
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Add()
            signature_text = doc.Content.Text  # Fetch the default Outlook signature
            doc.Close(False)
            word.Quit()

            mail.HTMLBody = body + "<br><br>" + signature_text  # Append signature
        else:
            mail.HTMLBody = body  # Without signature

        # Add attachments from the same folder as the script
        if attachments:
            for file_name in attachments:
                file_path = os.path.join(script_dir, file_name)  # Get full file path
                if os.path.exists(file_path):
                    mail.Attachments.Add(file_path)
                else:
                    print(f"Warning: Attachment not found - {file_path}")

        # Send the email
        mail.Send()
        return "Email sent successfully!"

    except Exception as e:
        return f"Error sending email: {str(e)}"

    finally:
        # Uninitialize COM library
        pythoncom.CoUninitialize()

# Example usage
to_list = ["mohammed.zeeshan1@sutherlandglobal.com", "mohammed.zeeshan1@sutherlandglobal.com"]
cc_list = ["mohammed.zeeshan1@sutherlandglobal.com"]
subject = "Test Email with Attachments"
body = "<h3>Hello,</h3><p>This is an automated email with attachments.</p>"

# Just specify file names, assuming they are in the same folder as the script
file_name = 'path_to_your_excel_file.xlsx'
attachments = [file_name]

# Send the email
send_mail = send_outlook_mail(to_list, cc_list, subject, body, signature=True, attachments=attachments)
print(send_mail)