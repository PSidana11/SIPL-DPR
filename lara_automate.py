import win32com.client  
import os
import subprocess
from datetime import datetime 

save_path = r"C:\Users\E01412\Desktop\REPORTS\DAILY\LARA"
sender_name = "SIPL lara"
target_script = ["lara.py", "lara_finance.py"]  # Script to run after download
 
today = datetime.now().date()
today_str = today.strftime("%Y-%m-%d")

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 = inbox    
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)  # Sort by latest
 
downloaded = False

for message in messages:
    try:
        received_date = message.ReceivedTime.date()
        actual_sender = message.SenderName
        print(f"Checking email from: {actual_sender} at {received_date}")

        if actual_sender.strip().lower() == sender_name.strip().lower() and received_date == today:
            for attachment in message.Attachments:
                if attachment.FileName.endswith(".xlsx"):
                    file_name = f"lara_{today_str}.xlsx"
                    file_path = os.path.join(save_path, file_name)
                    attachment.SaveAsFile(file_path)
                    print(f"Downloaded and saved as: {file_path}")
                    downloaded = True
                    break
            if downloaded:
                break
    except Exception as e:
        print(f"Error: {e}")

if downloaded:
    print(f"Running {target_script}...")
    subprocess.run(["python", target_script], check=True)
    print(f"Finished {target_script}")      
else:
    print("No matching email found for today.")
    