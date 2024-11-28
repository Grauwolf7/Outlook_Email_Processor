import win32com.client
import datetime
import json
from pathlib import Path
import os

def get_days_input():
    """Prompts the user for the number of days to process"""
    while True:
        try:
            days = int(input("How many days back would you like to retrieve emails for? (0 = today only): "))
            if days < 0:
                print("Please enter a positive number or 0.")
                continue
            return days
        except ValueError:
            print("Please enter a valid number.")

def create_directory_structure(date):
    """Creates the folder structure for email summaries"""
    base_dir = Path(__file__).parent
    summaries_dir = base_dir / "Email Summaries"
    date_dir = summaries_dir / date.strftime("%Y-%m-%d")
    date_dir.mkdir(parents=True, exist_ok=True)
    return date_dir

def find_outlook():
    """Finds and connects to the active Outlook instance"""
    try:
        outlook = win32com.client.GetActiveObject("Outlook.Application")
    except:
        try:
            outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        except:
            outlook = win32com.client.Dispatch("Outlook.Application")
    return outlook

def should_include_email(sender_email):
    """Checks if the email should be included in the summary"""
    excluded_addresses = {
        # Add excluded email addresses here
    }
    return sender_email.lower() not in {addr.lower() for addr in excluded_addresses}

def get_emails_for_date(target_date):
    """Retrieves emails for a specific date"""
    outlook = find_outlook()
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)
    
    start_date = target_date.replace(hour=0, minute=0, second=0, microsecond=0)
    end_date = target_date.replace(hour=23, minute=59, second=59, microsecond=999999)
    
    emails = []
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    
    print(f"Searching inbox: {inbox.Name}")
    filtered_count = 0
    
    for message in messages:
        try:
            if start_date.date() <= message.ReceivedTime.date() <= end_date.date():
                if should_include_email(message.SenderEmailAddress):
                    email_data = {
                        "subject": message.Subject,
                        "sender": message.SenderName,
                        "sender_email": message.SenderEmailAddress,
                        "body": message.Body,
                        "received": message.ReceivedTime.strftime("%H:%M:%S")
                    }
                    emails.append(email_data)
                else:
                    filtered_count += 1
            elif message.ReceivedTime.date() < start_date.date():
                break
        except Exception as e:
            print(f"Warning: Could not process email: {str(e)}")
            continue
    
    print(f"Filtered emails for {target_date.strftime('%Y-%m-%d')}: {filtered_count}")
    return emails

def save_emails_to_file(emails, output_dir):
    """Saves the emails to a JSON file"""
    json_file = output_dir / "emails.json"
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(emails, f, ensure_ascii=False, indent=4)
    return json_file

def create_summary_prompt(emails, date):
    """Creates a prompt for ChatGPT"""
    prompt = f"""Please summarize the following emails from {date.strftime('%m/%d/%Y')}. 
    Group them by topic and highlight important information:

    """
    for i, email in enumerate(emails, 1):
        prompt += f"\nEmail {i}:\n"
        prompt += f"From: {email['sender']} ({email['sender_email']})\n"
        prompt += f"Subject: {email['subject']}\n"
        prompt += f"Received at: {email['received']}\n"
        prompt += f"Content: {email['body']}\n"
        prompt += "-" * 50

    return prompt

def process_date(date):
    """Processes emails for a specific date"""
    print(f"\nProcessing emails for {date.strftime('%m/%d/%Y')}...")
    
    # Create directory structure for this date
    output_dir = create_directory_structure(date)
    print(f"Output directory created: {output_dir}")
    
    # Get emails for this date
    emails = get_emails_for_date(date)
    
    if not emails:
        print(f"No emails found for {date.strftime('%m/%d/%Y')}.")
        return
        
    print(f"{len(emails)} relevant emails found.")
    
    # Save emails to file
    json_file = save_emails_to_file(emails, output_dir)
    print(f"Emails have been saved to {json_file}")
    
    # Create and save prompt
    prompt = create_summary_prompt(emails, date)
    prompt_file = output_dir / "email_summary_prompt.txt"
    with open(prompt_file, 'w', encoding='utf-8') as f:
        f.write(prompt)
    print(f"Prompt has been saved to {prompt_file}")

def main():
    try:
        # Ask for number of days
        days = get_days_input()
        
        # Process each day
        for i in range(days + 1):
            date = datetime.datetime.now() - datetime.timedelta(days=i)
            process_date(date)
            
        print("\nProcessing completed!")
        print(f"You can find the files in the 'Email Summaries' folder.")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        print(f"Error type: {type(e).__name__}")
        import traceback
        print("Details:")
        print(traceback.format_exc())

if __name__ == "__main__":
    main()