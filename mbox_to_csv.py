import mailbox
import csv
import datetime
from email.utils import parsedate_to_datetime
import os
import re
import sys

def extract_email_body(msg):
    """Extract the email body from the message."""
    try:
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                
                # Skip attachments
                if "attachment" in content_disposition:
                    continue
                    
                if content_type == "text/plain":
                    payload = part.get_payload(decode=True)
                    if payload:
                        try:
                            return payload.decode(errors='replace')
                        except:
                            return str(payload)
            return ""
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                try:
                    return payload.decode(errors='replace')
                except:
                    return str(payload)
            return ""
    except Exception as e:
        print(f"Error extracting email body: {e}")
        return "[Error extracting email body]"

def convert_mbox_to_csv(mbox_file, csv_file, start_date=None, end_date=None):
    """
    Convert emails from an mbox file to CSV, filtering by date range.
    
    Args:
        mbox_file (str): Path to the mbox file
        csv_file (str): Path to the output CSV file
        start_date (datetime): Start date for filtering (inclusive)
        end_date (datetime): End date for filtering (inclusive)
    """
    # Ensure the mbox file exists
    if not os.path.exists(mbox_file):
        print(f"Error: {mbox_file} does not exist")
        return
    
    # Open the mbox file
    print(f"Opening mbox file: {mbox_file}")
    try:
        mbox = mailbox.mbox(mbox_file)
        total_emails = len(mbox)
        print(f"Total emails in mbox: {total_emails}")
    except Exception as e:
        print(f"Error opening mbox file: {e}")
        return
    
    # Prepare CSV file
    with open(csv_file, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Date', 'From', 'To', 'Subject', 'Body', 'Message-ID']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        
        # Counter for progress tracking
        processed_emails = 0
        included_emails = 0
        
        print(f"Filtering from {start_date} to {end_date}")
        
        # Process each email
        for message in mbox:
            processed_emails += 1
            
            if processed_emails % 100 == 0:
                print(f"Processed {processed_emails}/{total_emails} emails...")
            
            # Get the date of the email
            date_str = message['Date']
            if not date_str:
                continue  # Skip if no date
                
            try:
                # Parse the date from the email
                date = parsedate_to_datetime(date_str)
                
                # Convert to naive datetime for comparison
                email_date_naive = date.replace(tzinfo=None)
                
                # Filter by date if start_date and end_date are provided
                if (start_date and email_date_naive < start_date) or (end_date and email_date_naive > end_date):
                    continue
                    
                # Extract email data
                included_emails += 1
                
                # Get the email body
                body = extract_email_body(message)
                
                # Clean the body (remove extra spaces and newlines)
                body = re.sub(r'\s+', ' ', body).strip() if body else ""
                
                # Write to CSV
                writer.writerow({
                    'Date': date.strftime('%Y-%m-%d %H:%M:%S'),
                    'From': message['From'] or '',
                    'To': message['To'] or '',
                    'Subject': message['Subject'] or '',
                    'Body': body,
                    'Message-ID': message['Message-ID'] or ''
                })
                
            except (TypeError, ValueError) as e:
                print(f"Error processing email: {e}")
                continue  # Skip if date can't be parsed
    
    print(f"Conversion complete. {included_emails} emails within date range exported to {csv_file}")

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='Convert Gmail mbox file to CSV with date filtering')
    parser.add_argument('mbox_file', help='Path to the mbox file')
    parser.add_argument('csv_file', help='Path to the output CSV file')
    parser.add_argument('--start-date', help='Start date (YYYY-MM-DD)', required=False)
    parser.add_argument('--end-date', help='End date (YYYY-MM-DD)', required=False)
    
    args = parser.parse_args()
    
    # Convert dates to datetime objects
    start_date = None
    end_date = None
    
    if args.start_date:
        start_date = datetime.datetime.strptime(args.start_date, '%Y-%m-%d')
    
    if args.end_date:
        # Set the end date to the end of the day (23:59:59)
        end_date = datetime.datetime.strptime(args.end_date, '%Y-%m-%d')
        end_date = end_date.replace(hour=23, minute=59, second=59)
    
    convert_mbox_to_csv(args.mbox_file, args.csv_file, start_date, end_date)