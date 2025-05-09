Place mbox_to_csv.py and your .mbox file in the same directory. 

Run the script from the command line:

python3 mbox_to_csv.py inbox.mbox csv_name.csv --start-date YYYY-MM-DD --end-date YYYY-MM-DD

Make sure to adjust index.mbox to match whatever your .mbox file is named, csv_name.csv to whatever you want your output file to be named, and YYYY-MM-DD to the start and end dates you want to filter by.

The script will create a csv file in the same directory as the .mbox file and can be run multiple times with different start and end dates to create multiple csv files (make sure to adjust csv_name.csv each time when adjusting dates, to avoid overwriting the previous file).