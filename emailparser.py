import extract_msg
import sqlite3
import os, re
import win32com.client
import datetime
from datetime import datetime

def main():
    # Set up database
     db = setup()
     db.row_factory = lambda cursor, row: row[0]
     c = db.cursor()


    # define folder path with .msg files
    folder_path = r"C:/Users/gahane/Documents/Newfolder/Noris/"
    
    os.chdir(folder_path)
    #alternative to import .msg files to the script
    #for filename in os.listdir(folder_path):
        #new_name = re.sub('[^A-Za-z0-9]+', '', filename)
        #os.rename(os.path.join(folder_path,filename), os.path.join(folder_path,new_name+'.msg'))

    # Initialise & populate list of emails
    email_list = [file for file in os.listdir(folder_path) if file.endswith(".msg")]

    #defining all regex groups
    #regex for ticket number
    regex_ticket_no = re.compile(r"Ticket number\s*:\s*(\d+)")
    regex_ticketnummer = re.compile(r"Ticketnummer\s*([a-zA-Z0-9]+)")
    
    #regex for start time
    regex_start_date = re.compile(r"Start time\s*:\s*([0-9]{4}-[0-9]{2}-[0-9]{2}.*[0-9]+:[0-9]+)")
    regex_time_start = re.compile(r"Time Start\s*:\s*((0?[1-9]|[12][0-9]|3[01])\s[a-zA-Z]+\s[0-9]+\s+(0?[0-9]|1[0-9]|2[0-3]):[0-9]+)")
    regex_Start_Time = re.compile(r"Start Time\s*:\s*(([0-9]+(/[0-9]+)+).*[0-9]+:[0-9]+)")
    regex_start_time1 = re.compile(r"Time Start\s*:\s*([0-9]+\.\s[a-zA-Z]+\s[0-9]+\s+[0-9]+:[0-9]+)")

    
    #regex for end time
    regex_End_Time = re.compile(r"End Time\s*:\s*(([0-9]+(/[0-9]+)+).*[0-9]+:[0-9]+)")
    regex_end_date = re.compile(r"End time\s*:\s*([0-9]{4}-[0-9]{2}-[0-9]{2}.*[0-9]+:[0-9]+)")
    regex_time_end = re.compile(r"Time End\s*:\s*((0?[1-9]|[12][0-9]|3[01])\s[a-zA-Z]+\s[0-9]+\s+(0?[0-9]|1[0-9]|2[0-3]):[0-9]+)")
    regex_end_time1 = re.compile(r"Time End\s*:\s*([0-9]+\.\s[a-zA-Z]+\s[0-9]+\s+[0-9]+:[0-9]+)")
    regex_end_time2 = re.compile(r"Time End\s*:\s*((0?[1-9]|[12][0-9]|3[01]).*\s[a-zA-Z]+\s[0-9]+\s+[0-9]+:[0-9]+)")
    
    
    #regex for ticket type, class, status, announcement 
    regex_ticket_type = re.compile(r"Ticket type\s*:\s*([a-zA-Z]+.*[a-zA-Z]+)")
    regex_tickettype = re.compile(r"Type\s*:\s*([a-zA-Z]+.*[a-zA-Z]+)")
    regex_ticket_class = re.compile(r"Ticket class\s*:\s*([a-zA-Z]+)")
    
    regex_ticket_status = re.compile(r"Ticketstatus\s*:\s*([a-zA-Z]+)")
    regex_ticket_status1 = re.compile(r"FW: Announcement\s*([a-zA-Z]+\s)")
    
    regex_impacttoservice = re.compile(r"Impact to Service\s*:\s([a-zA-Z]+( [a-zA-Z]+)+).*([a-zA-Z]+([a-zA-Z]+)+)")
    regex_message = re.compile(r"Message\s*: (?<=>)([\w\s]+)(?=<\/)")
    regex_summary = re.compile(r"Summary\s*: +([a-zA-Z]+( [a-zA-Z]+)+)")
    #new
    #regex_summary1 = re.compile(r"Summary\s*:\s+([a-zA-Z]+[0-9]+-[a-zA-Z]+.*\s_\s[a-zA-Z]+\s- ([a-zA-Z]+( [a-zA-Z]+)+)\s+)")
    
    #regex for site location 
    regex_site_location = re.compile(r"Site Location\s*:\s+([a-zA-Z]+[0-9]+\s[a-zA-Z]+\s[a-zA-Z]+\s)")
    regex_site_location1 = re.compile(r"Site Location\s*\s+([a-zA-Z]+[0-9]+\s[a-zA-Z]+\s[a-zA-Z]+\s)")
    regex_site_location2 = re.compile(r"Site Location\s*:\s+([a-zA-Z]+[0-9]+/[0-9]+\s[a-zA-Z]+\s[a-zA-Z]+)")
    
    
    # Iterate through every email
    for i, _ in enumerate(email_list):
        
        
        msg_cont = extract_msg.Message(email_list[i])
        #msg_message= outlook.OpenSharedItem(os.path.join(folder_path,email_list[i]))
        msg_message = msg_cont.body
    
        msg_subject = msg_cont.subject
        msg_recieved = msg_cont.date
        msg_recieved = datetime.strptime(msg_recieved, '%a, %d %b %Y %H:%M:%S %z').strftime('%Y-%m-%d %I:%M')
        
        
        # for ticket number
        ticket_no = re.search(regex_ticket_no, msg_message)
        ticketnummer = re.search(regex_ticketnummer, msg_message)
        
        if ticket_no:
            ticket_no = ticket_no.group(1)
        elif ticketnummer:
            ticket_no = ticketnummer.group(1)
        else:
            ticket_no = None
            
            
        # for ticket start date and time
        ticket_start_date = re.search(regex_start_date, msg_message)
        ticket_time_start = re.search(regex_time_start, msg_message)
        ticket_Start_Time = re.search(regex_Start_Time, msg_message)
        ticket_start_time1 = re.search(regex_start_time1, msg_message)
#         ticket_start_time2 = re.search(regex_start_time2, msg_message)
        
        if ticket_start_date:
            ticket_start_date = ticket_start_date.group(1)
            
        elif ticket_time_start:
            ticket_start_date = ticket_time_start.group(1)
            ticket_start_date = datetime.strptime(ticket_start_date, '%d %B %Y %I:%M').strftime('%Y-%m-%d %I:%M')
            
        elif ticket_Start_Time:
            ticket_start_date = ticket_Start_Time.group(1)
            ticket_start_date = datetime.strptime(ticket_start_date, '%d/%m/%Y %H:%M').strftime('%Y-%m-%d %I:%M')
        
        elif ticket_start_time1:
            ticket_start_date = ticket_start_time1.group(1)
        else:
            ticket_start_date = None

            
        # for ticket end date and time
        ticket_end_date = re.search(regex_end_date, msg_message)
        ticket_time_end = re.search(regex_time_end, msg_message)
#         ticket_End_Time = re.search(regex_End_Time, msg_message)
        ticket_end_time1 = re.search(regex_end_time2, msg_message)
        
        
        if ticket_end_date:
            ticket_end_date = ticket_end_date.group(1)
            
        elif ticket_time_end:
            ticket_end_date = ticket_time_end.group(1)
            ticket_end_date = datetime.strptime(ticket_end_date, '%d %B %Y %H:%M').strftime('%Y-%m-%d %H:%M')
            
#         elif ticket_End_Time:
#             ticket_end_date = ticket_End_Time.group(1)
#             ticket_end_date = datetime.strptime(ticket_end_date, '%d/%m/%Y %H:%M').strftime('%Y-%m-%d %I:%M')
            
        elif ticket_end_time1:
            ticket_end_date = ticket_end_time1.group(1)
            ticket_end_date = datetime.strptime(ticket_end_date, '%d. %B %Y %H:%M').strftime('%Y-%m-%d %I:%M')
    
        else:
            ticket_end_date = None

            
        # for ticket type
        ticket_type = re.search(regex_ticket_type, msg_message)
        tickettype = re.search(regex_tickettype, msg_message)
        
        if ticket_type:
            ticket_type = ticket_type.group(1)
            
        elif tickettype:
            ticket_type = tickettype.group(1)
            
        else:
            ticket_type = None

            
        # for ticket class
        ticket_class = re.search(regex_ticket_class, msg_message)
        
        if ticket_class:
            ticket_class = ticket_class.group(1) 
        
        else:
            ticket_class = None
            
            
        # for status or announcment
        ticket_status = re.search(regex_ticket_status, msg_message)
        ticket_status1 = re.search(regex_ticket_status1, msg_subject)
        
        if ticket_status:
            ticket_status = ticket_status.group(1)  
        
        elif ticket_status1:
            ticket_status = ticket_status1.group(1)
            
        else:
            ticket_status = None
            
            
        # for impact to service
        impacttoservice = re.search(regex_impacttoservice, msg_message)
        
        if impacttoservice:
            impacttoservice = impacttoservice.group(1)
        else:
            impacttoservice = None
            
            
        # for site location    
        site_location = re.search(regex_site_location, msg_message)
        site_location1 = re.search(regex_site_location1, msg_message)
        site_location2 = re.search(regex_site_location2, msg_message)
        if site_location:
            site_location = site_location.group(1)
        elif site_location1:
            site_location = site_location1.group(1)
        elif site_location2:
            site_location = site_location2.group(1)
        else:
            site_location = None
        
        
        # summary
        summary = re.search(regex_summary, msg_message)
        
        if summary:
            summary = summary.group(1)
        else:
            summary = None
        
        
        # message    
        message = re.search(regex_message, msg_message)
        
        if message:
            message = message.group(1)
        else:
            message = None
                    
         sql_insert(db, ticket_no, msg_recieved, ticket_start_date, ticket_end_date, ticket_type, ticket_class, ticket_status, impacttoservice, summary, site_location, message)
#         print(ticket_no, ticket_type, ticket_class, ticket_status, impacttoservice, summary, site_location, message)

#         print(ticket_no, msg_subject)

 def sql_insert(db, ticket_no, msg_recieved, ticket_start_date, ticket_end_date, ticket_type, ticket_class, ticket_status, impacttoservice, summary, site_location, message):

     ticketnos = db.execute('SELECT ticket_no FROM PARSER_DB2').fetchall()
     # print(ticketnos)
     if ticket_no in ticketnos:
         # print('Ticket already exists and updated')
         db.execute("UPDATE PARSER_DB2 SET ticket_status = ? WHERE ticket_no = ?", (ticket_status, ticket_no))
         db.commit()
     else:
         db.execute("INSERT INTO PARSER_DB2 (ticket_no, msg_recieved, ticket_start_date, ticket_end_date, ticket_type, ticket_class, ticket_status, impacttoservice, summary, site_location, message) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", 
         (ticket_no, msg_recieved, ticket_start_date, ticket_end_date, ticket_type, ticket_class, ticket_status, impacttoservice, summary, site_location, message))
         db.commit()

     '''insert or replace into Book (ID, Name, TypeID, Level, Seen) values
 ((select ID from Book where Name = "SearchName"), "SearchName", ...);'''
    
 def setup():
     # Create & connect to database, add file path below
     db = sqlite3.connect("C:/Users/gahane/Documents/Newfolder/Intexterion/email_parser.db")

     # Create empty tables
     db.execute("""
     CREATE TABLE IF NOT EXISTS "PARSER_DB2" (
     ID INTEGER PRIMARY KEY AUTOINCREMENT,
     "ticket_no" VARCHAR(60),
     "msg_recieved" TEXT,
     "ticket_start_date" DATE,
     "ticket_end_date" DATE,
     "ticket_type" TEXT,
     "ticket_class" TEXT,
     "ticket_status" TEXT,
     "impacttoservice" TEXT,
     "summary" TEXT,
     "site_location" TEXT, 
     "message" TEXT)
      """)
    
     db.commit()

     return db
    
main()
 