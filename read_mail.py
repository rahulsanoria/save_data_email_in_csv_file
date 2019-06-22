import email
import imaplib 
from bs4 import BeautifulSoup
import os
import mimetypes
import csv
import io, json
import pandas as pd

username = 'your_email_id'
password = 'password'

mail = imaplib.IMAP4_SSL("imap.gmail.com") 
mail.login(username, password)

mail.select("inbox")

result, data = mail.uid('search', None, "ALL") 

inbox_item_list = data[0].split()
result2, email_data = mail.uid('fetch', inbox_item_list[-1], '(RFC822)')
raw_email = email_data[0][1]
email_message = email.message_from_string(raw_email)
to_ = email_message['To']
from_ = email_message['From']
subject_ = email_message['Subject']
date_ = email_message['date']

counter = 1

for part in email_message.walk():
    if part.get_content_maintype() == "multipart":
        continue 
    filename = part.get_filename()
    content_type = part.get_content_type()
    if not filename:
        ext = mimetypes.guess_extension(content_type)
        if not ext:
            ext = '.bin'
        if 'text' in content_type:
            ext = '.docx'
        elif 'html' in content_type:
            ext = '.html'
        filename = 'msg-part-%08d%s' %(counter, ext)
    counter += 1

if "plain" in content_type:
	body = part.get_payload()
elif "html" in content_type:
	html_ = part.get_payload()
	soup = BeautifulSoup(html_ , "html.parser")
	text = soup.get_text()
	body = text	  
else:
	body = content_type

p = {"FROM" : from_ , "TO" : to_ , "DATE" : date_ , "SUBJECT" : subject_ , "BODY" : body }
d = json.dumps(p)

f = open('append.csv' , 'a')
csv_file = csv.writer(f)
for item in d:
    csv_file.writerow(item)
    print item
f.close()

# csvo = open('append.csv', 'a')
# csvwriter = csv.writer(csvo)
# count = 0

# for data in p:
#     if count==0:
#         header = data.keys()
#         csvwriter.writerow(header)
#         count += 1
#     csvwriter.writerow(data.values())
# csvo.close()        


    

# with io.open('append.json', 'w', encoding='utf-8') as f:
#   f.write(json.dumps(data, ensure_ascii=False))

# df = pd.read_json("append.json")
# print df
# df.to_csv('append.csv')

