#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import csv
from time import sleep
import win32com.client as client
import pathlib


pdf_path = pathlib.Path("Myopia Focus - Become a Contributor.pdf")
pdf_absolute = str(pdf_path.absolute())

with open("test list.csv", newline="") as f:
    reader = csv.reader(f)
    distro = [row for row in reader]
    
chunks = [distro[x:x+30] for x in range(0, len(distro), 30)]



outlook = client.Dispatch("Outlook.Application")
for chunk in chunks:
    for name, address in chunk:
        message = outlook.CreateItem(0)
        message.To = address
        message.Subject = "Myopiafocus.org - Myopia Management Specialist"
        message.SentOnBehalfOfName = "MyopiaFocus.org"
        message.Attachments.Add(pdf_absolute)
        message.HTMLBody = html_body
        message.Send()
    
    sleep(60)

