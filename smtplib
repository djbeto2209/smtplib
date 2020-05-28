import csv, smtplib, ssl
import pandas as pd 


newhire = pd.read_csv('k:\\ITCAFTERMSKBCOPY.csv', delimiter = ",", skiprows=7,na_values="N/A")   #Open up the document and skip header
newhire = newhire.fillna(" ")
newhire["Email"] = "ithelp@peoplelinkgroup.com"     #Add a row 'Email' and add email address

mask = newhire["New IT Item"] == "New"      #brings back only new in "New IT Item column"

message = """\
Subject: {Effective_Date} -- {Employee} -- Kronos Term
From: kronos@peoplelinkgroup.com
To: ithelp@peoplelinkgroup.com
 


        Employee Name: {Employee}
        
        Title: {Title}

        Effective Date: {Effective_Date}

        Profit Center: {PC}

        Supervisor: {Supervisor}
        
        Manager 2 Name: {Manager_2_Name}

        Computer: {Computer}

        Avionte: {Avionte}

        
"""

from_address = "from address"
password = "password"

smtp = smtplib.SMTP("smtp.office365.com",587)
context = ssl.create_default_context()
with smtplib.SMTP("smtp.office365.com",587) as server:
    server.starttls(context=context)
    server.login(from_address, password)
    
    for i, r in newhire[mask].iterrows():     #This will unpack the contents of newhire
            server.sendmail(
                from_address,
                "ithelp@peoplelinkgroup.com",
                message.format(Employee=r["Employee Name"],
                   Effective_Date=r["Effective Date"],
                               PC=r["PC"],
                               Title=r["Title"],
                               Email=r["Email"], 
                               Supervisor=r["Supervisor Name"],
                               Computer =r["Laptop"],
                               Avionte = r["Avionte"],
                               GP=r["GP"],
                               Hotspot=r["Hotspot"],
                               Tablet=r["Tablet"],
                               Manager_2_Name=r["Manager 2 Name"]
                               
                )
            )

print('Term Emails were sent sucessfully!')
