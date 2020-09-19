"""
###### Birthday Wisher Programme Using Python 

NorthFox Py Developer https://bit.ly/NFGTwitter

Team : NorthFox Py Developers 
https://github.com/princu09/pybdaywisher 

The id password you enter in this file is 100% Safe & Secure.
Join Our Telegram Group For More Information and data http://t.me/ProgHub09

"""


# ======================== Main Code Starts Here ========================


import pandas as pd
import datetime
import smtplib

# ======================== Enter you mail Details Here ========================

Email_ID = '' # Your Email
Email_PSWD = '' # Your Password


def sendEmail(to , sub , msg):
    print(f"Email to {to} send with Subject : {sub} and massage {msg}")
    
    # ======================== Gmail SMTP Server ========================
    s = smtplib.SMTP('smtp.gmail.com' , 587)
    
    # ======================== STMP Server Connect or Start =============
    s.starttls()

    # ======================== Login to Email ===========================
    s.login(Email_ID , Email_PSWD)
    
    # ======================== Send Mail to User ========================
    s.sendmail(Email_ID , to , f"Subject : {sub}\n\n{msg}")
    
    # ======================== Mail Quit ================================
    s.quit()

if __name__ == "__main__":

    rd = pd.read_excel("data.xlsx") #read dara =  rd
    #print(rd) # When You want to check data in inserted successfully uncomment print function
    
    today = datetime.datetime.now().strftime("%d-%m")
    # print(today) #Print This year For send Email


    yearNow = datetime.datetime.now().strftime("%Y")
    writeYr = []
    # This will add the current year to the excel sheet so that the mail does not go again

    for index , item in rd.iterrows():
        # print(index , item['Birthday'])

        birthday = item['Birthday'].strftime("%d-%m")
        # Get today date for Send Birthday wish mail
        # print(birthday)

        if (today  == birthday) and yearNow not in str(item['Year']):
            sendEmail(item['Email'] , "Happy Birthday" , item['Dailogue'])
            writeYr.append(index)

        if (today  == birthday) and yearNow in str(item['Year']):
            print(f"Your wish mail alredy send to {item['Name']}")
    
    for i in writeYr:
        yr = rd.loc[i , 'Year']
        rd.loc[i , 'Year'] = str(yr) + ',' + str(yearNow)

    rd.to_excel('data.xlsx', index=False)