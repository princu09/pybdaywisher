### Python BirthDay Wisher

##### 📫 Connect with me here:<br />
 <br />
 <p>
  <a href="https://www.instagram.com/princu09">
    <img src="https://img.shields.io/badge/princu.09-386938188?style=flat&logo=instagram&color=black">
  </a> &nbsp; 
  <a href="https://twitter.com/princu09">
    <img src="https://img.shields.io/badge/@princu09-30302f?style=flat&logo=twitter&color=black">
  </a>&nbsp; 
  <a href="https://github.com/princu09">
    <img src="https://img.shields.io/badge/@princu09-30302f?style=flat&logo=github&color=black">
  </a>&nbsp;
    <a href="https://www.t.me/proghub09">
    <img src="https://img.shields.io/badge/ProgHub09-386938188?style=flat&logo=telegram&color=black">
  </a>
</p>


<br>

###### <p style="font-size:12px"> Note : One Excel Sheet Also Including This Git SO Don't Changes in Excel Sheet Because It's Connect into Python File . After You changes in Excel Sheet python file not working. Only Insert B'day and Some Data You send to Birthday Boy/Girl</p>

###### Excel File Link : <a href="data.xlsx">Data.xlsx</a>

##### Main File Code :

```
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
```
