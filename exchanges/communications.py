import numpy as np
from random import sample
import pandas as pd
import os
import win32com.client
from datetime import datetime, date, timedelta


class SendEmails:

    def __init__(self, college, leg, sender, date, away_details=None):
        self.college = college
        self.leg = leg
        self.sender = sender
        self.role = "MCR Social Secretary" if sender=="Megan" else "MCR Wine and Dine Rep" 
        self.date = date
        self.dress, self.price, self.start, self.dinner_time = away_details if away_details else (0, 0, 0, 0)

    def get_message_html(self, main):
        """Get string of the HTML code to use in the email."""
        message = f"""
            <html>
                <head></head>
                    <body>
                        <p>Dear all,</p>

                        <p>{main}</p>

                        <p>Best wishes,<br>
                        {self.sender}<br>
                        {self.role}</p><br>
                    </body>
                </html>
                """
        return message
    
    @property
    def find_photos(self):
        """Collect the paths of photos in a list."""
        path = os.getcwd() + "\\exchanges\\Pictures"
        print(path)
        pictures = []
        for root, dirs, files in os.walk(path):
            for file in files:
                if self.college in file:
                    pictures.append(os.path.join(root, file))
        return pictures

    def send_email(self, emails, main):
        """Interacts with Outlook app to send emails. Input is the required email addresses and main
        body of the message. If the leg is away it will also attach photos."""
        print("Getting message")
        message = self.get_message_html(main)
        print("Open outlook")
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = emails 
        print(f"Sending to {emails}.")
        mail.Subject = f'Exchange Formal Dinner - {self.college} ({self.leg}\'s leg)'
        pictures = self.find_photos if self.leg != "Mansfield" else []
        for pic in pictures:
            attachment  = pic
            mail.Attachments.Add(attachment)
        mail.HTMLBody = message
        mail.Send()
        print(f"Sent.")

    @property
    def email_sign_up(self):
        """Get messages and email for sign up and send off."""
        if self.leg == "Mansfield":
            message = f"""As our leg of the exchange formal dinner, 15 guests from {self.college} will be joining us 
                       at Mansfield on {self.date}. If you would like to be one of the 15 
                       Mansfield students to attend this formal (half-price) and sit with our guests, please 
                       fill out this form (make this a link)."""
        else:
            message = f"""We have an upcoming exchange dinner with {self.college} College. The leg at their college 
                       will be on {self.date}, and cost {self.price}. The drinks reception will begin at {self.start}, with dinner 
                       starting at {self.dinner_time}. The dress code is {self.dress}. To sign up for this opportunity, 
                       fill out this form (make this a link) within 24 hours of this email. Those that have been 
                       selected to attend will be informed by 10pm tomorrow."""
            
        emails = "megan.a05@icloud.com"  # mansfield-mcr-announce@maillist.ox.ac.uk
        self.send_email(emails, message)

    def find_excel(self, folder = "\\SignUps"):
        """Get path name of an excel file in the specified folder. E.g. 'SignUps' folder for
        emails of everyone who signed up or 'Chosen' folder for emails of people who are confirmed 
        to attend."""
        path = os.getcwd() + "\\exchanges" + folder
        for root, dirs, files in os.walk(path):
            for file in files:
                if self.college in file:
                    print(f"Using file: {file}")
                    path = os.path.join(root, file)
                    return pd.read_excel(path)
        return "No excel file found"

    def chose_attendents(self, attendents=15):
        """Randomly select 15 people to attend"""
        data = self.find_excel()
        n = len(data)
        rows = sample(range(n), attendents)
        winners = data.iloc[rows]
        path = os.getcwd()  + "\\exchanges\\Chosen" + f"\\{self.college}_selection.xlsx"
        winners.to_excel(path)
        print(f"Found {len(winners)} people")
        emails = "; ".join(winners.Email)
        return emails

    def email_selected(self, attendents=15):
        """Chose attendents, chose the message and send off the details."""
        if self.leg == "Mansfield":
            message = f"""We would like to invite you to join our exchange dinner at Mansfield on {self.date} with
                        {self.college} college. We will meet in our MCR at 7pm for a drinks reception. During the dinner,
                        you will sit with 15 guests from {self.college}. We will retire after the dinner to the MCR for more
                        drinks. Half the price of this formal dinner will be charged to your battels. Please note that you do <b>not<\b> 
                        need to book this formal dinner through the booking system."""
        else:
            message = f"""We would like to invite you to join our exchange dinner at <b>{self.college} college</b> on 
                        <b>{self.date}</b>. Please transfer {self.price} pounds (half-price) to our MCR account no later than 
                        <b>48 hours before</b> the formal. Here are the bank details of our MCR account:</p>

                        <p>Name: Mansfield College MCR<br>
                        Sort code: 40-35-34<br>
                        Account number: 74334922</p>

                        <p>The details for this dinner are as follows: Arrive at {self.college} at {self.start}. We will be 
                        meet at the gate and walked over to their MCR for a drinks reception. Dinner starts at 
                        {self.dinner_time}, where dress code is {self.dress}. After the dinner, we will head back to their 
                        MCR for digestifs."""
        emails = self.chose_attendents(attendents)
        self.send_email(emails, message)
        

    def find_winner_emails(self):
        """Find the emails of the people who are confirmed to go to the event, e.g. in 
        case we need to send them event details."""
        data = self.find_excel(folder = "Chosen")
        emails = "; ".join(data.Email)
        return emails

    def email_followup(self, message):
        """Send any extra details to the winners."""
        emails = self.find_winner_emails()
        self.send_email(emails, message)
