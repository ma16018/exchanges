from exchanges import SendEmails

college = "Nuffield"
leg = "Mansfield"
sender = "Megan"
date = "Wednesday 23rd November"
signup = ""  # add link
away_dets = ("formal", "10", "6:30pm", "7:30pm")  # dress, price, start, dinner_time
test = SendEmails(college, leg, sender, date, away_details=away_dets, link=signup)

# Email whole MCR about the event
test.email_sign_up

# Send confirmation to selected diners
test.email_selected(attendents=13)

# Send anything else to selected diners
main = f"""Hello"""
test.email_followup(main)
