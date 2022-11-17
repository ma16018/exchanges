from exchanges import SendEmails

college = "Nuffield"
leg = "Mansfield"
sender = "Megan"
date = "Wednesday 23rd November"
signup = ""  # add link
test = SendEmails(college, leg, sender, date, link=signup)

# Email whole MCR about the event
test.email_sign_up

# Send confirmation to selected diners
test.email_selected(attendents=13)

# Send anything else to selected diners
main = f"""Hello"""
test.email_followup(main)
