from exchanges import SendEmails
import numpy as np
import pandas as pd

college = "Nuffield"
leg = "Mansfield"
sender = "Megan"
date = "Wednesday 23rd November"

test = SendEmails(college, leg, sender, date)
test.email_selected(attendents=13)
