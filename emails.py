import win32com.client as win32

# Create Outlook application instance
outlook = win32.Dispatch('Outlook.Application')

# Email subject
subject = "Research Follow-Up: Share Your Thoughts on Feedback Approaches"

# HTML body (formatted)
html_body = """
<html>
<head></head>
<body style="font-family: Calibri, sans-serif; font-size: 11pt; color: #333;">
  <p>Dear Participant,</p>

  <p>I hope you are doing well. Thank you again for taking part in my research on maritime training.</p>

  <p>
    As a next step, I’m now collecting participants’ evaluations of the different feedback approaches used during the training.
    Your input will be instrumental in helping us understand what works best and how we can improve personalized training for seafarers.
  </p>

  <p>
    I would be very grateful if you could take a few minutes to complete the short assessment at the link below:
  </p>

    <p>
    
    https://stratheng.eu.qualtrics.com/jfe/form/SV_9Hp2WcZIERs6dP8
    
    </p>

  <p>
    <a href="https://stratheng.eu.qualtrics.com/jfe/form/SV_9Hp2WcZIERs6dP8" target="_blank" style="color: #0072C6; font-weight: bold;">
      Complete the Assessment Form
    </a>
  </p>

  <p>
    If you have any questions or additional comments, feel free to get in touch.
    Thank you once again for your valuable support.
  </p>

  <p>Best regards,<br><br>
     <strong>Furkan Tornaci</strong><br>
     PhD Researcher<br>
     University of Strathclyde<br>
     <a href="mailto:furkan.tornaci@strath.ac.uk">furkan.tornaci@strath.ac.uk</a>
  </p>
</body>
</html>
"""

# List of email recipients
recipients = [
    "shanaakhaledfiras@gmail.com",
    "anas_ali_ibrahim@yahoo.com",
    "Mob20722467@gmail.com",
    "abdulla3930099@gmail.com",
    "Waleedsalah1994@yahoo.com",
    "abdullatif.yonso@gmail.com",
    "can.tas@solent.ac.uk",
    "k_29@live.com",
    "etulloch@me.com",
    "nick.haslam@brookesbell.com",
    "mehmetali_bahce@yahoo.com",
    "yahyapekdas@hotmail.com",
    "cap_hany37@hotmail.com",
    "Heshamabdushkour@gmail.com"
]

# === SETTINGS ===
preview_before_send = False  # Set to True to preview instead of sending
log_sent_emails = True

# === SEND EMAILS ===
for email in recipients:
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = subject
    mail.HTMLBody = html_body

    if preview_before_send:
        mail.Display()
    else:
        mail.Send()

    if log_sent_emails:
        print(f"Email sent to: {email}")

print("✅ All emails processed.")
