from firebase_admin import db, credentials, firestore, initialize_app
from email.message import EmailMessage
import smtplib

def email_user():

    cred = credentials.Certificate("LabItEmailer/labit-b36bf-firebase-adminsdk-ntbal-cee91c0aa7.json")
    initialize_app(cred, {"databaseURL": "https://labit-b36bf-default-rtdb.firebaseio.com"})

    db = firestore.client()
    print(1)

    sender = "svc_cs_labit@gcc.edu"
    smtp = smtplib.SMTP("smtp-mail.outlook.com", port = 587)
    print(2)
    smtp.starttls()
    print(3)
    smtp.login(sender, "Fud10200")
    print(4)
    
    docs = (
        db.collection("EmailRequests")
        .stream()
    )

    for doc in docs:
        email_info = doc.to_dict()

        if (email_info.get('type') == "switchInto"):

            recipient = email_info.get('recipient')
            lab_into = email_info.get('labInto')
            lab_out_of = email_info.get('labOutOf')
            subject = email_info.get('subject')
            message = email_info.get('message')
            student_name = email_info.get('studentName')
            requires_action = email_info.get("requiresAction")

            if (requires_action == 1):
                print("switch\n")
                print("to: ", recipient, "\n")
                print("lab into: ", lab_into, "\n")
                print("lab out of: ", lab_out_of, "\n")
                print("subject: ", subject, "\n")
                print("message: ", message, "\n")
                print("student name: ", student_name, "\n")
                print("requires action: ", requires_action, "\n")

                email_body = student_name, " is switching from ", lab_out_of, " into ", lab_into, ".\nSincerely,\n-The LabIt Team"

                try:
                    email = EmailMessage()
                    email["From"] = sender
                    email["To"] = recipient
                    email["Subject"] = subject
                    email.set_content(email_body)

                    smtp.sendmail(sender, recipient, email.as_string())

                    # Update Firestore document
                    # doc.reference.update({"requiresAction": 0})
                    print("Email sent successfully.")
                except Exception as e:
                    print("Error sending email:", e)

        if (email_info.get('type') == "switchOutOf"):

            recipient = email_info.get('recipient')
            lab_into = email_info.get('labInto')
            lab_out_of = email_info.get('labOutOf')
            subject = email_info.get('subject')
            message = email_info.get('message')
            student_name = email_info.get('studentName')
            requires_action = email_info.get("requiresAction")

            if (requires_action == 1):
                print("switch\n")
                print("to: ", recipient, "\n")
                print("lab into: ", lab_into, "\n")
                print("lab out of: ", lab_out_of, "\n")
                print("subject: ", subject, "\n")
                print("message: ", message, "\n")
                print("student name: ", student_name, "\n")
                print("requires action: ", requires_action, "\n")

                email_body = student_name, " is switching from ", lab_out_of, " into ", lab_into, ".\nSincerely,\n-The LabIt Team"

                try:
                    email = EmailMessage()
                    email["From"] = sender
                    email["To"] = recipient
                    email["Subject"] = subject
                    email.set_content(email_body)

                    smtp.sendmail(sender, recipient, email.as_string())

                    # Update Firestore document
                    # doc.reference.update({"requiresAction": 0})
                    print("Email sent successfully.")
                except Exception as e:
                    print("Error sending email:", e)

        if (email_info.get('type') == "request"):
            
            recipient = email_info.get('recipient')
            lab_out_of = email_info.get('labOutOf')
            subject = email_info.get('subject')
            message = email_info.get('message')
            student_name = email_info.get('studentName')
            requires_action = email_info.get("requiresAction")
            replyTo = email_info.get('replyTo')

            if (requires_action == 1):
                print("request\n")
                print("to: ", recipient, "\n")
                print("lab out of: ", lab_out_of, "\n")
                print("subject: ", subject, "\n")
                print("message: ", message, "\n")
                print("student name: ", student_name, "\n")
                print("requires action: ", requires_action, "\n")

                email_body = student_name, " has requested to be absent from  ", lab_out_of, ".\n", message

                try:
                    email = EmailMessage()
                    email["From"] = sender
                    email["To"] = recipient
                    email["Subject"] = subject
                    email['Reply-To'] = replyTo
                    email.set_content(email_body)

                    smtp.sendmail(sender, recipient, email.as_string())

                    # Update Firestore document
                    # doc.reference.update({"requiresAction": 0})
                    print("Email sent successfully.")
                except Exception as e:
                    print("Error sending email:", e)


        if (email_info.get('type') == "contactStudent"):

            recipient = email_info.get('recipient')
            sectionFrom = email_info.get('sectionFrom')
            profFrom = email_info.get('prof')
            subject = email_info.get('subject')
            message = email_info.get('message')
            student_name = email_info.get('studentName')
            requires_action = email_info.get("requiresAction")

            if (requires_action == 1):
                print("contactStudent\n")
                print("to: ", recipient, "\n")
                print("lab section from: ", sectionFrom, "\n")
                print("prof from: ", profFrom, "\n")
                print("subject: ", subject, "\n")
                print("message: ", message, "\n")
                print("requires action: ", requires_action, "\n")

                try:
                    email = EmailMessage()
                    email_body = message
                    email.set_content(email_body)
                    email["From"] = sender
                    email["To"] = recipient
                    email["Subject"] = subject
                    email['Reply-To'] = profFrom

                    smtp.sendmail(sender, recipient, email.as_string())
                    # Update Firestore document
                    # doc.reference.update({"requiresAction": 0})
                    print("Email sent successfully.")
                except Exception as e:
                    print("Error sending email:", e)

    smtp.quit()

email_user()



# import win32com.client
# from firebase_admin import db, credentials, firestore, initialize_app

# cred = credentials.Certificate("labit-b36bf-firebase-adminsdk-ntbal-cee91c0aa7.json")
# initialize_app(cred, {"databaseURL": "https://labit-b36bf-default-rtdb.firebaseio.com"})
# db = firestore.client()

# def email_user():
    
#     docs = (
#         db.collection("EmailRequests")
#         .stream()
#     )

#     for doc in docs:
#         email_info = doc.to_dict()
#         id = email_info.get('id')
#         to = email_info.get('recipient')
#         lab_into = email_info.get('labInto')
#         lab_out_of = email_info.get('labOutOf')
#         subject = email_info.get('subject')
#         message = email_info.get('message')
#         student_name = email_info.get('studentName')
#         requires_action = email_info.get("requiresAction")

#         if (requires_action == 1):
#             print("id: ", id)
#             print("to: ", to)
#             print("lab into: ", lab_into)
#             print("lab out of: ", lab_out_of)
#             print("subject: ", subject)
#             print("message: ", message)
#             print("student name: ", student_name)
#             print("requires action: ", requires_action)

#             try:
#                 ol = win32com.client.Dispatch("outlook.application")
#                 olmailitem = 0x0
#                 newmail = ol.CreateItem(olmailitem)
#                 newmail.Subject = subject
#                 newmail.To = to
#                 newmail.Body = f'{student_name} is requesting to switch from {lab_out_of} into {lab_into}\n\nStudent message: {message}'
#                 newmail.Send()

#                 # Update Firestore document
#                 doc.reference.update({"requiresAction": 0})
#                 print("Email sent successfully.")
#             except Exception as e:
#                 print("Error sending email:", e)

# email_user()



# sender = "svc_cs_labit@gcc.edu"
# print(1)
# smtp = smtplib.SMTP("smtp-mail.outlook.com", port = 587)
# print(2)
# # smtp.ehlo()
# smtp.starttls()
# smtp.login(sender, "Fud10200")

# for i in range(5):
#     email = EmailMessage()
#     message = "Hello world!"
#     recipient = "sinkevitchdp19@gcc.edu"
#     email.set_content(message)
#     email["From"] = sender
#     email["To"] = recipient
#     email["Subject"] = "Test Email"
#     smtp.sendmail(sender, recipient, email.as_string())

# smtp.quit()

# msg = EmailMessage()
# msg['From'] = 'svc_cs_labit@gcc.edu'
# msg['To'] = 'sinkevitchdp19@gcc.edu'
# msg['Subject'] = 'Subject of the email'
# msg.set_content('Body of the email')

# # Connect to the SMTP server
# with smtplib.SMTP('smtp-mail.outlook.com', port=587) as smtp:
#     # Enable encryption
#     smtp.starttls()

#     # Login to the SMTP server (if authentication is required)
#     smtp.login('svc_cs_labit@gcc.edu', 'Fud10200')

#     # Send the email
#     smtp.send_message(msg)