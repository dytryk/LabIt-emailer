import win32com.client
from firebase_admin import db, credentials, firestore, initialize_app

cred = credentials.Certificate("labit-b36bf-firebase-adminsdk-ntbal-cee91c0aa7.json")
initialize_app(cred, {"databaseURL": "https://labit-b36bf-default-rtdb.firebaseio.com"})
db = firestore.client()

def email_user():
    
    docs = (
        db.collection("EmailRequests")
        .stream()
    )

    for doc in docs:
        email_info = doc.to_dict()
        id = email_info.get('id')
        to = email_info.get('recipient')
        lab_into = email_info.get('labInto')
        lab_out_of = email_info.get('labOutOf')
        subject = email_info.get('subject')
        message = email_info.get('message')
        student_name = email_info.get('studentName')
        requires_action = email_info.get("requiresAction")

        if (requires_action == 1):
            print("id: ", id)
            print("to: ", to)
            print("lab into: ", lab_into)
            print("lab out of: ", lab_out_of)
            print("subject: ", subject)
            print("message: ", message)
            print("student name: ", student_name)
            print("requires action: ", requires_action)

            try:
                ol = win32com.client.Dispatch("outlook.application")
                olmailitem = 0x0
                newmail = ol.CreateItem(olmailitem)
                newmail.Subject = subject
                newmail.To = to
                newmail.Body = f'{student_name} is requesting to switch from {lab_out_of} into {lab_into}\n\nStudent message: {message}'
                newmail.Send()

                # Update Firestore document
                doc.reference.update({"requiresAction": 0})
                print("Email sent successfully.")
            except Exception as e:
                print("Error sending email:", e)

email_user()