from firebase_admin import db, credentials, firestore, initialize_app

cred = credentials.Certificate("LabItEmailer/labit-b36bf-firebase-adminsdk-ntbal-cee91c0aa7.json")
initialize_app(cred, {"databaseURL": "https://labit-b36bf-default-rtdb.firebaseio.com"})
db = firestore.client()

def move_history():
    
    docs = (
        db.collection("EmailRequests")
        .stream()
    )

    i = 0

    for doc in docs:
        email_info = doc.to_dict()

        if (email_info.get('type') == "contactStudent"):
            print("contactStudent")

            if (email_info.get('requiresAction') == 1):
                print("requiresAction")
                recipient = email_info.get('recipient')
                sectionFrom = email_info.get('sectionFrom')
                profFrom = email_info.get('prof')
                subject = email_info.get('subject')
                message = email_info.get('message')
                student_name = email_info.get('studentName')
                requires_action = email_info.get("requiresAction")

                data = {
                    "recipient": recipient, 
                    "sectionFrom": sectionFrom, 
                    "prof": profFrom, 
                    "subject": subject, 
                    "message": message, 
                    "studentName": student_name, 
                    "requiresAction": requires_action
                }

                try:
                    db.collection("EmailHistory").document(doc.id).set(data)
                    db.collection("EmailRequests").document(doc.id).delete()
                    print("data added successfully")
                except Exception as e:
                    print("error writing to db:", e)


move_history()