from emails import email_user
import time

if __name__ == '__main__':
    
    while True:
        email_user()
        time.sleep(300)