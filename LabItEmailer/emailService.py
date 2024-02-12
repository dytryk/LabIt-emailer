from emails import email_user
from moveHistory import move_history
import time

count = 0

if __name__ == '__main__':
    
    while True:            
        email_user()
        time.sleep(300)
        count += 1

        if (count % 10 == 0):
            move_history()