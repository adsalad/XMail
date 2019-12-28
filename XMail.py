from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail
import pandas as pd

def list_iter(counter_value,list):
    return list[counter_value]

def main():
    df = pd.read_excel('Book1.xlsx', 'Sheet1')
    emailArray = df['Email'].values
    subjectArray = df["Subject"].values
    bodyArray = df['Body'].values

    counter = emailArray.size-1

    while counter != -1:
        message = Mail(
            from_email= 'adsalads1@gmail.com',
            to_emails=list_iter(counter, emailArray),
            subject=list_iter(counter, subjectArray),
            html_content=list_iter(counter, bodyArray))
        try:
            sg = SendGridAPIClient('APIKEY')
            sg.send(message)
            counter = counter - 1
        except Exception as e:
            print(e.message)

main()
