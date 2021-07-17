import smtplib as s                                                   # library used to send email
import speech_recognition as sr                                       # package used to recognize the speech said by the user
import pyttsx3                                                        # library which converts text to speech offline
import win32com.client                                                # package used to call outlook and read the emails from outlook
from email.message import EmailMessage                                # importing the structure from the package email
import imaplib                                                        # package that allows the end user to view and manipulate the messages

listener = sr.Recognizer()                                            # initializing the recognizer
engine = pyttsx3.init()                                               # initializing pyttsx3

Sender_email = ''                                                     # global variables
Password = ''
choice = 1


def talk(text):                                                       # function which speaks the text passed to it
    voice = engine.getProperty('voices')                              # getting details for voice
    engine.setProperty('voice', voice[choice].id)                     # changes the voice (0 for male and 1 for female)
    engine.setProperty('rate', 150)                                   # makes pyttsx3 module's voice go slower
    engine.say(text)
    engine.runAndWait()                                               # will make the speech audible in the system


def get_info():                                                       # function to get information from user through microphone
    try:                                                              # exception handling
        with sr.Microphone() as source:                               # records the speech from microphone
            print('we are listening')
            listener.pause_threshold = 10
            listener.energy_threshold = 100                           # defines what audio level above 100 should be considered speech
            listener.adjust_for_ambient_noise(source, duration=1)     # listen for 1 second to calibrate the energy threshold for ambient noise levels
            voice = listener.record(source, duration=4)               # records  what is coming from source
            info = listener.recognize_google(voice, language="en")    # converts voice to text
            return info.lower()                                      # returning the text in lower alphabet

    except sr.UnknownValueError:
        talk('sorry i did not listened please say again')
        get_info()


talk('welcome to voiced based automated email service')
talk(' this service provides you to send emails or read out your recent email')
talk('please choose your voice instructor : alexa or david')
voices = get_info()
print(voices)
if "alexa" in voices:
    choice = 1
    talk('hi i am alexa')

elif "david" in voices:
    choice = 0
    talk('hi i am david')

elif None in voices:
    choice = 1



email_list = {                                                         # dictionary which has the receiver's email information
    'work': 'example1@gmail.com',
    'myself': 'example2@gmail.com',
    'brother': 'example3@gmail.com',
    'piyush': 'example4@gmail.com',
    'monica': 'example5@gmail.com',
}




def send_email(receiver, subject, message):                             # function used to send email
    server = s.SMTP("smtp.gmail.com", 587)                              # creating a SMTP server s.Smtp('server name' , port_number)
    server.starttls()                                                   # starts transport layer security
    server.login(Sender_email, Password)                                # sender's info
    email = EmailMessage()                                              # creating the object of EmailMessage
    email['From'] = Sender_email
    email['To'] = receiver
    email['Subject'] = subject
    email.set_content(message)
    server.send_message(email)                                          # send the message through our own SMTP server


def login():                                                            # function to login in to the user account
    talk('speak  your email id')
    global Sender_email, Password                                       # using a global keyword
    Sender_email = get_info()
    print(Sender_email.replace(" ", ""))
    talk('your password')
    Password = get_info()
    engine.say("welcome {}".format(Sender_email.replace("@gmail.com", "")))
    unseen_email()


def unseen_email():                                                     # function to read unseen mail
    try:
        mail = imaplib.IMAP4_SSL('imap.gmail.com', 993)
        mail.login(Sender_email, Password)
        mail.select("inbox", True)                                      # connect to inbox.
        count = len(mail.search(None, 'UnSeen')[1][0].split())

    except Exception as e:
        count = 0
        print(" no emails", e)

    print(" unseen email in box - ", count)
    talk("you have {} new unread mails".format(count))



def receive_email():                                                                                             # function to read email
    unseen_email()
    outlook = win32com.client.Dispatch("Outlook.application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)                                                                          # 6 is for inbox folder
    messages = inbox.Items
    messages.sort("[ReceivedTime]", True)                                                                        # sorts the messages according to time
    message = messages.GetFirst()                                                                                # it will read the latest email
    print("Sender name - ", message.SenderName)
    print("SUBJECT - ", message.subject)
    print("Body - ", message.body)
    talk(
        "you have an email from {} , subject of email is {} , this is the mail {}".format(message.SenderName,
                                                                                          message.subject,
                                                                                          message.body))


talk('first log in to your account')
login()


def get_email_info():

    talk('do you want to sent email   or    read your recent email')
    send_or_receive = get_info()
    if "send" in send_or_receive:
        print(send_or_receive)
        talk('to whom you wanted to send email')
        name = get_info()
        receiver = email_list[name]
        print(receiver)
        talk('what is the subject of your email')
        subject = get_info()
        print(subject)
        talk('tell me the text in your email ')
        message = get_info()
        print(message)
        send_email(receiver, subject, message)
        talk('your message has been sent successfully ')
        talk(' do you want to send more email '
             'say yes or no')
        send_more = get_info()
        if 'yes' in send_more:
            print(send_more)
            get_email_info()
        elif 'no' in send_more:
            print(send_more)
            print('have a good day sir')
            talk('have a good day sir')
        else:
            talk('have a good day sir ')

    elif "read" in send_or_receive:
        print(send_or_receive)
        receive_email()


get_email_info()
