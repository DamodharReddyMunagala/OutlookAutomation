import win32com.client
import datetime
import os

class Outlook:

    def __init__(self):
        """
            3 - Deleted Mails
            4 - Outbox Mails
            5 - Sent Mails
            6 - Inbox Mails
            9 - Calendar
            10 - Contacts
            11 - Journal
            12 - Notes
            13 - Tasks
            16 - Drafts
        """
        self.__outlookMails = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.__deletedMails = self.__outlookMails.GetDefaultFolder(3).Items
        self.__outboxMails = self.__outlookMails.GetDefaultFolder(4).Items
        self.__sentMails = self.__outlookMails.GetDefaultFolder(5).Items
        self.__inboxMails = self.__outlookMails.GetDefaultFolder(6).Items.restrict("[SentOn] > '07/01/2018 12:00 AM'")
        self.__draftsMails = self.__outlookMails.GetDefaultFolder(16).Items

    
    def __loopingThroughOutlookMails(self, mailsList):
        """
            provided argument is COM objects
        """
        mail = mailsList.GetFirst()
        while mail:
            try:
                print("Sender : " + mail.Sender())
                print("To : " + mail.To)
                print("CC : " + mail.CC)
                print("Date : " + mail.SentOn.strftime("%d-%m-%y"))
                print("Time : " + mail.SentOn.strftime("%H:%M"))
                print("\nSubject : " + mail.Subject)
                print("\nBody : \n" + mail.Body)
                print("\nAttachments : " + mail.attachments)
            except:
                pass
            print("\n\n###################################################################\n\n")
            mail = mailsList.GetNext()


    def __loopingThroughDictionaryOfMails(self, mailsDictionary):
        """
            provided argument is a dictionary of mails
            Dictionary key   : Mail Subject ---- date :: time
            Dictionary value : Mail COM object
        """
        for mail in mailsDictionary.values():
            try:
                print("Sender : " + mail.Sender())
                print("To : " + mail.To)
                print("CC : " + mail.CC)
                print("Date : " + mail.SentOn.strftime("%d-%m-%y"))
                print("Time : " + mail.SentOn.strftime("%H:%M"))
                print("\nSubject : " + mail.Subject)
                print("\nBody : \n" + mail.Body)
                print("\nAttachments : " + mail.attachments)
            except:
                pass
            print("\n\n###################################################################\n\n")


    def readingAllInboxMails(self):
        
        """
            Just reading all the Inbox Folder Mails
        """
        self.__loopingThroughOutlookMails(self.__inboxMails)

    
    def readingAllSentMails(self):

        """
            Just reading all the Sent Folder Mails
        """
        self.__loopingThroughOutlookMails(self.__sentMails)

    
    def readingAllDraftsMails(self):

        """
            Just reading all the Drafts Folder Mails
        """
        mail = self.__draftsMails.GetFirst()
        while mail:
            try:
                print("Subject : " + mail.Subject)
                print("\nBody : \n" + mail.Body)
            except:
                pass
            print("\n\n###################################################################\n\n")
            mail = self.__draftsMails.GetNext()


    def readingAllDeletedMails(self):
        
        """
            Just reading all the Deleted Mails
        """
        self.__loopingThroughOutlookMails(self.__deletedMails)


    def readingInboxMailsByKeywordSearch(self, *args):

        """
            Reading Inbox Mails based on the keywords provided by the user on the mail subject.

            Note:

              Unsolved TestCases:
                
                Challenge 1:
                Problem : The mails in the output are not ordered
                Explanation : As I am using the dictionary data type, order is not guaranteed
        """

        # Creating a dictionary and storing the keyword searched mails without duplicates
        argumentedMails = {}

        inboxMail = self.__inboxMails.GetFirst()

        if len(args) == 0:
            self.readingAllInboxMails()

        while inboxMail:
            for keyword in args:
                if (inboxMail.Subject.find(keyword) != -1):
                    argumentedMails[inboxMail.Subject + " ---- " + inboxMail.CreationTime.strftime("%d-%m-%y :: %H:%M")] = inboxMail
            inboxMail = self.__inboxMails.GetNext()

        self.__loopingThroughDictionaryOfMails(argumentedMails)


    def replyingInboxMails(self, **kwargs):

        """
            Fetching the mails based on :
                1) Keywords we enter - number of keywords we enter doesn't matter
                2) Number of days before we received the mail
                
            Note:

              Solved TestCases:

                Challenge 1:
                Problem : Based on the first condition we get into a problem that we may get duplicate mails.
                Solution : To avoid that I used the dictionary with key as concatenated string of subject, date and time
            
                Challenge 2:
                Problem : Sending specific replies to the specific mails
                Solution : For this I used the named arguments

              Unsolved TestCases:
                
                Challenge 1:
                Problem : Replying to the same mail more than once is not possible through named arguments
                Explanation : As I am using dictionary the duplicate mails are not stored.
                Example : kwargs = {"Used named arguments" : "Testing the application", "Feeling happy" : "Testing the application"}

            Finally:
                To reply to the mails based on the keyword search.
                I looped through the collected mails and sent specific reply to specific mail
        """

        # Creating a dictionary and storing the keyword searched mails without duplicates
        argumentedMails = {}

        inboxMail = self.__inboxMails.GetFirst()

        if len(kwargs) == 0:
            self.readingAllInboxMails()

        while inboxMail:
            for reply, keyword in kwargs.items():
                if (inboxMail.Subject.find(keyword) != -1):
                    argumentedMails[reply + "::" + inboxMail.Subject + " ---- " + inboxMail.CreationTime.strftime("%d-%m-%y :: %H:%M")] = inboxMail
            inboxMail = self.__inboxMails.GetNext()

        for replyMessage, mail in argumentedMails.items():
            #print("Before Splitting : ",replyMessage)
            replyMessage = replyMessage.split("::")[0]
            #print("After Splitting : ",replyMessage)
            mailSubject = mail.Subject
            mail.Body = replyMessage
            reply = mail.Reply()
            reply.Send()
            print("The mail with the subject : '{subject}'has been replied.\
                  \nThe replied message : '{reply}'".format(subject = mailSubject, reply = replyMessage))
            print("\n\n***********************************************************\n\n")
