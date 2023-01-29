from datetime import datetime
import speech_recognition as sr
import pyttsx3 
import webbrowser
import wikipedia
import wolframalpha
import tkinter as tk
import threading
import win32com.client

# imports for backup
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload 
import os

class Assistant:
    def __init__(self):
        # Speech engine installation
        self.engine = pyttsx3.init()
        self.voices = self.engine.getProperty('voices')
        self.engine.setProperty('voice', self.voices[0].id) # 0=male, 1=female voice
        self.activationword = 'ranger' #single word to activate the software

        #configure browser
        #set the path
        self.edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
        webbrowser.register('edge', None, webbrowser.BackgroundBrowser(self.edge_path))

        # wolfram Alpha client
        self.appID = '4P4A3A-7UXV5GJL7R'
        self.wolframClient = wolframalpha.Client(self.appID)

        self.root = tk.Tk()
        self.label = tk.Label(text="\U0001F916", font=("Arial", 130, "bold"), bg="cyan")
        self.label.pack()

        threading.Thread(target=self.run_assistant).start()
        self.root.mainloop()
    
    def run_assistant(self):
        self.speak('All systems nominal.')
        while True:
            try: 
            # Parse as a list
                query = self.parseCommand().lower().split()
                if query[0] == self.activationword:
                    self.label.config(fg="red")
                    query.pop(0)
                    
                    # list commands
                    if query[0] == 'record':
                        if 'hello' in query:
                            self.speak('Greetings, all.')
                        else:
                            query.pop(0) 
                            speech = ' '.join(query)
                            self.speak(speech)
                    
                    # navigation
                    if query[0] == 'go' and query[1] == 'to':
                        self.speak('Opening...')
                        query = ' '.join(query[2:])
                        webbrowser.get('edge').open_new(query)

                    # wikipedia
                    if query[0] == 'wikipedia':
                        query = ' '.join(query[1:])
                        self.speak('Querying the universal databank.')
                        #print(self.search_wikipedia(query))
                        self.speak(self.search_wikipedia(query))
                        
                    # wolfram Alpha
                    if query [0] == 'compute' or query[0] == 'ranger' or query[0] == 'what':
                        query = ' '.join(query[1:])
                        self.speak('Computing')
                        try:
                            result = self.search_wolframAlpha(query)
                            self.speak(result)
                        except:
                            self.speak('Unable to compute.')

                    # taking note
                    if query[0] == 'log':
                        self.speak("Ready to record your note")
                        newNote = self.parseCommand().lower()
                        now = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
                        with open('note_%s.txt' % now, 'w') as newFile:
                            newFile.write(newNote)
                        self.speak('Note written')
                    
                    if query[0] == 'outlook':
                        self.outlook()

                    if query[0] == 'exit':
                        self.speak('Creating backup, please wait.')
                        try:
                            self.create_backup()
                        except:
                            # self.speak('Backup Created Successfully.')
                            return None
                        self.speak('Backup Created Successfully.')
                        self.speak('Goodbye')
                        self.engine.stop()
                        self.root.destroy()
                        break
                    else:
                        if query is not None:
                            response = self.assistant.request(query)
                            if response is not None:
                                self.speak(response)
                                self.engine.runAndWait()
                        self.label.config(fg="black")

            except:
                print("exception")
                self.label.config(fg="black")
                continue
            

    def speak(self, text, rate = 140):
        self.engine.setProperty('rate', rate)
        self.engine.say (text)
        self.engine.runAndWait()

    def parseCommand(self):
        listener = sr.Recognizer()
        #print('Listening for a command...')

        with sr.Microphone() as source:
            listener.phrase_threshold = 2
            input_speech = listener.listen(source)

        try:
            self.label.config(fg="red")
            #print('Recognizing speech...')
            self.speak("recognizing speech...")
            query = listener.recognize_google(input_speech, language='en_gb')
            #print(f'The input speech was: {query}')
        except Exception as exception:
            #print('I did not quite catch that')
            self.speak('I did not quite catch that')
            #print(exception)
            self.label.config(fg="balck")
            return 'None'
            
        return query
        

    def search_wikipedia(self, query = ''):
        searchResults = wikipedia.search(query)
        if not searchResults:
            #print('No wikipedia result')
            return 'No result received'
        try:
            wikiPage = wikipedia.page(searchResults[0])

        except wikipedia.DisambiguationError as error:
            wikiPage = wikipedia.page(error.options[0])
        #print(wikiPage.title)
        wikiSummarry = str(wikiPage.summary)
        return wikiSummarry

    def listOrDict(self, var):
        if isinstance(var, list):
            return var[0]['plaintext']
        else:
            return var['plaintext'] 

    def search_wolframAlpha(self, query = ''):
        response = self.wolframClient.query(query)

        # @success: Wolfram Alpha was able to resolve the query
        # @numpods: Number of results returned
        # pod: List of results. This can also contain subpods
        if response['@success'] == 'false':
            return 'Could not compute'

        # Query resolved 
        else:
            result = ''
            # Question
            pod0 = response['pod'][0]

            pod1 = response['pod'][1]
            # May contain the answer, has the highest confidence value
            # if it's primary, or has the title of result or definition, then it's the official result
            if (('result') in pod1['@title'].lower()) or (pod1.get('@primary', 'false') == 'true') or ('definition' in pod1['@title'].lower()):
                # Get the result
                result = self.listOrDict(pod1['subpod'])
                # Remove the bracket section
                return result.split('(')[0]
            else:
                question = self.listOrDict(pod0['subpod'])
                # Remove the bracket section
                return question.split('(')[0]
                # search wikipedia instead
                self.speak('Computation failed! Querying universal databank.')
                return search_wikipedia(question)

    def outlook(self):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items
        message = messages.GetLast()

        # print(message.SenderName)
        # print(message.subject)
        engine = pyttsx3.init()
        engine.say("You have an email from {}, Subject of the email is {}, This is what is written in the email: {}".format(message.SenderName, message.subject, message.body))
        engine.runAndWait()

    # def create_backup(self):
    #     backup.main()

    def list_files(self, page_size=10):
        try:
            # Call the Drive v3 API
            results = self.service.files().list(
                pageSize=page_size, fields="nextPageToken, files(id, name)").execute()
            items = results.get('files', [])

            if not items:
                print('No files found.')
                return
            print('Files:')
            for item in items:
                print(u'{0} ({1})'.format(item['name'], item['id']))

        except HttpError as error:
            # TODO(developer) - Handle errors from drive API.
            print(f'An error occurred: {error}')

    def upload_file(self, filename, path):
        folder_id = "1yKInRVn-ZDYNjNN8PCLXo72b2u3r0f3P"
        media = MediaFileUpload(f"{path}{filename}")

        response = self.service.files().list(
                                            q = f"name='{filename}' and parents='{folder_id}'",
                                            spaces='drive',
                                            fields='nextPageToken, files(id, name)',
                                            pageToken=None).execute()
        if len(response['files']) == 0:
            file_metadata = {
                'name': filename,
                'parents': [folder_id]
            }
            file = self.service.files().create(body=file_metadata, media_body=media, fields='id').execute()
            print(f"A new file was created {file.get('id')}")

        else:
            for file in response.get('files',[]):
                #Process change
                
                update_file = self.service.files().update(
                    fileId=file.get('id'),
                    media_body = media,
                ).execute()
                print(f'Updated File')

    def create_backup(self):
        # If modifying these scopes, delete the file token.json.
        SCOPES = ['https://www.googleapis.com/auth/drive']
        """Shows basic usage of the Drive v3 API.
        Prints the names and ids of the first 10 files the user has access to.
        """
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        try:
            self.service = build('drive', 'v3', credentials=creds)
        except HttpError as error:
            # TODO(developer) - Handle errors from drive API.
            print(f'An error occurred: {error}')


        path = "F:/Voice-based AI Assistant/final/"
        myfiles = os.listdir(path)
    
        # my_drive.list_files()
        
        for item in myfiles:
            self.upload_file(item,path)
        
    
bishesh = Assistant()


                








