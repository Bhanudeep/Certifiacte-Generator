from googleapiclient.discovery import build
import os
directory=os.getcwd()
from google.oauth2 import service_account
from PIL import Image, ImageDraw, ImageFont
print(directory)
import pandas as pd
import tkinter as tk
root=tk.Tk()
root.title('Certificates Generator')
# setting the windows size
root.geometry("800x300")
# declaring string variable
# for storing name and password
name_var=tk.StringVar()
sheet_var=tk.StringVar()
date_var=tk.StringVar()
def genr():
        SERVICE_ACCOUNT_FILE = 'keys.json'
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        creds=None
        creds = service_account.Credentials.from_service_account_file(
                SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        # If modifying these scopes, delete the file token.json.


        # The ID and range of a sample spreadsheet.
        SAMPLE_SPREADSHEET_ID = name_var.get() #upto39 from84
        name_var.set("")
        SAMPLE_SPREADSHEET_ID = SAMPLE_SPREADSHEET_ID[39:83]
        print(SAMPLE_SPREADSHEET_ID)
        SAMPLE_SPREADSHEET_ID.replace('/export?format=csv&gid=0',"")


        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=sheet_var.get()).execute()
        values = result.get('values', [])
        # print(values)

        # data = pull_sheet_data(SCOPES,SPREADSHEET_ID,DATA_TO_PULL)
        df = pd.DataFrame(values[1:], columns=values[0])

        # print(df)
        sub =date_var.get()
        
        # creating and passing series to new column
        df["Today"]= df["Timestamp"].str.find(sub)
        
        # display
        # print(df)

        df=df.loc[df['Today'] == 0]
        df=df.loc[ : , ['Timestamp', 'Name','Organisation','Phone','Email'] ]
        print(df)

        #cretificate generation
        font = ImageFont.truetype('arial.ttf',60)
        for index,j in df.iterrows():
                img = Image.open('certificate.jpg')
                draw = ImageDraw.Draw(img)
                draw.text(xy=(725,760),text='{}'.format(j['Name']),fill=(0,0,0),font=font)
                img.save('pictures/{}.jpg'.format(j['Name']))
name_label = tk.Label(root, text = 'Enter URL of sheets', font=('calibre',10, 'bold'))
sheet_label = tk.Label(root, text = 'Enter Sheet Name', font=('calibre',10, 'bold'))
date_label=tk.Label(root, text = 'Enter Date to Generate certificates from', font=('calibre',10, 'bold'))
note_label=tk.Label(root, text = 'NOTE: Make sure you enable share options in google sheets and \n'+'Add this  "certify@certify-319513.iam.gserviceaccount.com"\n and enable Editor access.', font=('calibre',10, 'bold'))
note_label.config(bg="gray")
# name using widget Entry
name_entry = tk.Entry(root,textvariable = name_var, font=('calibre',10,'normal'))
sheet_entry= tk.Entry(root,textvariable = sheet_var, font=('calibre',10,'normal'))
date_entry=tk.Entry(root,textvariable = date_var, font=('calibre',10,'normal'))
  
# creating a button using the widget
# Button that will call the submit function
sub_btn=tk.Button(root,text = 'Generate Certificates', command = genr)
  
# placing the label and entry in
# the required position using grid
# method
name_label.grid(row=0,column=0)
name_entry.grid(row=0,column=1)
sheet_label.grid(row=1,column=0)
sheet_entry.grid(row=1,column=1)
date_label.grid(row=2,column=0)
date_entry.grid(row=2,column=1)
sub_btn.grid(row=3,column=1)
note_label.grid(row=8,column=0)
# performing an infinite loop
# for the window to display
root.mainloop()