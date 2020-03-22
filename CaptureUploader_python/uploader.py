from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import datetime
import sys


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']

def searchTarget(targetName, isFolder, parentID):
    page_token = None
    drive_service = getService()

    Query = "name = '" + targetName;
    if (isFolder):
        Query += "' and mimeType = 'application/vnd.google-apps.folder'"
    else:
        Query += "' and mimeType != 'application/vnd.google-apps.folder'"

    if (parentID != None):
        Query += " and '" + parentID + "' in parents"
    while True:
        response = drive_service.files().list(q=Query,
                                              spaces='drive',
                                              fields='nextPageToken, files(id, name)',
                                              pageToken=page_token).execute()
        for file in response.get('files', []):
            # Process change
            return file
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break

def createGoogleDriveFolder(name, parentID):
    drive_service = getService()

    file_metadata = {
        'name': name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parentID]
    }
    file = drive_service.files().create(body=file_metadata, fields='id').execute()
    print('created: Folder ID: %s' % file.get('id'))

def createGoogleDriveFile(name, parentID):
	drive_service = getService()

    file_metadata = {
        'name': name,
        'parents': [parentID]
    }

    media = MediaFileUpload('files/photo.jpg',
                            mimetype='image/jpeg')
    file = drive_service.files().create(body=file_metadata,
                                        media_body=media,
                                        fields='id').execute()
    print 'File ID: %s' % file.get('id')

def getService():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('drive', 'v3', credentials=creds)
    return service

def getMetaList():
    # Call the Drive v3 API
    results = service.files().list(
        pageSize=10, fields="nextPageToken, files(id, name)").execute()
    items = results.get('files', [])

    if not items:
        print('No files found.')
    else:
        print('Files:')
        for item in items:
            print(u'{0} ({1})'.format(item['name'], item['id']))


def main():
    #service= getService()
    #print(service)

    folderKoreanFilterList = searchTarget("Korean filter list screenshots", True, None);
    if folderKoreanFilterList is None:
        print("No target to upload. (No folderKoreanFilterList)")
        return None
    else:
        print('Found fodler: %s (%s)' % (folderKoreanFilterList.get('name'), folderKoreanFilterList.get('id')))
    
    #createGoogleDriveFolder('2022', folderKoreanFilterList.get('id'))

    now = datetime.datetime.now()
    strYear= str(now.year)
    folderYear = searchTarget(strYear, True, folderKoreanFilterList.get('id'));    
    if folderYear is None:
        print("No target to upload. --> try to create a new year folder.")
        createGoogleDriveFolder(strYear, folderKoreanFilterList.get('id'))

    strMD = now.strftime("%b%d")
    folderUploadHere= searchTarget(strMD, True, folderYear.get('id'));    
    if folderUploadHere is None:
        print("No target to upload. --> try to create a new month+day folder.")
        createGoogleDriveFolder(strMD, folderYear.get('id'))

    if len(sys.argv) == 0:
    	print("No file to upload.")
    	return None
    
    print(sys.argv[0])
    createGoogleDriveFile(sys.argv[0], folderUploadHere)




if __name__ == '__main__':
    main()