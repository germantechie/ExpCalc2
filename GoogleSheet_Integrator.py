#!/usr/bin/python

# -------------------------------------------------------------------------------------------
# Developed By - Tom Thomas on 15 December 2014                                             |
# Current Version - 1.0                                                                     |
# Revision History                                                                          |
# 1.0 - 15 December 2014                                                                    |
#       Developed this script to read Google sheet that holds the input data for EXPCalc    |
# -------------------------------------------------------------------------------------------

# Client ID  1066503100456-a8ofhai5mk0c2s14oebfv8mha4emcm60.apps.googleusercontent.com
# Client secret THwk0AysqTYOmCeeF8NL8-wB
# Redirect URIs  urn:ietf:wg:oauth:2.0:oob   http://localhost

import httplib2        # This is used by GDrive API Client
import apiclient       # This is used by GDrive API Client
import webbrowser      # This is to launch browser and go to the link for user permission
import subprocess, os, sys
import logging 

from apiclient.discovery import build
from oauth2client.file import Storage
from oauth2client.client import OAuth2WebServerFlow

# Config_ExpenCalc.py is a configuration file which should be in the current directory as this script
from Config_ExpCalc import *

logging.basicConfig(filename='ExpnCalc.log', filemode='w', level=logging.DEBUG)
logging.info('Starting to connect Google Drive...')
# This is the file from Google Drive that I need to download. Used for drive_file variable down.
file_id = GOOGLE_SHEET_ID

# Copy your credentials from the Google Developer console - https://console.developers.google.com
CLIENT_ID = CLIENT_ID
CLIENT_SECRET = CLIENT_SECRET

# Check https://developers.google.com/drive/scopes for all available scopes
OAUTH_SCOPE = 'https://www.googleapis.com/auth/drive'

# Redirect URI for installed apps
REDIRECT_URI = 'urn:ietf:wg:oauth:2.0:oob' #'http://localhost'

# Create a credential storage object.  You pick the filename.
storage = Storage(sys.path[0] + '/ACCESS_TOKEN')

# Attempt to load existing credentials.  Null is returned if it fails.
credentials = storage.get()

if not credentials:
	# Run through the OAuth flow and retrieve credentials
	flow = OAuth2WebServerFlow(CLIENT_ID, CLIENT_SECRET, OAUTH_SCOPE, redirect_uri=REDIRECT_URI)
	authorize_url = flow.step1_get_authorize_url()

	#print 'Go to the following link in your browser: ' + authorize_url
	logging.info('Please enter user permission...')
	webbrowser.open_new_tab(authorize_url)
	code = raw_input('Enter verification code: ').strip()
	credentials = flow.step2_exchange(code)

	# This Credentials has refresh and access tokens. This should be stored. Find out how.
	storage.put(credentials)


# Create an httplib2.Http object and authorize it with our credentials
logging.info('Connected to Google Drive via existing token...')
http = httplib2.Http()
http = credentials.authorize(http)

drive_service = build('drive', 'v2', http=http)

drive_file = drive_service.files().get(fileId=file_id).execute()

downloadUrl = drive_file.get('exportLinks')['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']
#print 'Download URL is %s' % downloadUrl

resp, content = drive_service._http.request(downloadUrl)
if resp.status == 200:
    #print 'Status: %s' % resp
    logging.info('Writing data from Google to local file...')
    Gfile = open(sys.path[0] + DATA_INPUT_WORKBOOK_NAME, "wb")
    Gfile.write(content)
    Gfile.close()
    print 'Downloaded the file from Google. Ready to crunch statistics...'
    logging.info('Downloaded the file from Google. Ready to crunch statistics...')
    
    os.system('python ' + sys.path[0] + '/ExpenseCalc_nix.py')
    logging.info('Calling data crunching program, ExpenseCalc_nix.py...')
else:
    print 'An error occurred: %s' % resp
    logging.error('An error occurred during GDrive connection: %s' % resp)

#os.remove('/home/tom/Desktop/Personal-budget.xls')

# def download_file(service, drive_file):
#   """Download a file's content.
#
#   Args:
#     service: Drive API service instance.
#     drive_file: Drive File instance.
#
#   Returns:
#     File's content if successful, None otherwise.
#   """
#   download_url = file['exportLinks']['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']
#
#   if download_url:
#     resp, content = service._http.request(download_url)
#   else:
#       if resp.status == 200:
#           print 'Status: %s' % resp
#           return content
#       else:
#           print 'An error occurred: %s' % resp
#           return None
#       # The file doesn't have any content stored on Drive.
#       return None
