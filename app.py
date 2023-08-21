import streamlit as st
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
import json
import os
import pandas as pd
import re
import requests
import time

st.set_page_config(
    page_title='OfficeEditor',
    page_icon='📃'
) 

hide_streamlit_style = """ <style> #MainMenu {visibility: hidden;} footer {visibility: hidden;} </style> """ 
st.markdown(hide_streamlit_style, unsafe_allow_html=True)
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']

def is_valid_json(file_path):
    try:
        with open(file_path, 'r') as file:
            json.load(file)
        return True
    except json.JSONDecodeError:
        return False

def get_credentials():
    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
    creds = flow.run_local_server(port=0)
    return creds

def get_subfolder_names(folder_id, service, input_type):
    subfolder_names = []
    subfolder_ids = []
    query = ''
    if input_type == 'Files':
        query = f"'{folder_id}' in parents and mimeType!='application/vnd.google-apps.folder'"
    else:
        query = f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.folder'"

    try:
        response = service.files().list(
            q=query,
            spaces='drive',
            fields='files(id, name)',
            pageToken=None
        ).execute()

        subfolder_names = [file['name'] for file in response.get('files', [])]
        subfolder_ids = [file['id'] for file in response.get('files', [])]

    except HttpError as error:
        print(f'An error occurred: {error}')

    return subfolder_names, subfolder_ids

def create_public_link(file_id, service):
    try:
        permission = {
            'type': 'anyone',
            'role': 'writer'
        }
        service.permissions().create(fileId=file_id, body=permission).execute()
        file = service.files().get(fileId=file_id, fields='webViewLink').execute()
        return file['webViewLink']
    except HttpError as error:
        print(f'An error occurred: {error}')

def update_google_sheet(sheet_id, data, creds, starting_cell):
    service = build('sheets', 'v4', credentials=creds)
    body = {
        'values': data
    }
    result = service.spreadsheets().values().append(
        spreadsheetId=sheet_id,
        range=starting_cell,
        valueInputOption='RAW',
        body=body
    ).execute()

def extract_id_from_url(url):
    match = re.search(r'(?<=folders/)[a-zA-Z0-9_-]+', url)
    if match:
        return match.group(0)
    match = re.search(r'(?<=spreadsheets/d/)[a-zA-Z0-9_-]+', url)
    if match:
        return match.group(0)
    return None


try:   
    if 'begin_auth' not in st.session_state:
        st.session_state['creds'] = ""
        st.session_state['begin_auth'] = False
        st.session_state['final_auth'] = False

except:
    print(e)

# Title of the app
st.title("LI Office Editor")
st.caption("By Giacomo Pugliese")

with st.expander("Click to view full directions for this site"):
    st.subheader("Google Authentication")
    st.write("- Click 'Authenticate Google Account', and then on the generated link.")
    st.write("- Follow the steps of Google login until you get to the final page.")
    st.write("- Click on 'Finalize Authentication' to proceed to rest of website.")
    st.subheader("Google Drive Sharing Links Tool")
    st.write("- Enter the intended output Google sheets link, as well as the input Google drive folder link. ")
    st.write("- Enter the desired top left cell where data will start being written to the output Google sheet, as well as the input type within your input folder.")
    st.write("- Click 'Generate Share Links' to being link generation and view them in your destination Google drive sheet.")

st.header("Google authentication")

try:
    if st.button("Authenticate Google Account"):
        st.session_state['begin_auth'] = True
        # Request OAuth URL from the FastAPI backend
        response = requests.get(f"{'https://leadership-initiatives-0c372bea22f2.herokuapp.com'}/auth?user_id={'intros'}")
        if response.status_code == 200:
            # Get the authorization URL from the response
            auth_url = response.json().get('authorization_url')
            st.markdown(f"""
                <a href="{auth_url}" target="_blank" style="color: #8cdaf2;">
                    Click to continue to authentication page (before finalizing)


                </a>
                """, unsafe_allow_html=True)
            st.text("\n\n\n")
            # Redirect user to the OAuth URL
            # nav_to(auth_url)

    if st.session_state['begin_auth']:    
        if st.button("Finalize Google Authentication"):
            with st.spinner("Finalizing authentication..."):
                for i in range(6):
                    # Request token from the FastAPI backend
                    response = requests.get(f"{'https://leadership-initiatives-0c372bea22f2.herokuapp.com'}/token/{'sheets'}")
                    if response.status_code == 200:
                        st.session_state['creds'] = response.json().get('creds')
                        print(st.session_state['creds'])
                        st.success("Google account successfully authenticated!")
                        st.session_state['final_auth'] = True
                        break
                    time.sleep(1)
            if not st.session_state['final_auth']:
                st.error('Experiencing network issues, please refresh page and try again.')
                st.session_state['begin_auth'] = False

except Exception as e:
    print(e)

st.header('Google Drive Sharing Links Tool')


col1, col2 = st.columns(2)
with col1:
    sheet_url = st.text_input('Enter your Google Sheet URL:')
with col2:
    folder_url = st.text_input('Enter your Google Drive folder URL:')

col1, col2 = st.columns(2)
with col1:
    starting_cell = st.text_input('Enter desired top-left cell in output:')
with col2:
    input_type = st.selectbox("Input type", ['Subfolders', 'Files'])

if starting_cell == '':
    starting_cell = 'A1'

if st.button('Generate Share Links') and st.session_state['final_auth'] and sheet_url and folder_url:
    folder_id = extract_id_from_url(folder_url)
    sheet_id = extract_id_from_url(sheet_url)
    
    # Google Drive service setup
    CLIENT_SECRET_FILE = 'credentials.json'
    API_NAME = 'drive'
    API_VERSION = 'v3'
    SCOPES = ['https://www.googleapis.com/auth/drive']

    with open(CLIENT_SECRET_FILE, 'r') as f:
        client_info = json.load(f)['web']

    creds_dict = st.session_state['creds']
    creds_dict['client_id'] = client_info['client_id']
    creds_dict['client_secret'] = client_info['client_secret']
    creds_dict['refresh_token'] = creds_dict.get('_refresh_token')

    # Create Credentials from creds_dict
    creds = Credentials.from_authorized_user_info(creds_dict)

    # Build the Google Drive service
    drive_service = build('drive', 'v3', credentials=creds)

    subfolder_names, subfolder_ids = get_subfolder_names(folder_id, drive_service, input_type)

    data = []
    for name, folder_id in zip(subfolder_names, subfolder_ids):
        link = create_public_link(folder_id, drive_service)
        data.append([name, link])

    update_google_sheet(sheet_id, data, creds, starting_cell)


