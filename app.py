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
import boto3
from collections import defaultdict
from googleapiclient.http import MediaFileUpload
import openai

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

def reset_s3():
    # Set AWS details (replace with your own details)
    AWS_REGION_NAME = 'us-east-2'
    AWS_ACCESS_KEY = 'AKIARK3QQWNWXGIGOFOH'
    AWS_SECRET_KEY = 'ClAUaloRIp3ebj9atw07u/o3joULLY41ghDiDc2a'

    # Initialize the S3 client
    s3 = boto3.client('s3',
        region_name=AWS_REGION_NAME,
        aws_access_key_id=AWS_ACCESS_KEY,
        aws_secret_access_key=AWS_SECRET_KEY
    )

    # Delete objects within subdirectories in the bucket 'li-general-tasks'
    subdirs = ['input_videos/', 'output_videos/', 'images/']
    for subdir in subdirs:
        objects = s3.list_objects_v2(Bucket='li-general-tasks', Prefix=subdir)
        for obj in objects.get('Contents', []):
            if obj['Key'] != 'input_videos/outro.mp4':
                s3.delete_object(Bucket='li-general-tasks', Key=obj['Key'])
                
        # Add a placeholder object to represent the "directory"
        s3.put_object(Bucket='li-general-tasks', Key=subdir)


try:   
    if 'begin_auth' not in st.session_state:
        reset_s3()
        st.session_state['creds'] = ""
        st.session_state['begin_auth'] = False
        st.session_state['final_auth'] = False

except Exception as e:
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
    st.write("- Enter the intended output Google sheets link, as well as the input Google drive folder link.")
    st.write("- Enter the desired top left cell where data will start being written to the output Google sheet, as well as the input type within your input folder.")
    st.write("- Click 'Generate Share Links' to being link generation and view them in your destination Google drive sheet.")
    st.subheader("Essay Editing Tool")
    st.write("- Upload the CSV of the gathered data to be transformed into the University Connection Google Doc Template. Columns should be titled PRECISELY: 'Student Email', 'Student Cell', 'High School Name', 'List of colleges that will be receiving this essay or application information are:', 'Supplemental Essays Prompt', 'The writing being edited is for', and 'How many times have you turned in this essay or application information in for review by University Connection?').")
    st.write("- Enter the Google Doc Link of the University Connetion Template with filler text PRECISELY named 'Student-Name-Filler', 'Student-Email-Filler', 'Student-Cell-Filler', 'High-School-Name-Filler', 'University-Filler', 'Application-Material-Filler', 'Application-Type-Filler', and 'Review-Round-Filler'.")
    st.write("- Enter a Google Sheets link so you will have the output of all the COMPLETED University Connection templates.")
    st.subheader("IIP Conjoining")
    st.write("- Upload the CSV of the unsorted IIP team members with columns PRECISELY titled 'Your Full Name (First Name)', 'Your Full Name (Last Name)', 'Your Email Address', 'Your Phone Number', 'Your Skype Name', 'Preffered Pronouns', 'Your Current Grade', 'City', 'State / Province', 'Country', and 'High School Name'.")
    st.write("- Enter your desired output sheet title.")
    st.write("- Click 'Generate Sheet' and receive a new Google sheet that is sorted by IIP team number.")

st.header("Google Authentication")

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
    sheet_url = st.text_input('Google Sheet URL:')
with col2:
    folder_url = st.text_input('Google Drive folder URL:')

col1, col2 = st.columns(2)
with col1:
    starting_cell = st.text_input('Desired top-left cell in output:')
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

def read_google_doc(doc_url):
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

    drive_service = build('drive', 'v3', credentials=creds)
    doc_id = doc_url.split('/')[-2]
    request = drive_service.files().export_media(fileId=doc_id, mimeType='text/plain')
    response = request.execute()
    return response.decode('utf-8')


st.header("Essay Editing Tool")

uploaded_file = st.file_uploader("Upload CSV file", type="csv")
spreadsheet_url = st.text_input("Spreadsheet URL:")
template_doc_link = st.text_input("Template Google Docs URL:")

if st.button("Process Data") and uploaded_file is not None and spreadsheet_url and template_doc_link and st.session_state['final_auth']:
    data = pd.read_csv(uploaded_file, na_values='NaN', keep_default_na=False)
    data = data.fillna("")

    SPREADSHEET_ID = spreadsheet_url.split('/')[-2]
    DOCUMENT_ID = template_doc_link.split('/')[-2]

    if True:
        try:
            with st.spinner("Processing docs (may take a few minutes)..."):
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

                drive_service = build('drive', 'v3', credentials=creds)
                docs_service = build('docs', 'v1', credentials=creds)
                sheets_service = build('sheets', 'v4', credentials=creds)

                for index, row in data.iterrows():
                    doc_title = f"{row['First Name']} {row['Last Name']}"

                    copy_request = {"name": doc_title}
                    try:
                        copied_doc = drive_service.files().copy(fileId=DOCUMENT_ID, body=copy_request).execute()
                        copy_id = copied_doc["id"]
                    except HttpError as error:
                        st.error(f"An error occurred while copying the document: {error}")
                        continue
                    
                    essay_link = row['Please provide Google Doc Link to Essay']
                    if not essay_link:
                        continue
                    essay_content = read_google_doc(essay_link)

                    find_and_replace_requests = [
                        {"replaceAllText": {"containsText": {"text": "Student-Name-Filler", "matchCase": True}, "replaceText": doc_title}},
                        {"replaceAllText": {"containsText": {"text": "Student-Email-Filler", "matchCase": True}, "replaceText": row['Student Email']}},
                        {"replaceAllText": {"containsText": {"text": "Student-Cell-Filler", "matchCase": True}, "replaceText": row['Student Cell']}},
                        {"replaceAllText": {"containsText": {"text": "High-School-Name-Filler", "matchCase": True}, "replaceText": row['High School Name']}},
                        {"replaceAllText": {"containsText": {"text": "University-Filler", "matchCase": True}, "replaceText": row['List of colleges that will be receiving this essay or application information are:']}},
                        {"replaceAllText": {"containsText": {"text": "Application-Material-Filler", "matchCase": True}, "replaceText": row['Supplemental Essays Prompt']}},
                        {"replaceAllText": {"containsText": {"text": "Application-Type-Filler", "matchCase": True}, "replaceText": row['The writing being edited is for']}},
                        {"replaceAllText": {"containsText": {"text": "Review-Round-Filler", "matchCase": True}, "replaceText": row['How many times have you turned in this essay or application information in for review by University Connection?']}},
                        {"replaceAllText": {"containsText": {"text": "Essay", "matchCase": True}, "replaceText": essay_content}},
                    ]

                    try:
                        response = docs_service.documents().batchUpdate(
                            documentId=copy_id,
                            body={"requests": find_and_replace_requests}
                        ).execute()
                    except HttpError as error:
                        st.error(f"An error occurred while replacing text: {error}")

                    # Make the copied document publicly editable
                    drive_service.permissions().create(
                        fileId=copy_id,
                        body={"role": "writer", "type": "anyone"},
                    ).execute()

                    # Get the shareable link
                    file = drive_service.files().get(fileId=copy_id, fields='webViewLink').execute()
                    share_link = file['webViewLink']

                    update_request = {
                        "values": [[doc_title, share_link]]
                    }

                    try:
                        response = sheets_service.spreadsheets().values().append(
                            spreadsheetId=SPREADSHEET_ID,
                            range="A:B",
                            valueInputOption="RAW",
                            body=update_request,
                        ).execute()
                    except HttpError as error:
                        st.error(f"An error occurred while updating the spreadsheet: {error}")

                st.success("Data processed successfully!")
        except Exception as e:
            st.error(f"An error occurred: {e}")

def process_input(input_csv):
    df = pd.read_csv(input_csv)
    teams = defaultdict(list)
    for team_number, group in df.groupby("International Internship Program Team Number"):
        tl_first = group["First Name"].iloc[0]
        tl_last = group["Last Name"].iloc[0]
        team_leader_mask = (
            (group["Your Full Name (First Name)"] == tl_first)
            & (group["Your Full Name (Last Name)"] == tl_last)
        )
        if not any(team_leader_mask):
            continue
        team_leader_data = group[team_leader_mask].iloc[0]
        team = {
            "Team Number": team_number,
            "TL First": team_leader_data["Your Full Name (First Name)"],
            "TL Last": team_leader_data["Your Full Name (Last Name)"],
            "TL Email": team_leader_data["Your Email Address"],
            "TL Phone": team_leader_data["Your Phone Number"],
            "TL Skype": team_leader_data["Your Skype Name"],
            "TL Pronouns": team_leader_data["Preferred Pronouns"],
            "TL Grade": team_leader_data["Your Current Grade"],
            "TL City": team_leader_data["City"],
            "TL State/Province": team_leader_data["State / Province"],
            "TL Country": team_leader_data["Country"],
            "TL High School": team_leader_data["High School Name"],
        }
        for idx, row in enumerate(group.iterrows()):
            index, data = row
            prefix = f"S{idx+1}"
            team[f"{prefix} First"] = data["Your Full Name (First Name)"]
            team[f"{prefix} Last"] = data["Your Full Name (Last Name)"]
            team[f"{prefix} Email"] = data["Your Email Address"]
            team[f"{prefix} Phone"] = data["Your Phone Number"]
            team[f"{prefix} Skype"] = data["Your Skype Name"]
            team[f"{prefix} Pronouns"] = data["Preferred Pronouns"]
            team[f"{prefix} Grade"] = data["Your Current Grade"]
            team[f"{prefix} City"] = data["City"]
            team[f"{prefix} State/Province"] = data["State / Province"]
            team[f"{prefix} Country"] = data["Country"]
            team[f"{prefix} High School"] = data["High School Name"]
        teams[team_number].append(team)
    output_df = pd.DataFrame([item for sublist in teams.values() for item in sublist])
    return output_df

def upload_to_drive(filename, sheet_title):
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

    service_drive = build('drive', 'v3', credentials=creds)
    file_metadata = {
        'name': sheet_title,
        'mimeType': 'application/vnd.google-apps.spreadsheet'
    }
    media = MediaFileUpload(filename, mimetype='text/csv', resumable=True)
    file = service_drive.files().create(body=file_metadata, media_body=media, fields='id').execute()
    spreadsheet_id = file.get('id')
    return f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit'

st.header("IIP Conjoining")
uploaded_file = st.file_uploader("Choose a CSV file", type="csv")
sheet_title = st.text_input('Title for the Google Sheet that will be created:')

if st.button("Generate Sheet") and uploaded_file is not None and sheet_title and st.session_state['final_auth']:
    df = process_input(uploaded_file)

    output_file = 'output.csv'
    df.to_csv(output_file, index=False)
    try:
        url = upload_to_drive(output_file, sheet_title)
        st.write(f'Success! Spreadsheet uploaded to Google Drive: {url}')
    except Exception as e:
        st.write(f'Error occurred during the upload: {str(e)}')

import ast

def process_response(response):
    # Process the response from the GPT-4 API
    content = response['choices'][0]['message']['content']
    matched_pairs = ast.literal_eval(content)
    return matched_pairs


def call_gpt4_api(data):
    # Set up OpenAI API key and model
    openai.api_key = st.secrets['OPENAI_KEY']
    model = "gpt-4"

    # Prepare the input data for GPT-4 API
    bios = data['bio'].tolist()
    names = data['name'].tolist()

    # Create a system message to instruct the model
    messages = [{"role": "system", "content": "You are providing output in a very structured and specific way to help me match people into pairs of roommates based on the similarity of their bios. The only output you should create should be in the form of a 2d array that looks like: [[person1, person2], [person3, person4]] with absolutely no extra text."}]

    # Create a single user message with all the bios
    bio_text = " ".join([f"{names[i]}: {bio}" for i, bio in enumerate(bios)])
    messages.append({"role": "user", "content": f"You are providing output in a very structured and specific way to help me match people into pairs of roommates based on the similarity of their bios. The only output you should create should be in the form of a 2d array that looks like: [[person1, person2], [person3, person4]] with absolutely no extra text. Match the following people into pairs based on their bios: {bio_text}"})

    # Call the GPT-4 API
    try:
        response = openai.ChatCompletion.create(
            model=model,
            messages=messages
        )
        print(response)
        matched_pairs = process_response(response)
        return matched_pairs
    except Exception as e:
        print("Error calling GPT-4 API:", e)
        return []

st.header("Roommate Matcher")

input_sheet = st.file_uploader("Upload input CSV file", type="csv")

if st.button("Match Roommates") and input_sheet and st.session_state['final_auth']:
    with st.spinner("Matching Roommates (may take a few minutes)..."):
        # Google Drive service setup
        CLIENT_SECRET_FILE = 'credentials.json'
        API_NAME = 'drive'
        API_VERSION = 'v3'
        SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']

        with open(CLIENT_SECRET_FILE, 'r') as f:
            client_info = json.load(f)['web']

        creds_dict = st.session_state['creds']
        creds_dict['client_id'] = client_info['client_id']
        creds_dict['client_secret'] = client_info['client_secret']
        creds_dict['refresh_token'] = creds_dict.get('_refresh_token')

        # Create Credentials from creds_dict
        creds = Credentials.from_authorized_user_info(creds_dict)

        service_drive = build('drive', 'v3', credentials=creds)
        
        # Load CSV data
        data = pd.read_csv(input_sheet)
        
        # Send data to GPT-4 model (code to be added)
        # This code should call the GPT-4 API and get the matches as a response
        matches = call_gpt4_api(data)
        
        # Create a Google Sheet with the matched roommates
        service_sheets = build('sheets', 'v4', credentials=creds)
        spreadsheet_body = {
            'properties': {
                'title': 'Sorted Roommates'
            }
        }
        spreadsheet = service_sheets.spreadsheets().create(body=spreadsheet_body).execute()
        spreadsheet_id = spreadsheet['spreadsheetId']
        worksheet_id = spreadsheet['sheets'][0]['properties']['sheetId']

        
        # Format the Google Sheet
        requests = [
            {
                'appendCells': {
                    'sheetId': worksheet_id,
                    'rows': [
                        {
                            'values': [{'userEnteredValue': {'stringValue': 'Roommate 1'}},
                                    {'userEnteredValue': {'stringValue': 'Roommate 2'}}]
                        },
                        *[
                            {
                                'values': [{'userEnteredValue': {'stringValue': match[0]}},
                                        {'userEnteredValue': {'stringValue': match[1]}}]
                            } for match in matches
                        ]
                    ],
                    'fields': 'userEnteredValue'
                }
            }
        ]
        service_sheets.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': requests}
        ).execute()

    st.write(f"Google Sheet created! [View here](https://docs.google.com/spreadsheets/d/{spreadsheet_id})")
