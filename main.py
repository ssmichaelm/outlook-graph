import configparser
from venv import create
from graph import Graph
import os.path

def main():
    print('Python Graph Tutorial\n')

    # Load settings
    config = configparser.ConfigParser()
    config.read(['config.cfg', 'config.dev.cfg'])
    azure_settings = config['azure']

    graph: Graph = Graph(azure_settings)

    folders = list_folders(graph)

    # Check to see if the new folder that messages will be moved to exists (i.e. was created)
    new_Folder_Name = 'completed_Cybersecurity_Certs'
    
    try:
        completed_folder_id = get_folder_id(new_Folder_Name, folders)
    except:
        print("Creating folder")
        completed_folder_id = create_folder(graph)
        print(completed_folder_id['id'])

    messages = list_inbox(graph, get_folder_id('cybersecurity',folders))
    download_attachments(graph, messages)

def list_folders(graph: Graph) -> list:
    folders = graph.get_folders().get('value')

    return folders

def get_folder_id(folder_name: str, folders: list) -> str:
    for x in folders:
        if x['displayName'] == folder_name:
            folder_id = x['id']

    return folder_id

def create_folder(graph: Graph) -> str:
    folder = { 'displayName': 'completed_Cybersecurity_Certs' }
    folder_creation = graph.create_folder(folder)

    return folder_creation.json()

def list_inbox(graph: Graph, folder_id: str) -> list:
    message_page = graph.get_inbox(folder_id).get('value')
    messages_list = []
    for message in message_page:
        x = {}
        x['id']      = message['id']
        x['subject'] = message['subject']
        x['email']   = message['sender']['emailAddress']['address']
        x['sender']  = message['sender']['emailAddress']['name']
        x['user']    = parse_name(x['sender'])
        messages_list.append(x)

    return messages_list

def parse_name(sender: str) -> dict:
    user = {}

    # Last, First format
    if ',' in sender:
        sender = sender.split(', ')
        user['First name'] = sender[1]
        user['Last name'] = sender[0]

        temp = ''
        i = 0
        if '[' in user['First name']:
            while user['First name'][i] != ' ':
                temp += user['First name'][i]
                i += 1

        user['First name'] = temp

    # First Last format
    else:
        sender = sender.split(' ')
        user['First name'] = sender[0]
        user['Last name'] = sender[1]

    return user

def download_attachments(graph: Graph, messages: list):
    for message in messages:
        attachment = graph.get_attachments(message['id']).get('value')
        file_suffix = attachment[0]['contentType']
        file_suffix = file_suffix.split('/')[1]
        attachment_content = graph.download_attachments(message['id'], attachment[0]['id'])

        with open("{0}_{1}_Cybersecurity.{2}".format(message['user']['Last name'], 
                                                    message['user']['First name'], file_suffix), 
        'wb') as _f:
            _f.write(attachment_content.content)

def mark_read(graph: Graph, message_id: str):
    return 1

def move_message(graph: Graph, message_id: str):
    return 1

# Run main
main()
