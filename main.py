import configparser
from graph import Graph

def main():
    print('Python Graph Tutorial\n')

    # Load settings
    config = configparser.ConfigParser()
    config.read(['config.cfg', 'config.dev.cfg'])
    azure_settings = config['azure']

    graph: Graph = Graph(azure_settings)

    folder_id = list_folders(graph)
    list_inbox(graph, folder_id)
    
def greet_user(graph: Graph):
    user = graph.get_user()
    print('Hello,', user['displayName'])
    # For Work/school accounts, email is in mail property
    # Personal accounts, email is in userPrincipalName
    print('Email:', user['mail'] or user['userPrincipalName'], '\n')

def display_access_token(graph: Graph):
    token = graph.get_user_token()
    print('User token:', token, '\n')

def list_folders(graph: Graph):
    folders = graph.get_folders().get('value')

    for x in folders:
        if x['displayName'] == 'cybersecurity':
            folder_id = x['id']

    return folder_id

def list_inbox(graph: Graph, folder_id: str):
    message_page = graph.get_inbox(folder_id)

    # Output each message's details
    for message in message_page['value']:
        print('Message:', message['subject'])
        print('  From:', message['from']['emailAddress']['name'])
        print('  Status:', 'Read' if message['isRead'] else 'Unread')
        print('  Received:', message['receivedDateTime'])

    # If @odata.nextLink is present
    more_available = '@odata.nextLink' in message_page
    print('\nMore messages available?', more_available, '\n')

# # Run main
main()