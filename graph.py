import json
from configparser import SectionProxy
from select import select
from wsgiref.util import request_uri
from azure.identity import DeviceCodeCredential, ClientSecretCredential
from msgraph.core import GraphClient
from platformdirs import user_cache_dir

class Graph:
    settings: SectionProxy
    device_code_credential: DeviceCodeCredential
    user_client: GraphClient
    client_credential: ClientSecretCredential
    app_client: GraphClient

    def __init__(self, config: SectionProxy):
        self.settings = config
        client_id = self.settings['clientId']
        tenant_id = self.settings['authTenant']
        graph_scopes = self.settings['graphUserScopes'].split(' ')

        self.device_code_credential = DeviceCodeCredential(client_id, tenant_id = tenant_id)
        self.user_client = GraphClient(credential=self.device_code_credential, scopes=graph_scopes)

    def get_user_token(self):
        graph_scopes = self.settings['graphUserScopes']
        access_token = self.device_code_credential.get_token(graph_scopes)
        return access_token.token

    def get_user(self):
        endpoint = '/me'
        # Only request specific properties
        select = 'displayName,mail,userPrincipalName'
        request_url = f'{endpoint}?$select={select}'

        user_response = self.user_client.get(request_url)
        return user_response.json()

    def get_folders(self):
        endpoint = '/me/mailFolders'
        select = 'displayName,id'
        request_url = f'{endpoint}?$select={select}'

        folder_response = self.user_client.get(request_url)
        return folder_response.json()

    def get_inbox(self, folder_id: str):
        endpoint = f'/me/mailFolders/{folder_id}/messages'
        select = 'from,isRead,receivedDateTime,subject,hasAttachments,id'
        top = 25
        order_by = 'receivedDateTime DESC'
        request_url = f'{endpoint}?$select={select}&$top={top}&$orderBy={order_by}'

        inbox_response = self.user_client.get(request_url)
        return inbox_response.json()

    def get_attachments(self, message_id: str):
        endpoint = f'/me/messages/{message_id}/attachments'
        top = 1
        request_url = f'{endpoint}'

        attachment_response = self.user_client.get(request_url)
        return attachment_response.json()