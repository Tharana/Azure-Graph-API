import os
import json
from datetime import datetime, timedelta
from configparser import SectionProxy
from azure.identity import DeviceCodeCredential
from msgraph import GraphServiceClient
from msgraph.generated.users.item.user_item_request_builder import UserItemRequestBuilder
from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (
    MessagesRequestBuilder)
from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
    SendMailPostRequestBody)
from msgraph.generated.models.message import Message
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.recipient import Recipient
from msgraph.generated.models.email_address import EmailAddress

class Graph:
    settings: SectionProxy
    device_code_credential: DeviceCodeCredential
    user_client: GraphServiceClient
    token_cache_file = 'token_cache.json'

    def __init__(self, config: SectionProxy):
        self.settings = config
        client_id = self.settings['clientId']
        tenant_id = self.settings['tenantId']
        graph_scopes = self.settings['graphUserScopes'].split(' ')

        self.device_code_credential = DeviceCodeCredential(client_id, tenant_id=tenant_id)
        self.user_client = GraphServiceClient(self.device_code_credential, graph_scopes)
        self.token = self._load_token()

    def _load_token(self):
        if os.path.exists(self.token_cache_file):
            with open(self.token_cache_file, 'r') as f:
                return json.load(f)
        return None

    def _save_token(self, token):
        with open(self.token_cache_file, 'w') as f:
            json.dump(token, f)

    async def _get_access_token(self):
        if self.token:
            # Check if token is expired
            if 'expires_on' in self.token and datetime.fromtimestamp(self.token['expires_on']) > datetime.now():
                return self.token['access_token']

        # Acquire new token using refresh token if available
        if self.token and 'refresh_token' in self.token:
            new_token = self.device_code_credential.get_token(
                scopes=self.settings['graphUserScopes'].split(' '),
                refresh_token=self.token['refresh_token']
            )
        else:
            new_token = self.device_code_credential.get_token(self.settings['graphUserScopes'].split(' '))

        self.token = {
            'access_token': new_token.token,
            'refresh_token': new_token.refresh_token,
            'expires_on': (datetime.now() + timedelta(seconds=new_token.expires_in)).timestamp()
        }
        self._save_token(self.token)
        return self.token['access_token']

    async def get_user_token(self):
        return await self._get_access_token()

    async def get_user(self):
        query_params = UserItemRequestBuilder.UserItemRequestBuilderGetQueryParameters(
            select=['displayName', 'mail', 'userPrincipalName']
        )

        request_config = UserItemRequestBuilder.UserItemRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        user = await self.user_client.me.get(request_configuration=request_config)
        return user

    async def get_inbox(self):
        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            select=['from', 'isRead', 'receivedDateTime', 'subject'],
            top=25,
            orderby=['receivedDateTime DESC']
        )
        request_config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        messages = await self.user_client.me.mail_folders.by_mail_folder_id('inbox').messages.get(
                request_configuration=request_config)
        return messages

    async def send_mail(self, subject: str, body: str, recipient: str):
        message = Message()
        message.subject = subject

        message.body = ItemBody()
        message.body.content_type = BodyType.Text
        message.body.content = body

        to_recipient = Recipient()
        to_recipient.email_address = EmailAddress()
        to_recipient.email_address.address = recipient
        message.to_recipients = []
        message.to_recipients.append(to_recipient)

        request_body = SendMailPostRequestBody()
        request_body.message = message

        await self.user_client.me.send_mail.post(body=request_body)

    async def make_graph_call(self):
        # INSERT YOUR CODE HERE
        return
