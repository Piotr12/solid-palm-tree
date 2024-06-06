import requests
import msal
import os
import atexit

class SillyTeamsClient:
    VERY_SECRET_TOKEN_STORE_FILENAME = 'token_cache123.bin'
    TENANT_ID = os.getenv("AZURE_TENANT_ID")
    CLIENT_ID = os.getenv("AZURE_CLIENT_ID")    
    AUTHORITY = 'https://login.microsoftonline.com/' + TENANT_ID
    ENDPOINT = 'https://graph.microsoft.com/v1.0'
    SCOPES = ['User.Read', 'User.ReadBasic.All', 'Chat.ReadWrite', 'Presence.Read.All']

    def __init__(self):
        """ This is a simple client for Microsoft Teams API. 
        It is not meant to be used in production.
        Especially given the token storage mechanism, which is not that secure.
        Upon init it will do its best to authenticate the user and get the token confiming delegated permissions.
        """
        self.cache = msal.SerializableTokenCache()
        if os.path.exists(SillyTeamsClient.VERY_SECRET_TOKEN_STORE_FILENAME):
            self.cache.deserialize(open(SillyTeamsClient.VERY_SECRET_TOKEN_STORE_FILENAME, 'r').read())
        atexit.register(lambda: open(SillyTeamsClient.VERY_SECRET_TOKEN_STORE_FILENAME, 'w').write(self.cache.serialize()) if self.cache.has_state_changed else None)
        app = msal.PublicClientApplication(
            SillyTeamsClient.CLIENT_ID,
            authority = SillyTeamsClient.AUTHORITY,
            token_cache = self.cache)

        self.accounts = app.get_accounts()
        self.login_result = None
        if len(self.accounts) > 0: # silent login if account is already in cache
            self.login_result = app.acquire_token_silent(SillyTeamsClient.SCOPES, account=self.accounts[0])
        if self.login_result is None: # interactive login if account is not in cache or silent login failed
            print ("Requesting token from Microsoft...")
            flow = app.initiate_device_flow(scopes=SillyTeamsClient.SCOPES,  )
            if 'user_code' not in flow:
                raise Exception('Failed to create device flow')
            print(flow['message']) # here I get the url to login presented on the screen
            self.login_result = app.acquire_token_by_device_flow(flow)
        if 'access_token' in self.login_result:
            self.access_token = self.login_result['access_token']
            self.headers={'Authorization': 'Bearer ' + self.access_token, "ConsistencyLevel" : "eventual"}
        else:
            raise Exception('Some miserable failure occured while logging. I am VERY sorry. Bye.')

    def get_user_id_from_email(self,email):
        """be careful with this method, it will return **THE FIRST USER** with the email starting with the string provided."""
        result = requests.get(f'https://graph.microsoft.com/v1.0/users?$filter=startswith(mail,\'{email}\')&$orderby=displayName&$count=true&$top=10',
            headers = self.headers) 
        result.raise_for_status()
        s = result.json()
        user_id = s["value"][0]["id"]
        return user_id

    def create_chat_for_users(self,users_list, title = "automatic chat"):
        """technically it is possible to create a chat with more than 2 users, but I am not sure if it is a good idea."""
        if len (users_list)>2:
            body_head = f"{{\"topic\": \"{title}\", \"chatType\": \"group\", \"members\": [ "
        else:
            body_head = f"{{\"chatType\": \"oneOnOne\", \"members\": [ "

        body_tail = "]}"
        real_body = ""
        for user in users_list:
            real_body += f"""
             {{
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": ["owner"],
                "user@odata.bind": "https://graph.microsoft.com/v1.0/users('{user}')"
                }},"""
        body = body_head + real_body[:-1] + body_tail
        
        result = requests.post(f"https://graph.microsoft.com/v1.0/chats", 
            headers={'Authorization': 'Bearer ' + self.access_token, 'Content-Type': 'application/json'}, 
            data=body)
        
        result.raise_for_status()
        chat_id = result.json()["id"]
        return chat_id
    
    def send_message_to_chat(self,chat_id,message_string):
        post_message_body = f"""
        {{
            "body": {{
            "content": "{message_string}"
            }}
        }}
        """   
        self._send_message(chat_id,post_message_body)
    
    def send_html_message_to_chat(self,chat_id,message_html):
        post_message_body = f"""
        {{
            "body": {{
                "contentType": "html",
                "content": "{message_html}"
            }}
        }}
        """
        self._send_message(chat_id,post_message_body)

    def _send_message(self,chat_id,post_message_body):
        result = requests.post(f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages", 
            headers={'Authorization': 'Bearer ' + self.access_token, 'Content-Type': 'application/json'}, 
            data= post_message_body.encode('utf-8'))    
        result.raise_for_status()
        return result