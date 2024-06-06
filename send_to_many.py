import click
import time
import os 
from datetime import datetime, timezone
import silly_teams_client

LITTLE_WHILE = 1

def send_to_many(list_of_users,message, do_message_log=False):
    if (do_message_log):
        log_filename = datetime.now(tz=timezone.utc).strftime("%Y-%m-%dT%H_%M_%S%z") + ".log"
        f = open (log_filename, "w")
        f.write (str(list_of_users) + "\n\n" + message)
    conn = silly_teams_client.SillyTeamsClient()
    for user in list_of_users:
        user1 = conn.get_user_id_from_email(os.getenv("MY_EMAIL"))
        user2 = conn.get_user_id_from_email(user)
        chat_id = conn.create_chat_for_users((user1,user2))
        conn.send_html_message_to_chat(chat_id,message_html=message)
        time.sleep(LITTLE_WHILE)

def assert_os_env_string(env_name):
    if not env_name in os.environ:
        raise ValueError(f"environment variable {env_name} is not set")

@click.command()
@click.option('--audience', help='list of users, comma separated')
@click.option('--message', help='message to send (accepts HTML)')
def main_function(audience,message):        
    audience = audience.replace (" ","")
    list_of_users = audience.split(",")
    send_to_many(list_of_users,message)

if __name__ == '__main__':
    assert_os_env_string("MY_EMAIL")
    assert_os_env_string("AZURE_TENANT_ID")
    assert_os_env_string("AZURE_CLIENT_ID")
    main_function()