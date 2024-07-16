from O365 import Account
from O365.utils.token import FileSystemTokenBackend

def send_email():
    try:
        # add attachment
        tk = FileSystemTokenBackend(token_path=".", token_filename='o365_token.txt')
        credentials = ('client_id')
        account = Account(credentials=credentials, auth_flow_type='public', token_backend=tk)
        m = account.new_message()
        m.to.add(['someone@somewhere.com', 'another@somewhere.com']) # read emails from confing.ini file
        m.subject = 'Email Subject' # config.ini
        m.body = "Insert email text here" # config.ini
        m.send()
        # logging
    except Exception as e:
        pass
        # log the error into log file
    