


import requests
import json

# Azure AD App details
# tenant_id = 
# client_id = 
# client_secret = 

# Microsoft authentication endpoint for token retrieval
token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

# Define the body to request an access token
body = {
    'grant_type': 'client_credentials',
    'scope': 'https://graph.microsoft.com/.default',
    'client_id': client_id,
    'client_secret': client_secret
}

# Request the access token from Azure AD
token_response = requests.post(token_url, data=body)

# Check if the token was successfully retrieved
if token_response.status_code == 200:
    access_token = token_response.json().get('access_token')
    print("Access token acquired")
else:
    print(f"Failed to retrieve access token: {token_response.status_code}")
    print(token_response.text)
    exit(1)

# Define the email content
mail_body = {
    "message": {
        "subject": "Получилось",
        "body": {
            "contentType": "Text",
            "content": "Салам, все, получилось, рахмет!)"
        },
        "toRecipients": [
            {
                "emailAddress": {
                    "address": "a"
                }
            }
        ]
    },
    "saveToSentItems": "true"
}

# Use the custom URI provided in your PowerShell script
# Replace {user_email} with the email address of the user you want to send as
user_email = "systemreports@norsec.kz"  # Replace with actual user email
send_mail_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/sendMail"

# Send the email via Microsoft Graph API
headers = {
    'Authorization': f'Bearer {access_token}',
    'Content-Type': 'application/json'
}

response = requests.post(send_mail_url, headers=headers, data=json.dumps(mail_body))

# Check if the email was successfully sent
if response.status_code == 202:
    print("Email sent successfully!")
else:
    print(f"Failed to send email: {response.status_code}")
    print(response.text)