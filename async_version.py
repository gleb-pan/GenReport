from O365 import Account

scopes =  ["IMAP.AccessAsUser.All", "POP.AccessAsUser.All", "SMTP.Send", "Mail.Send", "offline_access"]

account = Account(credentials=('be3b533b-6900-4e84-8338-a5934799565d', 'nwl8Q~xa56X1r2TUcHHm5VG8mOXACbi66dX_Ras7'))
result = account.authenticate(scopes=scopes)  # request a token for this scopes