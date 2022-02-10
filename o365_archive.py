from asyncio.windows_events import NULL
from exchangelib import OAuth2Credentials, OAuth2AuthorizationCodeCredentials, Identity, OAUTH2, Account, Configuration, IMPERSONATION
from types import SimpleNamespace
import os, json

if not os.path.exists('config.json'):
    raise Exception(f"config.json not found in {os.getcwd()}!")

with open('config.json','r') as f: appCfg = json.load(f, object_hook=lambda d:SimpleNamespace(**d))

archiveBasePath = os.path.join(os.getcwd(),'archive')
if not os.path.exists(archiveBasePath):
    os.makedirs(archiveBasePath)

creds = OAuth2Credentials(
    client_id=appCfg.config.App.Id,
    client_secret=appCfg.config.App.Secret,
    tenant_id=appCfg.config.Ms365TenantId,
    identity=Identity(primary_smtp_address=appCfg.config.MbxUpn)
    )

conf = Configuration(
    credentials=creds,
    auth_type=OAUTH2,
    service_endpoint='https://outlook.office365.com/EWS/Exchange.asmx'
    )

account = Account(
    primary_smtp_address=appCfg.config.MbxUpn,
    credentials=creds,
    autodiscover=False,
    config=conf,
    access_type=IMPERSONATION
    )

for item in account.inbox.all()[:10]:

    cleanId = (item.id).replace("<","")
    cleanId = cleanId.replace(">","")
    cleanId = cleanId.replace(":","")
    cleanId = cleanId.replace("\"","")
    cleanId = cleanId.replace("/","")
    cleanId = cleanId.replace("\\","")
    cleanId = cleanId.replace("|","")
    cleanId = cleanId.replace("?","")
    cleanId = cleanId.replace("*","")

    msgPath = f"{archiveBasePath}\\{item.subject}_{cleanId}.eml"
    with open(msgPath,"wb", buffering=0) as f: f.write(item.mime_content)
    
    for att in item.attachments:
        attPath=os.path.join(archiveBasePath, att.name)
        with open(attPath, "wb") as f: f.write(att.content)


