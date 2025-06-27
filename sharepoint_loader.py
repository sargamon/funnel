'''
sharepoint_loader.py

Provides load_the_datafile to download and pivot an Excel file from SharePoint using delegated (user login) permissions via MSAL Device Code Flow.

Requirements:
    pip install pandas msal office365-rest-python-client
'''
import os
import io
import pandas as pd
import msal
import requests

def acquire_token_device_flow(client_id: str, tenant_id: str, scope: list) -> str:
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.PublicClientApplication(client_id, authority=authority)

    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(scope, account=accounts[0])
    if not result:
        flow = app.initiate_device_flow(scopes=scope)
        if "user_code" not in flow:
            raise ValueError(f"Failed to start device flow: {flow}")
        print(flow["message"])  # Replace with st.info(...) in Streamlit
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise Exception(f"Could not acquire token: {result.get('error_description')}")
    return result["access_token"]

def download_excel_from_sharepoint(site_id: str, item_id: str, access_token: str) -> pd.DataFrame:
    headers = {
        "Authorization": f"Bearer {access_token}"
    }
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/content"
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return pd.read_excel(io.BytesIO(response.content), engine="openpyxl")

# Example usage
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
TENANT_ID = os.getenv("AZURE_TENANT_ID")
SITE_ID = os.getenv("SP_SITE_ID")       # e.g., 'contoso.sharepoint.com,site-id,web-id'
ITEM_ID = os.getenv("SP_ITEM_ID")       # e.g., 'item-id'

SCOPE = ["https://graph.microsoft.com/.default"]

if not all([CLIENT_ID, TENANT_ID, SITE_ID, ITEM_ID]):
    raise EnvironmentError("Missing one or more required environment variables.")

token = acquire_token_device_flow(CLIENT_ID, TENANT_ID, SCOPE)
df = download_excel_from_sharepoint(SITE_ID, ITEM_ID, token)

print(df.head())

