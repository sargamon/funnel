'''
sharepoint_loader.py

Provides load_datafile to download and pivot an Excel file from SharePoint using Azure AD App credentials.

Requirements:
    pip install pandas office365-rest-python-client
'''  
import io
import os
import pandas as pd
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext

def load_datafile(site_url: str,
                  relative_url: str,
                  client_id: str,
                  client_secret: str) -> pd.DataFrame:
    """
    Download an Excel file from SharePoint via Azure AD App credentials,
    pivot on ST_Program × LeadStatus summing GroupTotal, and return the pivot table.

    Parameters:
        site_url (str): SharePoint site URL (e.g. https://yourorg.sharepoint.com/sites/Team)
        relative_url (str): Server-relative file path (e.g. /sites/Team/Shared Documents/data.xlsx)
        client_id (str): Azure AD application (client) ID
        client_secret (str): Azure AD application client secret

    Returns:
        pd.DataFrame: Pivot table with ST_Program index, LeadStatus columns, and summed GroupTotal.

    Raises:
        Exception: If download or parsing fails.
    """
    # Authenticate using client credentials
    creds = ClientCredential(client_id, client_secret)
    ctx = ClientContext(site_url).with_credentials(creds)

    # Download file into memory
    response = ctx.web.get_file_by_server_relative_url(relative_url).download().execute_query()

    # Read into DataFrame
    df = pd.read_excel(io.BytesIO(response.content), sheet_name=0)

    # Pivot as requested
    pivot_df = (
        df
        .pivot_table(
            index='ST_Program',
            columns='LeadStatus',
            values='GroupTotal',
            aggfunc='sum'
        )
        .fillna(0)
    )
    return pivot_df

def load_the_datafile() -> pd.DataFrame:
    global SITE_URL      = os.getenv("SP_SITE_URL")
    global RELATIVE_URL  = os.getenv("SP_RELATIVE_URL")
    global CLIENT_ID     = os.getenv("AZURE_CLIENT_ID")
    global CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")

    missing = [name for name, val in {
        "SP_SITE_URL": SITE_URL,
        "SP_RELATIVE_URL": RELATIVE_URL,
        "AZURE_CLIENT_ID": CLIENT_ID,
        "AZURE_CLIENT_SECRET": CLIENT_SECRET
    }.items() if not val]
    if missing:
        print(f"Missing environment vars: {', '.join(missing)}")
        exit(1)
    return load_datafile(SITE_URL, RELATIVE_URL, CLIENT_ID, CLIENT_SECRET)
    

if __name__ == "__main__":
    import os
    load_the_datafile()
