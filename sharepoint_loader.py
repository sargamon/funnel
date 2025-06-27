'''
sharepoint_loader.py

Provides load_the_datafile to download and pivot an Excel file from SharePoint using delegated (user login) permissions via MSAL Device Code Flow.

Requirements:
    pip install pandas msal office365-rest-python-client
'''
import io
import os
import pandas as pd
import msal
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import streamlit as st


def acquire_token_device_flow(client_id: str,
                              tenant_id: str,
                              authority: str = None,
                              scope: list = None) -> str:
    """
    Acquire an access token using MSAL Device Code Flow.
    """
    if authority is None:
        authority = f"https://login.microsoftonline.com/{tenant_id}"
    if not scope:
        raise ValueError("Scope must be provided, e.g. ['Sites.Read.All'] or [f'{site_url}/.default']")

    app = msal.PublicClientApplication(client_id, authority=authority)
    # Try silent acquisition
    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(scope, account=accounts[0])
    # Interactive device code flow if silent fails
    if not result:
        flow = app.initiate_device_flow(scopes=scope)
        if "user_code" not in flow:
            raise ValueError(f"Failed to start device flow: {flow}")
        st.info(flow["message"])
        result = app.acquire_token_by_device_flow(flow)


    if "access_token" not in result:
        raise Exception(f"Could not acquire token: {result.get('error_description')}")
    return result["access_token"]


def load_datafile(site_url: str,
                  relative_url: str,
                  client_id: str,
                  tenant_id: str,
                  scope: list) -> pd.DataFrame:
    """
    Download an Excel file from SharePoint via delegated user login,
    pivot on ST_Program × LeadStatus summing GroupTotal, and return the pivot table.
    """
    # Acquire token interactively
    token = acquire_token_device_flow(client_id, tenant_id, scope=scope)

    # Create client context with the acquired token
    ctx = ClientContext(site_url).with_access_token(token)

    # Download file into memory
    response = File.open_binary(ctx, relative_url)

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
    """
    Load configuration from environment variables and invoke delegated load_datafile.
    Required env vars:
      - SP_SITE_URL
      - SP_RELATIVE_URL
      - AZURE_CLIENT_ID
      - AZURE_TENANT_ID
    """
    SITE_URL     = os.getenv("SP_SITE_URL")
    RELATIVE_URL = os.getenv("SP_RELATIVE_URL")
    CLIENT_ID    = os.getenv("AZURE_CLIENT_ID")
    TENANT_ID    = os.getenv("AZURE_TENANT_ID")
    # Use SharePoint resource default scope
    SCOPE        = [f"{SITE_URL}/.default"]

    missing = [name for name, val in {
        "SP_SITE_URL": SITE_URL,
        "SP_RELATIVE_URL": RELATIVE_URL,
        "AZURE_CLIENT_ID": CLIENT_ID,
        "AZURE_TENANT_ID": TENANT_ID
    }.items() if not val]
    if missing:
        raise EnvironmentError(f"Missing environment vars: {', '.join(missing)}")

    return load_datafile(SITE_URL, RELATIVE_URL, CLIENT_ID, TENANT_ID, SCOPE)


if __name__ == "__main__":
    pivot_df = load_the_datafile()
    print(pivot_df)
