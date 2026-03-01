from crewai.tools import BaseTool
from typing import Type, List
from pydantic import BaseModel, Field
import os
from datetime import datetime
from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.models.o_data_errors.o_data_error import ODataError

class SharePointTool(BaseTool):
    name: str = "SharePoint Navigator"
    description: str = (
        "A tool to navigate and query SharePoint sites. "
        "Use this tool to find information about board meetings, like dates and times."
    )
    args_schema: Type[BaseModel] = None

    def _run(self) -> str:
        """
        Connects to SharePoint and retrieves board meeting dates.
        
        To use this tool, you need to set the following environment variables:
        - AZURE_TENANT_ID: The ID of your Azure tenant.
        - AZURE_CLIENT_ID: The client ID for your Azure AD app.
        - AZURE_CLIENT_SECRET: The client secret for your Azure AD app.
        - SHAREPOINT_SITE_ID: The ID of your SharePoint site.
        - SHAREPOINT_CALENDAR_ID: The ID of the calendar.
        """
        tenant_id = os.environ.get("AZURE_TENANT_ID")
        client_id = os.environ.get("AZURE_CLIENT_ID")
        client_secret = os.environ.get("AZURE_CLIENT_SECRET")
        site_id = os.environ.get("SHAREPOINT_SITE_ID")
        calendar_id = os.environ.get("SHAREPOINT_CALENDAR_ID")

        if not all([tenant_id, client_id, client_secret, site_id, calendar_id]):
            return "Azure and SharePoint credentials are not fully configured. Please set AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, SHAREPOINT_SITE_ID and SHAREPOINT_CALENDAR_ID environment variables."

        try:
            credential = ClientSecretCredential(tenant_id, client_id, client_secret)
            graph_client = GraphServiceClient(credentials=credential)

            now = datetime.now()
            
            # Get upcoming events
            upcoming_events = graph_client.sites.by_site_id(site_id).lists.by_list_id(calendar_id).items.get(
                query_params={"filter": f"fields/EventDate ge '{now.isoformat()}'", "orderby": "fields/EventDate asc"}
            )
            # Get past events
            past_events = graph_client.sites.by_site_id(site_id).lists.by_list_id(calendar_id).items.get(
                query_params={"filter": f"fields/EventDate lt '{now.isoformat()}'", "orderby": "fields/EventDate desc"}
            )

            next_meeting = upcoming_events.value[0].fields.additional_data['EventDate'] if upcoming_events.value else "None scheduled"
            last_meeting = past_events.value[0].fields.additional_data['EventDate'] if past_events.value else "None"
            
            return f"The last board meeting was on {last_meeting}. The next board meeting is on {next_meeting}."

        except ODataError as odata_error:
            return f"An error occurred while connecting to SharePoint: {odata_error.error.message}"
        except Exception as e:
            return f"An unexpected error occurred: {e}"

class OneNoteToolInput(BaseModel):
    """Input schema for OneNoteTool."""
    notebook_id: str = Field(..., description="The ID of the OneNote notebook.")

class OneNoteTool(BaseTool):
    name: str = "OneNote Extractor"
    description: str = (
        "A tool to extract meeting minutes from OneNote notebooks. "
        "Use this tool to get the raw text from a specific committee's meeting."
    )
    args_schema: Type[BaseModel] = OneNoteToolInput

    def _run(self, notebook_id: str) -> str:
        """
        Connects to OneNote and retrieves the content of a page.
        
        To use this tool, you need to set the following environment variables:
        - AZURE_TENANT_ID: The ID of your Azure tenant.
        - AZURE_CLIENT_ID: The client ID for your Azure AD app.
        - AZURE_CLIENT_SECRET: The client secret for your Azure AD app.
        """
        tenant_id = os.environ.get("AZURE_TENANT_ID")
        client_id = os.environ.get("AZURE_CLIENT_ID")
        client_secret = os.environ.get("AZURE_CLIENT_SECRET")

        if not all([tenant_id, client_id, client_secret]):
            return "Azure credentials are not configured. Please set AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET environment variables."

        try:
            credential = ClientSecretCredential(tenant_id, client_id, client_secret)
            graph_client = GraphServiceClient(credentials=credential)

            # Get pages in the notebook
            pages = graph_client.me.onenote.notebooks.by_notebook_id(notebook_id).pages.get()
            
            if not pages.value:
                return f"No pages found in notebook with ID {notebook_id}."

            # For simplicity, we'll just get the content of the first page.
            # In a real-world scenario, you would need to identify the correct page.
            page_id = pages.value[0].id
            page_content = graph_client.me.onenote.pages.by_onenote_page_id(page_id).content.get()

            return page_content.decode('utf-8')

        except ODataError as odata_error:
            return f"An error occurred while connecting to OneNote: {odata_error.error.message}"
        except Exception as e:
            return f"An unexpected error occurred: {e}"

def get_tools() -> List[BaseTool]:
    """A helper function to return a list of all the tools."""
    return [SharePointTool(), OneNoteTool()]
