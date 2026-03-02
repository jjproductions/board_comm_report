from crewai.tools import BaseTool
from typing import Type, List, Optional
from pydantic import BaseModel, Field
import os
import json
from datetime import datetime
from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from kiota_abstractions.base_request_configuration import RequestConfiguration
from kiota_abstractions.headers_collection import HeadersCollection
from msgraph.generated.sites.item.lists.item.items.items_request_builder import ItemsRequestBuilder

# Module-level cache for the Microsoft Graph client to avoid redundant authentication
# and to prevent Pydantic serialization issues within CrewAI tools.
_cached_graph_client: Optional[GraphServiceClient] = None

import asyncio

def get_graph_client() -> Optional[GraphServiceClient]:
    global _cached_graph_client
    if _cached_graph_client is None:
        tenant_id = os.environ.get("AZURE_TENANT_ID")
        client_id = os.environ.get("AZURE_CLIENT_ID")
        client_secret = os.environ.get("AZURE_CLIENT_SECRET")
        if tenant_id and client_id and client_secret:
            credential = ClientSecretCredential(tenant_id, client_id, client_secret)
            # You must provide the .default scope for client credentials
            _cached_graph_client = GraphServiceClient(
                credentials=credential, 
                scopes=['https://graph.microsoft.com/.default']
            )
    return _cached_graph_client

class SharePointToolInput(BaseModel):
    pass

class SharePointTool(BaseTool):
    name: str = "SharePoint Navigator"
    description: str = (
        "A tool to navigate and query SharePoint sites. "
        "Use this tool to find information about board meetings, like dates and times."
    )
    args_schema: Type[BaseModel] = SharePointToolInput

    def _run(self, **kwargs) -> str:
        """
        Connects to SharePoint and retrieves board meeting dates.
        """
        site_id = os.environ.get("SHAREPOINT_SITE_ID", "instituteofmusic.sharepoint.com:/sites/BoardConnect:")
        calendar_id = os.environ.get("SHAREPOINT_CALENDAR_ID", "7b632d43-ee15-4950-a7a1-1f6be09caea5")
        
        if not site_id or not calendar_id:
            return json.dumps({"error": "SharePoint credentials are not fully configured. Please set SHAREPOINT_SITE_ID and SHAREPOINT_CALENDAR_ID environment variables."})

        graph_client = get_graph_client()
        if not graph_client:
            return json.dumps({"error": "Azure credentials are not configured. Please set AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET environment variables."})

        try:
            now = datetime.now()
            # MS Graph OData queries require ISO format with Z for UTC, and no microseconds
            now_str = now.replace(microsecond=0).isoformat() + "Z"
            
            # OData v4 query parameters
            # Create typed query parameter object and configuration instead of raw dict
            
            headers = HeadersCollection()
            headers.add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")

            upcoming_query = ItemsRequestBuilder.ItemsRequestBuilderGetQueryParameters(
                filter=f"fields/EventDate ge '{now_str}'",
                orderby=["fields/EventDate asc"],
                top=10,
                expand=["fields($select=Title,EventDate)"]
            )
            upcoming_config = RequestConfiguration(query_parameters=upcoming_query, headers=headers)
            
            past_query = ItemsRequestBuilder.ItemsRequestBuilderGetQueryParameters(
                filter=f"fields/EventDate lt '{now_str}'",
                orderby=["fields/EventDate desc"],
                top=10,
                expand=["fields($select=Title,EventDate)"]
            )
            past_config = RequestConfiguration(query_parameters=past_query, headers=headers)

            # Use asyncio to run the async graph client calls synchronously for the crew base tool
            async def run_queries():
                u_events = await graph_client.sites.by_site_id(site_id).lists.by_list_id(calendar_id).items.get(request_configuration=upcoming_config)
                p_events = await graph_client.sites.by_site_id(site_id).lists.by_list_id(calendar_id).items.get(request_configuration=past_config)
                return u_events, p_events

            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                upcoming_events, past_events = loop.run_until_complete(run_queries())
            finally:
                loop.close()

            def find_valid_meeting(events_collection):
                if not events_collection or not events_collection.value:
                    return None
                
                for event in events_collection.value:
                    fields = event.fields
                    if not fields or not fields.additional_data:
                        continue
                    
                    title = fields.additional_data.get('Title', '')
                    title_lower = title.lower()
                    
                    if 'board meeting' in title_lower and 'postpone' not in title_lower and 'cancel' not in title_lower:
                        
                        date_val = fields.additional_data.get('EventDate', 'Unknown Date')
                        if isinstance(date_val, datetime):
                            date_val = date_val.isoformat()
                            
                        return {
                            "title": title,
                            "date": date_val
                        }
                        
                return None

            next_meeting = find_valid_meeting(upcoming_events)
            last_meeting = find_valid_meeting(past_events)
            
            return json.dumps({
                "last_board_meeting": last_meeting,
                "next_board_meeting": next_meeting
            }, indent=2)

        except ODataError as odata_error:
            error_msg = odata_error.error.message if odata_error.error else str(odata_error)
            return json.dumps({"error": f"An error occurred while connecting to SharePoint: {error_msg}"})
        except Exception as e:
            return json.dumps({"error": f"An unexpected error occurred: {e}"})

class OneNoteToolInput(BaseModel):
    """Input schema for OneNoteTool."""
    user_id: str = Field(..., description="The ID or User Principal Name (UPN) of the user who owns the OneNote notebook. Required because app-only authentication doesn't have a default user context.")
    notebook_id: str = Field(..., description="The ID of the OneNote notebook.")

class OneNoteTool(BaseTool):
    name: str = "OneNote Extractor"
    description: str = (
        "A tool to extract meeting minutes from OneNote notebooks. "
        "Use this tool to get the raw text from a specific committee's meeting."
    )
    args_schema: Type[BaseModel] = OneNoteToolInput

    def _run(self, user_id: str, notebook_id: str) -> str:
        """
        Connects to OneNote and retrieves the content of a page.
        """
        graph_client = get_graph_client()
        if not graph_client:
            return "Azure credentials are not configured. Please set AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET environment variables."

        try:
            # We must run graphite async methods in a sync wrapper for the crew tools
            async def get_page_content():
                # 1. Get the notebook's sections
                sections_response = await graph_client.users.by_user_id(user_id).onenote.notebooks.by_notebook_id(notebook_id).sections.get()
                if not sections_response or not sections_response.value:
                    return f"No sections found in notebook {notebook_id} for user {user_id}."
                
                first_section_id = sections_response.value[0].id
                
                # 2. Get pages from that section
                pages_response = await graph_client.users.by_user_id(user_id).onenote.sections.by_onenote_section_id(first_section_id).pages.get()
                if not pages_response or not pages_response.value:
                    return f"No pages found in the first section of notebook {notebook_id}."
                    
                # 3. Get the first page's content
                page_id = pages_response.value[0].id
                page_content = await graph_client.users.by_user_id(user_id).onenote.pages.by_onenote_page_id(page_id).content.get()
                
                if isinstance(page_content, bytes):
                    return page_content.decode('utf-8')
                return str(page_content)
                
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                content = loop.run_until_complete(get_page_content())
                return content
            finally:
                loop.close()

        except ODataError as odata_error:
            error_msg = odata_error.error.message if odata_error.error else str(odata_error)
            return f"An error occurred while connecting to OneNote: {error_msg}"
        except Exception as e:
            return f"An unexpected error occurred: {e}"

def get_tools() -> List[BaseTool]:
    """A helper function to return a list of all the tools."""
    return [SharePointTool(), OneNoteTool()]
