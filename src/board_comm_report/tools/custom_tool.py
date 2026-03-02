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
_cached_sharepoint_client: Optional[GraphServiceClient] = None
_cached_onenote_client: Optional[GraphServiceClient] = None

import asyncio

def get_graph_client(client_type: str = "sharepoint") -> Optional[GraphServiceClient]:
    global _cached_sharepoint_client, _cached_onenote_client
    
    tenant_id = os.environ.get("AZURE_TENANT_ID")
    
    if client_type == "sharepoint":
        if _cached_sharepoint_client is None:
            client_id = os.environ.get("SHAREPOINT_CLIENT_ID") or os.environ.get("AZURE_CLIENT_ID")
            client_secret = os.environ.get("SHAREPOINT_CLIENT_SECRET") or os.environ.get("AZURE_CLIENT_SECRET")
            if tenant_id and client_id and client_secret:
                credential = ClientSecretCredential(tenant_id, client_id, client_secret)
                _cached_sharepoint_client = GraphServiceClient(
                    credentials=credential, 
                    scopes=['https://graph.microsoft.com/.default']
                )
        return _cached_sharepoint_client
        
    elif client_type == "onenote":
        if _cached_onenote_client is None:
            tenant_id = os.environ.get("ONENOTE_TENANT_ID") or os.environ.get("AZURE_TENANT_ID")
            client_id = os.environ.get("ONENOTE_CLIENT_ID") or os.environ.get("AZURE_CLIENT_ID")
            client_secret = os.environ.get("ONENOTE_CLIENT_SECRET") or os.environ.get("AZURE_CLIENT_SECRET")
            if tenant_id and client_id and client_secret:
                credential = ClientSecretCredential(tenant_id, client_id, client_secret)
                _cached_onenote_client = GraphServiceClient(
                    credentials=credential, 
                    scopes=['https://graph.microsoft.com/.default']
                )
        return _cached_onenote_client
        
    return None

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

        graph_client = get_graph_client("sharepoint")
        if not graph_client:
            return json.dumps({"error": "Azure SharePoint credentials are not configured. Please set AZURE_TENANT_ID, AZURE_CLIENT_ID (or SHAREPOINT_CLIENT_ID), and AZURE_CLIENT_SECRET environment variables."})

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
    site_id: str = Field(..., description="The SharePoint Site ID.")
    section_id: str = Field(..., description="The OneNote Section ID.")
    last_meeting_date: str = Field(..., description="The date of the last board meeting in ISO format (e.g., YYYY-MM-DD).")

class OneNoteTool(BaseTool):
    name: str = "OneNote Extractor"
    description: str = (
        "A tool to extract meeting minutes from OneNote notebooks. "
        "Use this tool to get the raw text from a specific committee's meeting."
    )
    args_schema: Type[BaseModel] = OneNoteToolInput

    def _run(self, site_id: str, section_id: str, last_meeting_date: str) -> str:
        """
        Connects to OneNote and retrieves the content of pages that fall between the given date and today.
        """
        graph_client = get_graph_client("onenote")
        if not graph_client:
            return "Azure OneNote credentials are not configured. Please set AZURE_TENANT_ID, ONENOTE_CLIENT_ID, and ONENOTE_CLIENT_SECRET environment variables."

        try:
            from datetime import datetime, timezone
            
            # We must run graphite async methods in a sync wrapper for the crew tools
            async def get_page_content():
                try:
                    # Clean the date string for fromisoformat if needed
                    date_str = last_meeting_date.replace('Z', '+00:00')
                    # fromisoformat handles +00:00, or standard YYYY-MM-DD
                    if 'T' not in date_str:
                        target_date = datetime.strptime(date_str, "%Y-%m-%d").replace(tzinfo=timezone.utc)
                    else:
                        target_date = datetime.fromisoformat(date_str)
                        if target_date.tzinfo is None:
                            target_date = target_date.replace(tzinfo=timezone.utc)
                            
                    target_date_iso = target_date.strftime("%Y-%m-%dT%H:%M:%SZ")
                except Exception as e:
                    return f"Invalid last_meeting_date format: {e}"

                from msgraph.generated.sites.item.onenote.sections.item.pages.pages_request_builder import PagesRequestBuilder
                
                query_params = PagesRequestBuilder.PagesRequestBuilderGetQueryParameters(
                    top=100,
                    select=["title", "createdDateTime", "contentUrl"],
                    orderby=["createdDateTime desc"],
                    filter=f"createdDateTime ge {target_date_iso}"
                )
                request_config = RequestConfiguration(query_parameters=query_params)
                pages_response = await graph_client.sites.by_site_id(site_id).onenote.sections.by_onenote_section_id(section_id).pages.get(request_configuration=request_config)
                if not pages_response or not pages_response.value:
                    return f"No pages found in section {section_id} for site {site_id}."
                    
                now = datetime.now(timezone.utc)
                
                # Filter pages by created_date_time between last_meeting_date and today (as a safeguard/upper bound limit)
                valid_pages = []
                for page in pages_response.value:
                    page_date = getattr(page, 'created_date_time', None)
                    if not page_date:
                        continue
                        
                    if isinstance(page_date, str):
                        try:
                            p_str = page_date.replace('Z', '+00:00')
                            page_date = datetime.fromisoformat(p_str)
                        except Exception:
                            continue
                            
                    if isinstance(page_date, datetime):
                        if page_date.tzinfo is None:
                            page_date = page_date.replace(tzinfo=timezone.utc)
                        
                        if target_date <= page_date <= now:
                            valid_pages.append(page)
                
                if not valid_pages:
                    return f"No pages found between {last_meeting_date} and today in section {section_id}."
                
                content_results = []
                for page in valid_pages:
                    page_id = getattr(page, 'id', 'unknown_id')
                    try:
                        page_content = await graph_client.sites.by_site_id(site_id).onenote.pages.by_onenote_page_id(page_id).content.get()
                        
                        if isinstance(page_content, bytes):
                            text_content = page_content.decode('utf-8')
                        else:
                            text_content = str(page_content)
                            
                        page_title = getattr(page, 'title', page_id)
                        p_date = getattr(page, 'created_date_time', 'Unknown Date')
                        content_results.append(f"--- Page: {page_title} (Created: {p_date}) ---\n{text_content}")
                    except Exception as e:
                        content_results.append(f"--- Error loading page {page_id}: {str(e)} ---")
                
                return "\n\n".join(content_results)
                
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
