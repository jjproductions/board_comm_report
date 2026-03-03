from crewai.tools import BaseTool
from typing import Type
from pydantic import BaseModel, Field
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from kiota_abstractions.base_request_configuration import RequestConfiguration

import asyncio
from .graph_client import get_graph_client

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
                    select=["id", "title", "createdDateTime", "contentUrl"],
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
                    if page_id is None or page_id == 'unknown_id':
                        content_results.append(f"--- Error loading page ID for site: {site_id} ---")
                        continue
                        
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
