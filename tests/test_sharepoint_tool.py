import pytest
import json
import os
import sys

# Add the src folder to the python path so it runs smoothly locally
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from unittest.mock import patch, MagicMock

from board_comm_report.tools.custom_tool import SharePointTool

@pytest.fixture
def mock_graph_client():
    with patch('board_comm_report.tools.custom_tool.get_graph_client') as mock_get_client:
        mock_client = MagicMock()
        mock_get_client.return_value = mock_client
        yield mock_client

@pytest.fixture
def mock_env_vars():
    with patch.dict(os.environ, {
        "SHAREPOINT_SITE_ID": "instituteofmusic.sharepoint.com:/sites/BoardConnect",
        "SHAREPOINT_CALENDAR_ID": "7b632d43-ee15-4950-a7a1-1f6be09caea5",
        "AZURE_TENANT_ID": "mock",
        "AZURE_CLIENT_ID": "mock",
        "AZURE_CLIENT_SECRET": "mock"
    }):
        yield

def create_mock_event(title, date):
    event = MagicMock()
    event.fields.additional_data = {
        'Title': title,
        'EventDate': date
    }
    return event

def test_sharepoint_tool_retrieval(mock_graph_client, mock_env_vars):
    """
    Test that the SharePointTool retrieves and filters board meeting dates correctly.
    """
    # Create mock events for upcoming meetings
    upcoming_events_collection = MagicMock()
    upcoming_events_collection.value = [
        create_mock_event("Finance Committee", "2026-04-01T10:00:00Z"), # Ignored
        create_mock_event("Board Meeting - Postponed", "2026-04-10T10:00:00Z"), # Ignored
        create_mock_event("Q2 Board Meeting", "2026-04-15T14:00:00Z"), # Valid next meeting
        create_mock_event("Q3 Board Meeting", "2026-07-15T14:00:00Z") # Valid but shouldn't be selected (it stops at earliest valid one)
    ]
    
    # Create mock events for past meetings
    past_events_collection = MagicMock()
    past_events_collection.value = [
        create_mock_event("Cancelled Board Meeting", "2025-11-15T14:00:00Z"), # Ignored
        create_mock_event("Q1 Board Meeting", "2025-10-15T14:00:00Z"), # Valid last meeting
    ]

    # Setup the mock chain:
    # graph_client.sites.by_site_id(site_id).lists.by_list_id(calendar_id).items.get(...)
    mock_items = mock_graph_client.sites.by_site_id.return_value.lists.by_list_id.return_value.items
    
    # We need to differentiate between the two gets based on query params or just side_effect
    # The first call is upcoming, second is past.
    mock_items.get.side_effect = [
        upcoming_events_collection,
        past_events_collection
    ]

    # Initialize and run the tool
    tool = SharePointTool()
    result_json = tool._run()
    
    # Parse the result
    result = json.loads(result_json)
    
    # Assertions
    assert "error" not in result

    assert result["next_board_meeting"] is not None
    assert result["next_board_meeting"]["title"] == "Q2 Board Meeting"
    assert result["next_board_meeting"]["date"] == "2026-04-15T14:00:00Z"
    
    assert result["last_board_meeting"] is not None
    assert result["last_board_meeting"]["title"] == "Q1 Board Meeting"
    assert result["last_board_meeting"]["date"] == "2025-10-15T14:00:00Z"

    # Verify that the msgraph client was called with correct filter and expand params
    calls = mock_items.get.call_args_list
    assert len(calls) == 2
    
    # Check upcoming call
    upcoming_call_kwargs = calls[0].kwargs['query_params']
    assert "fields/EventDate ge" in upcoming_call_kwargs['filter']
    assert upcoming_call_kwargs['top'] == 10
    assert upcoming_call_kwargs['expand'] == "fields($select=Title,EventDate)"
    
    # Check past call
    past_call_kwargs = calls[1].kwargs['query_params']
    assert "fields/EventDate lt" in past_call_kwargs['filter']
    assert past_call_kwargs['top'] == 10
    assert past_call_kwargs['expand'] == "fields($select=Title,EventDate)"

def test_sharepoint_tool_no_valid_meetings(mock_graph_client, mock_env_vars):
    """
    Test when no valid board meetings are found.
    """
    empty_collection = MagicMock()
    empty_collection.value = []

    mock_items = mock_graph_client.sites.by_site_id.return_value.lists.by_list_id.return_value.items
    mock_items.get.side_effect = [empty_collection, empty_collection]

    tool = SharePointTool()
    result_json = tool._run()
    
    result = json.loads(result_json)
    assert result["next_board_meeting"] is None
    assert result["last_board_meeting"] is None

def test_sharepoint_tool_missing_credentials():
    """
    Test that the tool fails gracefully when environment variables are not set.
    """
    with patch.dict(os.environ, {}, clear=True):
        tool = SharePointTool()
        result_json = tool._run()
        result = json.loads(result_json)
        assert "error" in result
        assert "Azure credentials are not configured" in result["error"]
