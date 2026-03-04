from dotenv import load_dotenv
import json
from unittest.mock import patch
from board_comm_report.crew import BoardCommReport
from crewai import Crew, Process

load_dotenv()

def run_single_task():
    # Mock the SharePointTool._run method
    with patch('board_comm_report.tools.custom_tool.SharePointTool._run') as mock_run:
        # Define the mock return value
        mock_return_value = {
            "last_board_meeting": {
                "title": "Q1 Board Meeting",
                "date": "2025-10-15T14:00:00Z"
            },
            "next_board_meeting": {
                "title": "Q2 Board Meeting",
                "date": "2026-04-15T14:00:00Z"
            }
        }
        mock_run.return_value = json.dumps(mock_return_value)

        # Instantiate the base crew class to access our configs
        board_crew = BoardCommReport()
        
        # Grab just the agent and task we want to test
        coordinator = board_crew.meeting_coordinator()
        collection_task = board_crew.collection_objective()
        
        # Create a temporary crew with ONLY this agent and task.
        # We use sequential process to bypass the supervisor logic for this isolated test.
        test_crew = Crew(
            agents=[coordinator],
            tasks=[collection_task],
            process=Process.sequential,
            verbose=True
        )
        
        print("Kick-starting test crew...")
        result = test_crew.kickoff()
        
        print("\n\n======== FINAL RESULT ========\n")
        print(result)

        # Assert that the result is the same as the mock return value
        assert result == mock_return_value, f"Expected {mock_return_value}, but got {result}"
        print("\n\n======== ASSERTION PASSED ========\n")


if __name__ == '__main__':
    run_single_task()
