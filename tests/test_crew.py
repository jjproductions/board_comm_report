import pytest
from unittest.mock import patch
from board_comm_report.crew import BoardCommReport

def test_crew_kickoff():
    """
    Test that the crew's kickoff method is called.
    """
    with patch('crewai.crew.Crew.kickoff') as mock_kickoff:
        # Instantiate your crew
        crew = BoardCommReport().crew()
        # Run the kickoff method
        crew.kickoff()
        # Assert that the kickoff method was called
        mock_kickoff.assert_called_once()
