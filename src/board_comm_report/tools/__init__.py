from typing import List
from crewai.tools import BaseTool

from .sharepoint_tool import SharePointTool
from .onenote_tool import OneNoteTool

def get_tools() -> List[BaseTool]:
    """A helper function to return a list of all the tools."""
    return [SharePointTool(), OneNoteTool()]
