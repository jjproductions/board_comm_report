from pydantic import BaseModel
from typing import List, Optional

class CommitteeMinutes(BaseModel):
    name: str
    content: str
    has_data: bool

class CoordinatorOutput(BaseModel):
    last_board_meeting: str
    upcoming_board_meeting: str
    all_committee_data: List[CommitteeMinutes]

class ReportOutput(BaseModel):
    file_name: str
    content: str