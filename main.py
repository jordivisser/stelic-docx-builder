from fastapi import FastAPI, Response
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import Any, Dict, List, Optional
import io
from minutes_builder import build_minutes_docx

app = FastAPI(title="Stelic Docx Builder")


class MeetingInfo(BaseModel):
    title: str
    date: str
    date_iso: Optional[str] = None
    duration_minutes: int
    attendees: List[str]
    participant_emails: List[str] = []
    transcript_id: str


class Subscore(BaseModel):
    score: float
    evidence: str


class L10Score(BaseModel):
    meeting_date: str
    duration_minutes: int
    attendees_present: List[str] = []
    attendees_missing: List[str] = []
    subscores: Dict[str, Subscore]
    overall_score: float
    flags: List[str] = []
    one_line_summary: str


class Todo(BaseModel):
    owner: str
    description: str
    due_date: Optional[str] = None
    source_quote: str
    category: str


class TodoBundle(BaseModel):
    all: List[Todo] = []
    carryover: List[Todo] = []
    new: List[Todo] = []
    total_count: int


class RetreatItem(BaseModel):
    topic: str
    context: str
    source_quote: str
    raised_by: str


class Agenda(BaseModel):
    carryover_scorecard_items: List[str] = []
    carryover_rocks_offtrack: List[str] = []
    open_todos_for_review: List[str] = []
    unresolved_issues: List[str] = []
    tabled_items: List[str] = []


class MinutesRequest(BaseModel):
    meeting: MeetingInfo
    l10_score: L10Score
    todos: TodoBundle
    retreat_items: List[RetreatItem] = []
    agenda: Agenda
    filename: str


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/build-minutes")
def build_minutes(req: MinutesRequest):
    buffer = build_minutes_docx(req.model_dump())
    return StreamingResponse(
        io.BytesIO(buffer),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{req.filename}"'}
    )