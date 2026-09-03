import io
import os
import re
from typing import Dict, List, Optional

from fastapi import Depends, FastAPI, Header, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field

from interview_builder import (
    build_interview_summary_docx,
    build_interview_transcript_docx,
)
from minutes_builder import build_minutes_docx


DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

app = FastAPI(
    title="Stelic DOCX Builder",
    version="2.0.0",
    description=(
        "Builds branded Stelic DOCX files for interview artifacts and "
        "backward-compatible executive meeting minutes."
    ),
)

DOCX_TOKEN = os.getenv("DOCX_TOKEN", "")


def verify_token(x_auth_token: Optional[str] = Header(default=None)):
    if DOCX_TOKEN and x_auth_token != DOCX_TOKEN:
        raise HTTPException(status_code=401, detail="Invalid token")


class MeetingInfo(BaseModel):
    title: str
    date: str
    date_iso: Optional[str] = None
    duration_minutes: int
    attendees: List[str] = Field(default_factory=list)
    participant_emails: List[str] = Field(default_factory=list)
    transcript_id: str
    meeting_type: str = "other"
    type_confidence: Optional[float] = None
    type_reasoning: Optional[str] = ""


class Subscore(BaseModel):
    score: float
    evidence: str


class Scoring(BaseModel):
    meeting_date: str
    duration_minutes: int
    attendees_present: List[str] = Field(default_factory=list)
    attendees_missing: List[str] = Field(default_factory=list)
    subscores: Dict[str, Subscore]
    overall_score: float
    flags: List[str] = Field(default_factory=list)
    one_line_summary: str


class Todo(BaseModel):
    owner: str
    description: str
    due_date: Optional[str] = None
    source_quote: str
    category: str


class TodoBundle(BaseModel):
    all: List[Todo] = Field(default_factory=list)
    carryover: List[Todo] = Field(default_factory=list)
    new: List[Todo] = Field(default_factory=list)
    total_count: int


class RetreatItem(BaseModel):
    topic: str
    context: str
    source_quote: str
    raised_by: str


class Agenda(BaseModel):
    carryover_scorecard_items: List[str] = Field(default_factory=list)
    carryover_rocks_offtrack: List[str] = Field(default_factory=list)
    open_todos_for_review: List[str] = Field(default_factory=list)
    unresolved_issues: List[str] = Field(default_factory=list)
    tabled_items: List[str] = Field(default_factory=list)


class MeetingNote(BaseModel):
    topic: str
    discussion: str
    decision: Optional[str] = ""
    open_questions: List[str] = Field(default_factory=list)


class MinutesRequest(BaseModel):
    meeting: MeetingInfo
    meeting_notes: List[MeetingNote] = Field(default_factory=list)
    scoring: Optional[Scoring] = None
    todos: Optional[TodoBundle] = None
    retreat_items: List[RetreatItem] = Field(default_factory=list)
    agenda: Optional[Agenda] = None
    filename: str


class InterviewArtifactRequest(BaseModel):
    meeting: MeetingInfo
    meeting_notes: List[MeetingNote] = Field(default_factory=list)
    filename: str


def safe_download_filename(filename: str, fallback: str) -> str:
    value = os.path.basename(str(filename or "")).strip()
    value = re.sub(r"[\x00-\x1f\x7f\"\\/]+", "-", value)
    value = value[:180].strip(" .-") or fallback
    if not value.lower().endswith(".docx"):
        value = f"{value}.docx"
    return value


def docx_response(buffer: bytes, filename: str, fallback: str) -> StreamingResponse:
    safe_name = safe_download_filename(filename, fallback)
    return StreamingResponse(
        io.BytesIO(buffer),
        media_type=DOCX_MIME,
        headers={"Content-Disposition": f'attachment; filename="{safe_name}"'},
    )


@app.get("/health")
def health():
    return {"status": "ok", "version": app.version}


@app.post("/build-minutes")
def build_minutes(req: MinutesRequest, auth: None = Depends(verify_token)):
    buffer = build_minutes_docx(req.model_dump())
    return docx_response(buffer, req.filename, "meeting-notes.docx")


@app.post("/build-interview-transcript")
def build_interview_transcript(
    req: InterviewArtifactRequest,
    auth: None = Depends(verify_token),
):
    buffer = build_interview_transcript_docx(req.model_dump())
    return docx_response(
        buffer,
        req.filename,
        "interview-transcript.docx",
    )


@app.post("/build-interview-summary")
def build_interview_summary(
    req: InterviewArtifactRequest,
    auth: None = Depends(verify_token),
):
    buffer = build_interview_summary_docx(req.model_dump())
    return docx_response(
        buffer,
        req.filename,
        "interview-summary.docx",
    )
