# Stelic DOCX Builder

Small authenticated FastAPI service that generates branded Word documents for
Stelic automations.

## Endpoints

| Method | Endpoint | Purpose |
| --- | --- | --- |
| GET | `/health` | Container health check |
| POST | `/build-interview-transcript` | Branded interview transcript |
| POST | `/build-interview-summary` | Branded interview summary |
| POST | `/build-minutes` | Existing executive-meeting document |

The existing `/build-minutes` endpoint remains backward-compatible.

## Authentication

Set `DOCX_TOKEN` in the container environment. Authenticated endpoints expect:

```http
X-Auth-Token: <DOCX_TOKEN>
```

If `DOCX_TOKEN` is empty, authentication is disabled for local development.

## Interview request

Both interview endpoints accept the payload already produced by the n8n
Fireflies workflow:

```json
{
  "meeting": {
    "title": "Interview Transcript - Jane Doe",
    "date": "September 3, 2026 at 10:00 AM ET",
    "date_iso": "2026-09-03T14:00:00.000Z",
    "duration_minutes": 45,
    "attendees": ["Rob Crouch", "Jane Doe"],
    "participant_emails": ["rcrouch@stelic.com", "jane@example.com"],
    "transcript_id": "fireflies-transcript-id",
    "meeting_type": "other",
    "type_confidence": null,
    "type_reasoning": "Fireflies interview transcript"
  },
  "meeting_notes": [
    {
      "topic": "Full Transcript",
      "discussion": "STELIC INTERVIEW TRANSCRIPT\n\nCandidate: Jane Doe\n...",
      "decision": "",
      "open_questions": []
    }
  ],
  "filename": "Interview Transcript - Jane Doe - fireflies-transcript-id.docx"
}
```

Extra fields sent by n8n are ignored, allowing the same payload shape that was
previously sent to `/build-minutes`.

## n8n URLs

```text
http://docx-builder:8000/build-interview-transcript
http://docx-builder:8000/build-interview-summary
```

Configure the HTTP Request node to return a file. n8n stores the response in
binary field `data`, which can be passed directly to the Zoho attachment
upload node.

## Local development

```bash
python -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --reload
```

Run the test suite:

```bash
python -m unittest discover -s tests -v
```

Build the container:

```bash
docker build -t stelic-docx-builder .
docker run --rm -p 8000:8000 \
  -e DOCX_TOKEN=development-token \
  stelic-docx-builder
```
