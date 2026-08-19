# UDD → CRR Document Automation

A web application that automatically extracts information from a **User Design Document (UDD)** and populates the corresponding blank fields in a **Code Review Record (CRR)** template — while fully preserving the original CRR formatting.

## Quick Start

```bash
cd udd-crr-app
npm install
npm start
```

Then open **http://localhost:3000** in your browser.

## How It Works

1. **Upload** your UDD (`.docx`) and CRR template (`.docx`)
2. **Review** the extracted fields on the confirmation screen — edit any field that was not found or incorrect
3. **Confirm** — the app generates a new DOCX with only the blank fields populated
4. **Download** the completed CRR

## Fields Extracted from UDD → Populated in CRR

| Field | CRR Location |
|---|---|
| Name | First page |
| CRR Document Title | Section 1.1 |
| UDD Creation Date | Section 1.1 |
| Development Type | Section 1.1 |
| Reviewer | Section 1.2 |
| Developer Function | Section 1.2 |
| Developer Name | Section 1.2 |

## Architecture

```
server.js        — Express server, file upload, API routes, session management
extractor.js     — Intelligent UDD text extraction (label matching, CRR pattern)
populator.js     — CRR DOCX XML surgery (preserves all formatting)
public/index.html — Single-page frontend (upload → preview → download)
test.js          — Integration tests (run with: node test.js)
```

## Key Design Decisions

- **OOXML direct manipulation** (`pizzip`): the CRR DOCX XML is edited surgically — only blank target cells are populated; all fonts, tables, borders, margins, headers/footers remain untouched.
- **Intelligent label matching**: fields are located by their label text using case-insensitive fuzzy matching, not by fixed page/paragraph numbers.
- **No hallucination**: if a field cannot be identified, the user is shown an error and must supply the value — nothing is guessed.
- **Session-based temp storage**: uploaded files and generated output are scoped to a UUID session and cleaned up automatically after 1 hour (or after download).

## Requirements

- Node.js 18+
- No external database required
