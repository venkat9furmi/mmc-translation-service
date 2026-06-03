# MMC Sport Translation Service

Automated DOCX translation service for UEFA TV scripts.
Receives a Word script → translates COMMENTARY cells via Gemini AI → returns translated DOCX.

---

## Deploy to Railway (Free)

1. Push this folder to a GitHub repo
2. Go to railway.app → New Project → Deploy from GitHub
3. Add environment variable: `GEMINI_API_KEY=your_key_here`
4. Deploy → Railway gives you a public URL like: `https://mmc-translation.up.railway.app`

---

## API Endpoints

### GET /health
Check the service is running.

### POST /translate
Translate a DOCX script.

**Form fields:**
| Field      | Required | Description                          |
|------------|----------|--------------------------------------|
| file       | Yes      | The .docx script file                |
| language   | Yes      | `de` for German, `fr` for French     |
| api_key    | No       | Gemini API key (uses env var if not set) |

**Returns:** Translated .docx file as download

---

## n8n Workflow Setup

After Gmail Trigger + Gmail Get Attachment:

### Node: HTTP Request (German Translation)
```
Method:    POST
URL:       https://your-app.up.railway.app/translate
Body Type: Form Data (multipart)
Fields:
  - file:     [binary from Gmail attachment]
  - language: de
```

### Node: HTTP Request (French Translation)  
```
Method:    POST
URL:       https://your-app.up.railway.app/translate
Body Type: Form Data (multipart)
Fields:
  - file:     [binary from Gmail attachment]
  - language: fr
```

Both nodes return a .docx binary — attach to Gmail send nodes.

---

## What Gets Translated

| Table         | Column       | What                              |
|---------------|--------------|-----------------------------------|
| Running Order | Item (col 1) | Feature description paragraphs    |
| Feature tables| Commentary   | VO lines, SOT transcripts         |

### Skipped (kept as-is):
- Timecodes
- SOT/UPSOT/NATSOT labels
- [NO COMMENTARY]
- Production cues (Break Bumper, Clock Into Part 2, etc.)
- Empty cells

---

## Translation Format in Output

Each translated cell contains:
```
Original English line
DE: German translation line

Original English line 2
DE: German translation line 2
```
