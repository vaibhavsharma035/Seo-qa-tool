# SEO QA Validator — Backend API

Flask backend for the SEO QA Validator Chrome Extension.  
Audits live URLs against expected H1, Title Tag, and Meta Description from an Excel file.

## Endpoints

| Method | Route | Description |
|--------|-------|-------------|
| POST | `/upload` | Upload Excel file, returns sheet names |
| POST | `/get-columns` | Get columns for a sheet + header row |
| POST | `/run-test` | Run full SEO audit, returns results |
| GET  | `/download/<file>` | Download colour-coded Excel report |
| GET  | `/` | Health check |

## Local Development

```bash
pip install -r requirements.txt
python app.py
```

Server runs at: http://127.0.0.1:5000
