
# OOS Mobile Auditor v2

Phone-first version of the OOS Pick Sheet Auditor. It runs in a browser on Android, iPhone, iPad, Windows, or Mac.

## Employee workflow
1. Open the app from the phone home screen.
2. Upload the pick-sheet PDF.
3. Upload Verified Outs (Excel/CSV/PDF/photo).
4. Tap **RUN AUDIT**.
5. Review exceptions.
6. Download the two PDFs.

## Business rule
**Anything in Column A / the left column of Verified Outs is a verified OOS item.**
Column B and later columns are notes only.

- Pick sheet OS + item in Column A -> VERIFIED OOS
- Pick sheet OS + item absent from Column A -> FALSE OOS / PICKING EXCEPTION
- Uncertain handwriting -> NEEDS REVIEW

## Recommended hosting
This is a Streamlit web app. Host it once and employees use the same URL on Android or iPhone.

Good choices:
- Streamlit Community Cloud for testing
- Render / Railway / Azure App Service for internal production use
- Your existing company cloud environment if preferred

## Secure API setup
Do NOT give employees the API key and do not put it in source code.

Set a server environment variable:
OPENAI_API_KEY=...

The app uses the OpenAI Responses API for scanned PDFs/handwriting.

## Local test
Python 3.11+:

    python -m venv .venv
    # Windows:
    .venv\Scripts\activate
    # Mac/Linux:
    source .venv/bin/activate

    pip install -r requirements.txt
    streamlit run app.py

## Add to Android home screen
Open the hosted URL in Chrome -> menu -> Add to Home screen.

## Add to iPhone home screen
Open the hosted URL in Safari -> Share -> Add to Home Screen.

## Production improvements recommended next
- Employee login / PIN
- Audit history database
- Picker scorecards
- Manager correction screen that stores aliases
- Company logo
- Automatic email/report archive
