# AI Skill Matcher - Viva Room

A Streamlit application for practicing project vivas with questions and feedback
grounded in uploaded study material. It also includes a study-buddy and general
chatbot powered by Ollama.

## Requirements

- Python 3.11
- [Ollama](https://ollama.com/) running on the same machine as the app
- The `llama3.2` model (or another model configured in the app)

## Run locally

In PowerShell, from the project directory:

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
ollama pull llama3.2
streamlit run app.py
```

If PowerShell blocks virtual-environment activation, run the commands using
`.venv\Scripts\python.exe` directly instead.

Upload a text-based PPTX, PDF, or TXT file in the app. Scanned PDFs are not
supported. The uploaded document is sent to the configured Ollama server for
question generation, assessment, and study-buddy responses.

## GitHub and hosting

This repository can be pushed to GitHub and connected to a Streamlit hosting
service. The app requires a reachable Ollama server; a hosted Streamlit app
cannot use Ollama installed on your personal computer through `localhost`.
Configure a reachable Ollama endpoint and network access with your hosting
provider before deploying the live app. Do not commit credentials or
`.streamlit/secrets.toml`.

The viva chat history is kept in the Streamlit session. Feedback submissions
are written to `customer_feedback.json` locally; that runtime data is ignored
by Git and is not included in the repository.
