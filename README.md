# Case Builder

Validate RFP facts, build a proposal storyline, and generate branded PowerPoint decks.

## Run

Requires Python 3.11+. Install dependencies with `pip install -r requirements.txt`, then run `python app.py`. Open http://localhost:5000. Set `GROQ_API_KEY` in your environment for AI interpretation and synthesis. Never commit API keys. `PORT` defaults to 5000; PDF OCR is off unless `ENABLE_PDF_OCR=true`.

## Tests

`python -m unittest discover -s tests -q`

Research uses only supplied URLs. An objective is required before drafting.

This repository contains a current source snapshot; uploaded documents, credentials, temporary files, and agent memory are excluded.
