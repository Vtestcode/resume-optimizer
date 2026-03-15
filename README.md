# Resume Optimizer

Resume Optimizer is a Streamlit app that rewrites a Word `.docx` resume against a pasted job description while preserving the document structure and formatting.

The UI lives in `web_app.py`, and the resume-processing engine runs in `app.py`.

## What It Does

- Rewrites experience bullets to maximize JD alignment while staying grounded in the original resume
- Updates summary bullets and skills to better match the target role
- Preserves Word layout, paragraph structure, and formatting during rebuild
- Generates ATS-style analysis with missing skills and keyword frequency
- Drafts a tailored cover letter from the aligned resume and JD

## Requirements

- Python 3.11 recommended
- An [OpenAI API key](https://platform.openai.com/api-keys)

## Local Setup

```bash
# 1. Create and activate a virtual environment
python -m venv venv
venv\Scripts\activate

# 2. Install dependencies
pip install -r requirements.txt

# 3. Configure environment
copy .env.example .env
# Add OPENAI_API_KEY to .env

# 4. Run the UI
streamlit run web_app.py
```

The app opens at `http://localhost:8501`.

## Usage

1. Upload your `.docx` resume.
2. Paste the full job description.
3. Click `Align Resume`.
4. Download `aligned_resume.docx`.
5. Review the ATS panel, cover letter, and bullet diff.

The UI does not expose configuration fields. It reads `OPENAI_API_KEY` from the environment and uses `OPENAI_MODEL` if set, otherwise defaults to `gpt-4o`.

