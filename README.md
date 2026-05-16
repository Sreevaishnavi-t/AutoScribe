# 📄 AutoScribe: AI Document Maker & Formatter

AutoScribe is a dual-purpose document utility built with Python and Streamlit. It leverages the power of Large Language Models (Google Gemini) to generate comprehensive document drafts, and utilizes advanced `python-docx` manipulation to ensure every generated or uploaded file perfectly adheres to strict, print-ready MS Word formatting standards.

## Core Features

*   ** AI Content Generation:** Provide a title and an outline, and AutoScribe will generate a fully fleshed-out document draft using the Gemini 2.5 Flash model.
*   ** Native Word Formatting:** Unlike web-based text editors, AutoScribe constructs native MS Word XML Styles. This ensures downloaded `.docx` files are perfectly editable without messy inline-styling clashes.
*   ** Template Inheritance:** Upload a reference document (e.g., a college syllabus or corporate template), and the AI-generated text will automatically adopt its exact fonts, margins, sizes, and layout styles.
*   ** The "Messy Document" Cleaner:** A dedicated formatting mode. Upload a poorly formatted `.docx` file, and AutoScribe will strip out excessive blank pages, normalize chaotic font sizes to a professional standard, and enforce strict A4 dimensions with 1-inch margins.

## Technology Stack

*   **Frontend/UI:** [Streamlit](https://streamlit.io/) (Python)
*   **AI Engine:** [Google Gemini API](https://aistudio.google.com/) (`google-genai`)
*   **Document Engine:** `python-docx`

## Local Installation & Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/Sreevaishnavi-t/AutoScribe.git
   cd AutoScribe
   ```

2. **Install the required dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Get your API Key:**
   Get a free Gemini API key from [Google AI Studio](https://aistudio.google.com/app/apikey). You will paste this directly into the app's sidebar.

4. **Run the Application:**
   ```bash
   streamlit run app.py
   ```
   *The app will automatically open in your default web browser at `http://localhost:8501`.*
