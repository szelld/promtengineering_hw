# AI-Powered PDF to PPTX Generator

This Python application automatically generates a styled PowerPoint presentation (PPTX) from a long PDF document. It leverages the **Google Gemini API** (`gemini-2.5-flash`) to intelligently extract key points and organize them into structured slides with titles and bullet points.

## Features
- **Smart Summarization:** Reads a PDF and condenses it into logical slides without cutting off mid-thought.
- **Configurable Length:** Set a strict limit on the maximum number of slides generated (`MAX_SLIDES`).
- **Dynamic Theming:** Built-in support for multiple presentation styles via the `THEME` variable:
  - `"dark"`: Futuristic dark mode with neon cyan text and electric purple accent lines.
  - `"light"`: Clean white background with black text and a smooth gradient line fading from dark blue to sky blue.
  - `"none"`: Standard unstyled blank layout.

---

## Setup Tutorial

Follow these steps to get the project running on your local machine.

### 1. Prerequisites
Ensure you have Python 3 installed on your system.

### 2. Install Dependencies
Open your terminal in the project directory and run the following command to install the required Python packages:

```bash
pip install -r requirements.txt
```

### 3. Get a Google Gemini API Key
1. Go to [Google AI Studio](https://aistudio.google.com/).
2. Sign in with your Google account and generate a new API key.
3. *Note: The free tier has usage limits. If you process very large PDFs frequently, you may hit a "Quota Exceeded" error and need to wait or enable billing.*

### 4. Configure Environment Variables
To keep your API key secure, this project uses a `.env` file. 
1. In the root directory of this project, create a new file named exactly `.env`.
2. Open the `.env` file in a text editor and add your API key like this:

```env
GEMINI_API_KEY=AIzaSyYourActualApiKeyHere...
```

### 5. Configure the Script
Open `main.py` in your code editor. At the top of the file, you can modify the following configuration variables:

```python
# The PDF file you want to convert
INPUT_FILE = "beszamolo.pdf"

# Presentation style ("light", "dark", or "none")
THEME = "light"

# Maximum number of slides to generate (e.g. 10), or None for unlimited
MAX_SLIDES = 10
```

### 6. Run the Application
Make sure your target PDF (e.g., `beszamolo.pdf`) is in the same folder as the script. Then, run:

```bash
python main.py
```

The script will read the PDF, send it to Gemini, generate the slides, and save the output as `beszamolo_light_pptx.pptx` (or your configured output name). 

> **Important:** Make sure you do not have the output PowerPoint file open in PowerPoint while the script is running, or the script will crash with a "Permission denied" error when it tries to overwrite it!
