# AI-Powered Presentation Content Generator

A Python project for the Generative AI practical exam. It accepts a topic, prepares a presentation title and objectives, creates at least six content slides with bullet points and speaker notes, and saves a `.pptx` PowerPoint file.

## Features

- Accepts a topic related to computer applications, IT, or emerging technology.
- Generates a title, objectives, and six content slides.
- Adds speaker notes to each content slide.
- Saves the presentation in the `output` folder.
- Includes an offline template fallback if no API key is configured.
- With an OpenAI API key, asks an AI model to generate topic-specific content.

## Requirements

- Python 3.10 or newer recommended
- Visual Studio Code
- Internet connection for installing packages and for live AI generation
- Optional: an OpenAI API key for live AI content generation. API usage may incur charges.

## Setup in VS Code (Windows)

1. Extract the ZIP file.
2. Open the extracted `AI_Presentation_Generator` folder in VS Code.
3. Open **Terminal → New Terminal**.
4. Install the libraries:

   ```bash
   python -m pip install -r requirements.txt
   ```

   If `python` is not recognized, try:

   ```bash
   py -m pip install -r requirements.txt
   ```

5. Run the program:

   ```bash
   python main.py
   ```

   Or:

   ```bash
   py main.py
   ```

6. Enter a topic such as `Artificial Intelligence`, `Cloud Computing`, `Cybersecurity`, or `Internet of Things`.
7. Find the PowerPoint file in the `output` folder.

## Enable live AI generation (optional)

1. Create or access your API key through your AI provider's official platform. API access may be separate from a ChatGPT subscription and may cost money.
2. Copy `.env.example` to a new file named `.env`.
3. Open `.env` and set `OPENAI_API_KEY` to your own key. Keep this file private.
4. Run `python main.py` again.

If no key is configured or the API call fails, the program uses the built-in template so the PPT can still be created. Template mode is automatic PPT creation, not live AI content generation.

## Expected output

The generated presentation normally contains 8 slides:
1. Title
2. Objectives
3. Introduction
4. Key Concepts and Features
5. How It Works
6. Applications
7. Benefits and Challenges
8. Future Scope and Conclusion

Each content slide includes speaker notes. Check the generated file in PowerPoint and review topic-specific facts before submitting it.

## Practical write-up

**Aim:** To develop an AI-based presentation generator that creates structured presentation content from a given topic.

**Algorithm:**
1. Start.
2. Accept a technology-related topic from the user.
3. Generate the presentation title and objectives.
4. Generate at least six slide headings and bullet points.
5. Add speaker notes to each content slide.
6. Create a PowerPoint presentation.
7. Save the file in PPTX format.
8. Review accuracy, logical sequence, and readability.
9. Stop.

**Result:** A PowerPoint presentation is generated with structured slide content and speaker notes.

**Conclusion:** The project demonstrates automatic presentation creation using Python and PowerPoint-generation tools. When configured with an API key, an AI model generates topic-specific content dynamically.

## Troubleshooting

- **`python` is not recognized:** Try `py` instead, or install Python and enable “Add Python to PATH”.
- **Module not found:** Run `python -m pip install -r requirements.txt`.
- **No live AI output:** Check the `.env` filename, API key, internet connection, API access, and available account credits. The program falls back to template mode if AI generation fails.
- **Speaker notes:** Open the PPT in PowerPoint and look for the Notes pane under the slide.
