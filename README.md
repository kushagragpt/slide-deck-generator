# AI-Powered Slide Deck Generator

## Description
This Python script automatically generates a professional PowerPoint presentation on any given topic. It uses the Google Gemini LLM for content synthesis, real-time web search for up-to-date information, and a stock photo API for relevant visuals.

## Features
-   **Topic Input**: Generates a presentation on any user-provided topic.
-   **Live Web Search**: Uses SerpApi to gather current information.
-   **AI Content Generation**: Leverages Google Gemini to write titles, bullet points, and speaker notes.
-   **Automated Visuals**: Generates charts with Matplotlib and finds stock photos with the Pexels API.
-   **Template-Based Design**: Uses a library of `.pptx` templates for professional styling and layout.
-   **Intelligent Template Selection**: Automatically chooses the best template (e.g., business, tech, science) based on the topic.

## Setup & Installation
1.  Clone the repository.
2.  Create and activate a Python virtual environment:
    `python -m venv venv`
    `venv\Scripts\activate`
3.  Install the required dependencies:
    `pip install -r requirements.txt`
4.  Create a `config.json` file and add your API keys:
    ```json
    {
      "GEMINI_API_KEY": "YOUR_KEY_HERE",
      "SERPAPI_API_KEY": "YOUR_KEY_HERE",
      "PEXELS_API_KEY": "YOUR_KEY_HERE"
    }
    ```

## Usage
Run the script from your terminal:
`python slidegen.py --topic "Your Topic Here"`
