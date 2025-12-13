# OS3M Sheet

### This project is an entry into the [Awarri](https://www.awarri.com/) [N-ATLaS](https://huggingface.com/NCAIR1/N-ATLaS) Developer Challenge 2025.

OS3M Sheet is an advanced spreadsheet assistant that integrates Large Language Models (LLMs) directly into LibreOffice Calc (with Excel Add-in support planned for the future). It enables users to perform complex data operations, formula generation, and visualization using natural language prompts.

- 📝[Goto](#system-requirements) for list of requirements.
#### 📁Note: I used Minis-Forum UM790-Pro in this project with a Ryzen 9 CPU, 32GB RAM and up to 16Gb configurable iGPU RAM. Most inference was run on CPU with Ollama and it's fast enough. If you don't have these, it is recommended to launch on any cloud compute service provider of your choice e.g modal, google cloud, etc.

## Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Architecture](#architecture)
- [Project Structure](#project-structure)
- [System Requirements](#system-requirements)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [API Reference](#api-reference)

## Overview
This project bridges the gap between traditional spreadsheets and AI. It uses a Python backend (FastAPI) to process requests via DSPy and an LLM (like, N-ATLaS from NCAIR1 (via vLLM or Ollama locally), Gemini or GPT), and a LibreOffice Basic/Python script to interact with the spreadsheet UI via UNO (Universal Network Objects). While currently optimized for LibreOffice via UNO, future updates will include a dedicated Excel Add-in.

The system is designed to support **[N-ATLaS](https://huggingface.com/NCAIR1/N-ATLaS) from NCAIR1** as the primary model (via vLLM or Ollama locally), alongside other DSPy-compatible models.

## Features
- **Autofill**: Intelligently fills cells based on input patterns and natural language descriptions.
- **Formula By Example (PBE)**: Infers complex formulas by providing input data and desired output examples.
- **Formula Explanation**: Translates cryptic spreadsheet formulas into plain English.
- **Formula Compatibility Check**: Validates formulas for cross-compatibility (e.g., Excel vs. LibreOffice).
- **Batch Processing**: Performs bulk data transformations.
- **Data Summarization**: Generates textual summaries of data ranges.
- **Smart Selection**: Highlights cells (Green/Red/Yellow) based on semantic criteria.
- **Visualization**: Automatically creates charts (Bar, Line, Pie, etc.) from data.
- **Feedback System**: Allows users to refine AI outputs through a feedback loop.

## Architecture
The system operates on a Client-Server model:

1.  **Client (LibreOffice)**:
    -   Script: `main.py` (packaged in Extension)
    -   Technology: Python UNO (Universal Network Objects).
    -   Role: Renders the GUI dialog, reads cell data, sends HTTP requests, and updates the spreadsheet.
    -   *Note*: Uses standard `urllib` to avoid external dependency issues within the LibreOffice Python environment.

2.  **Server (Backend)**:
    -   Entry Point: `api.py`
    -   Technology: FastAPI, Uvicorn.
    -   Role: Handles business logic, maintains conversation history, and routes tasks to the AI processor.

3.  **AI Processor**:
    -   Technology: DSPy (Declarative Self-improving Language Programs).
    -   Role: Manages prompts, context, and structured outputs (JSON) from the LLM.

### Context & Feedback Loop
The project employs a `ContextManager` to enable iterative interactions:
- **Session State**: The system temporarily stores the `Analysis` object of the most recent operation, which includes the input data, selected ranges, and the initial prompt.
- **Refinement**: When a user submits feedback, this context is retrieved. The feedback is injected into the DSPy signature as a constraint or hint, and the query is re-processed. This allows the model to self-correct or adjust its output based on user critique (e.g., "Use a different formula" or "Format as currency") without restarting the task.

## Project Structure
```text
os3m_sheet/
├── api.py                  # FastAPI application and route definitions
├── extensions/             # LibreOffice Extension source files
│   └── LibreOffice/        # Source for the .oxt extension
├── .env                    # Configuration file (API keys)
├── scripts/                # Utility and deployment scripts
│   └── vllm_modal.py       # Deploy the model on modal
├── installables/             # Folder containaing prepackaged extensions
└── processors/             # Core logic package
    ├── context.py          # Context management for feedback loops
    ├── dspy_config.py      # DSPy setup and LLM configuration
    ├── llm.py              # Abstract base classes
    ├── matcher.py          # DSPy signatures (prompts) and data parsing
    └── operations.py       # Operation handlers (Autofill, Summary, etc.)
```

## System Requirements
Operating System: Windows, Linux, or macOS.
Software: LibreOffice 7.0 or later.
Python: Python 3 embedded within LibreOffice (usually included).

### Backend Server
Operating System: Windows, Linux, or macOS.
Python: Python 3.8 or higher.

### Local LLM Inference (Optional)
If running models like N-ATLaS locally via vLLM or Ollama: 
GPU: NVIDIA GPU with CUDA support (16GB+ VRAM recommended for decent performance).
RAM: 16GB+ system RAM.
Software: Ollama or vLLM.

## Installation

### Backend Requirements
Ensure you have Python 3.8+ installed.
```bash
pip install -r requirements.txt
```

### LibreOffice Setup
To install the extension:
1.  Download the os3msheet.oxt extension file files from installables folder or clone the repo.
2.  Open LibreOffice.
2.  Go to **Tools** > **Extensions**.
3.  Click **Add**, locate and select the `os3msheet.oxt` file, and restart LibreOffice.

## Configuration
Create a `.env` file in the root directory of the project:

```ini
# For OpenAI API models
API_KEY=your_llm_api_key
BASE_URL=https://api.openai.com/v1  # or your custom endpoint
MODEL_NAME=gpt-4o                   # or gemini-2.5-pro, etc.

# For vLLM/Ollama based models
API_KEY=your_api_key                # e.g., "EMPTY" for local vLLM/Ollama
BASE_URL=http://localhost:11434/v1  # vLLM or Ollama endpoint
MODEL_NAME=n-atlas                  # The name of the model you served with vLLM/Ollama N-ATLaS from NCAIR1, or other DSPy supported models
```

## Usage

1.  **Start the Backend Server**:
    ```bash
    uvicorn api:app --reload
    ```
    The server listens on `http://127.0.0.1:8000`.

2.  **Run in LibreOffice**:
    -   Open LibreOffice Calc.
    -   **Extension**: Click the **OS3M Sheet** menu bar (if installed via Extension).

3.  **Operation**:
    -   Select a range of cells in your sheet.
    -   The dialog will auto-populate the "Input Range".
    -   Select an **Action** (e.g., `create_visual`).
    -   Provide a **Description** (e.g., "Plot a bar chart of sales over time").
    -   Click **Execute**.

## API Reference

| Endpoint | Method | Description |
| :--- | :--- | :--- |
| `/autofill` | POST | Fills output range based on input patterns. |
| `/rangesel` | POST | Returns cell colors for highlighting based on criteria. |
| `/summary` | POST | Returns a text summary of the input data. |
| `/formula_exp` | POST | Explains the logic of provided formulas. |
| `/batchproc` | POST | Transforms input data in-place. |
| `/formula_pbe` | POST | Generates formulas from input/output examples. |
| `/create_visual` | POST | Returns chart configuration (title, type). |
| `/formula_chk` | POST | Checks for formula errors or compatibility issues. |
| `/feedback` | POST | Sends user feedback to refine the previous context. |
| `/history` | GET | Retrieves the session conversation history.
