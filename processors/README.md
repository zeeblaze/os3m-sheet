# Processors

This directory contains the core logic and processing units for OS3M Sheet. It's responsible for managing context, configuring LLMs, defining DSPy signatures, and handling various spreadsheet operations.

## Structure

- `context.py`: Manages the session context, including the `Analysis` object, which stores input data, selected ranges, and the initial prompt. This is crucial for the feedback loop, allowing the system to refine previous operations.
- `dspy_config.py`: Handles the setup and configuration of DSPy, including the initialization of the Large Language Model (LLM) to be used. It acts as the central point for defining how DSPy interacts with the chosen LLM.
- `llm.py`: Defines abstract base classes and interfaces for interacting with different Large Language Models (LLMs). This module ensures that OS3M Sheet can flexibly integrate with various LLM providers (e.g., OpenAI, Google Gemini, local models via vLLM/Ollama) by adhering to a common interface.
- `matcher.py`: Contains DSPy signatures, which are essentially structured prompts used to guide the LLM in generating specific outputs (e.g., formulas, summaries, chart configurations). It also includes logic for parsing and validating the LLM's responses, as well as utility classes for handling spreadsheet cell and section references.
- `operations.py`: Implements the business logic for each specific spreadsheet operation supported by OS3M Sheet, such as `Autofill`, `Summary`, `Formula by Example`, `Create Visual`, etc. Each operation handler processes the input, interacts with the LLM via DSPy, and formats the output for the client.

## Key Concepts

### DSPy Integration

DSPy is used extensively throughout the processors to define clear, modular, and optimizable language model programs. This allows for robust and reliable interaction with LLMs, ensuring that the model's outputs are structured and relevant to the spreadsheet tasks.

### Context Management

The `ContextManager` in `context.py` is vital for enabling iterative interactions and the feedback system. By storing the state of the most recent operation, the system can re-process queries with user feedback, allowing the LLM to self-correct and improve its output without starting from scratch.

### Modularity
The design emphasizes modularity, with each file having a distinct responsibility, making the system easier to understand, maintain, and extend with new features or LLM integrations.