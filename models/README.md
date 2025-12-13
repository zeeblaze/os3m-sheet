# Models

This directory is intended to house various Large Language Models (LLMs) or configurations related to them, which can be used with OS3M Sheet.

## Supported Models

OS3M Sheet is designed to be flexible and can integrate with any DSPy-compatible LLM. Currently, it is optimized for:

- **N-ATLaS from NCAIR1**: This is the primary recommended model, especially when served locally via `vLLM` or `Ollama`.
  - **Hugging Face**: You can find N-ATLaS models on [Hugging Face](https://huggingface.com/NCAIR1/N-ATLaS).
- **OpenAI Models**: Such as `gpt-4o`, `gpt-3.5-turbo`, etc.
- **Google Gemini Models**: Such as `gemini-2.5-flash`.

## Local LLM Setup (vLLM/Ollama)

For optimal performance and privacy, it is recommended to run N-ATLaS or other compatible models locally using `vLLM` or `Ollama`. Refer to the main `README.md` for configuration details on setting `BASE_URL` and `MODEL_NAME` in your `.env` file.