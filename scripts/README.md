# Scripts

This directory contains utility and deployment scripts for the OS3M Sheet backend.

## ☁️ Modal Deployment

We use [Modal](https://modal.com/) for serverless deployment. You can deploy the **LLM** (inference) on Modal.

### Prerequisites

1.  **Install Modal**:
    ```bash
    pip install modal
    ```
2.  **Setup Modal**:
    ```bash
    modal setup
    ```

### 1. Model Deployment (Optional)

If you wish to host the **N-ATLaS** model yourself on high-performance GPUs (A100) instead of running it locally, use `vllm_modal.py`.

1.  **Create a Hugging Face Secret**:
    Ensure you have a Modal secret named `my-huggingface-secret` containing your `HF_TOKEN` if the model is gated (though N-ATLaS is public, it's good practice).
    ```bash
    modal secret create my-huggingface-secret HF_TOKEN=your_token
    ```
2.  **Deploy the Model**:
    ```bash
    modal deploy scripts/vllm_modal.py
    ```
    *Note the URL provided by Modal (e.g., `https://your-user--natlas-vllm-serve.modal.run`). You will use this as the `BASE_URL` for the backend.*