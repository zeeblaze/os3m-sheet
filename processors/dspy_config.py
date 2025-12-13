import os
import dspy
from dotenv import load_dotenv
from .llm import Context, LLM

class DSPyContext(Context):
    def __init__(self, lm=None):
        super().__init__()
        self.lm = lm if lm else dspy.settings.lm
        self.history = []

    def query(self, input: str) -> str:
        messages = []
        for h in self.history:
            messages.append(f"{h['role']}: {h['content']}")
        messages.append(f"User: {input}")
        
        prompt = "\n\n".join(messages)
        
        if self.lm:
            response = self.lm(prompt)[0]
            self.history.append({"role": "user", "content": input})
            self.history.append({"role": "assistant", "content": response})
            return response
        return ""

class DSPyLLM(LLM):
    def __init__(self, lm=None):
        super().__init__()
        self.lm = lm

    def getContext(self):
        context = DSPyContext(lm=self.lm)
        self.context.append(context)
        return context

def setup_dspy():
    load_dotenv()
    api_key = os.getenv("API_KEY")
    base_url = os.getenv("BASE_URL")
    model_name = os.getenv("MODEL_NAME")

    if api_key:
        model = model_name if model_name else "gpt-3.5-turbo"
        if base_url and not model.startswith("openai/"):
            model = f"openai/{model}"

        lm = dspy.LM(
            model=model,
            api_key=api_key,
            api_base=base_url
        )
        dspy.settings.configure(lm=lm)
        return lm
    return None