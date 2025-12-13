import os
from typing import Optional, List, Any
from dotenv import load_dotenv
from fastapi import FastAPI
from pydantic import BaseModel
from processors.operations import (
    handle_autofill,
    handle_feedback,
    handle_rangesel,
    handle_summary,
    handle_formula_exp,
    handle_batchproc,
    handle_formula_pbe,
    handle_create_visual,
    handle_formula_chk
)

load_dotenv()

app = FastAPI()

# Global variable to store conversation history
conversation_history = []

class AutofillRequest(BaseModel):
    inputRange: str
    inputData: List[List[Any]]
    outputRange: str
    outputData: List[List[Any]]
    description: str

@app.post("/autofill")
async def autofill_route(request: AutofillRequest):
    msg = request.dict()
    print(f"Received autofill request: {msg}")
    request_summary = f"Action: autofill\nInput Range: {msg['inputRange']}\nOutput Range: {msg['outputRange']}\nDescription: {msg['description']}"
    result = handle_autofill(msg)
    conversation_history.append((request_summary, result))
    return {"message": "Autofill processed", "result": result}

class FeedbackRequest(BaseModel):
    feedbackMsg: str

@app.post("/feedback")
async def feedback_route(request: FeedbackRequest):
    msg = request.dict()
    print(f"Received feedback request: {msg}")
    request_summary = f"Action: feedback\nFeedback: {msg['feedbackMsg']}"
    result = handle_feedback(msg)
    conversation_history.append((request_summary, result))
    return {"message": "Feedback processed", "result": result}

class RangeselRequest(BaseModel):
    inputRange: str
    inputData: List[List[Any]]
    description: str
    
@app.post("/rangesel")
async def rangesel_route(request: RangeselRequest):
    msg = request.dict()
    print(f"Received rangesel request: {msg}")
    request_summary = f"Action: rangesel\nInput Range: {msg['inputRange']}\nDescription: {msg['description']}"
    result = handle_rangesel(msg)
    conversation_history.append((request_summary, result))
    return {"message": "Rangesel processed", "result": result}

class SummaryRequest(BaseModel):
    inputRange: str
    inputData: List[List[Any]]
    outputRange: Optional[str] = None
    outputData: Optional[List[List[Any]]] = None
    description: str

@app.post("/summary")
async def summary_route(request: SummaryRequest):
    msg = request.dict()
    print(f"Received summary request: {msg}")
    request_summary = f"Action: summary\nInput Range: {msg['inputRange']}\nDescription: {msg['description']}"
    result = handle_summary(msg)
    conversation_history.append((request_summary, result))
    return {"message": "Summary processed", "result": result}

class FormulaExpRequest(BaseModel):
    inputRange: str
    inputData: List[List[Any]]
    outputRange: Optional[str] = None
    outputData: Optional[List[List[Any]]] = None
    description: str

@app.post("/formula_exp")
async def formula_exp_route(request: FormulaExpRequest):
    msg = request.dict()
    print(f"Received formula explanation request: {msg}")
    request_summary = f"Action: formula_exp\nInput Range: {msg['inputRange']}\nDescription: {msg['description']}"
    result = handle_formula_exp(msg)
    conversation_history.append((request_summary, result))
    return {"message": "Formula explanation processed", "result": result}

class BatchprocRequest(BaseModel):
    inputRange: str
    inputData: List[List[Any]]
    description: str

@app.post("/batchproc")
async def batchproc_route(request: BatchprocRequest):
    msg = request.dict()
    print(f"Received batch processing request: {msg}")
    request_summary = f"Action: batchproc\nInput Range: {msg['inputRange']}\nDescription: {msg['description']}"
    result = handle_batchproc(msg)
    conversation_history.append((request_summary, result))
    return {"message": "Batch processing processed", "result": result}

class FormulaPBERequest(BaseModel):
    inputRange: str
    inputData: List[List[Any]]
    outputRange: str
    outputData: List[List[Any]]
    description: str

@app.post("/formula_pbe")
async def formula_pbe_route(request: FormulaPBERequest):
    msg = request.dict()
    print(f"Received formula PBE request: {msg}")
    request_summary = f"Action: formula_pbe\nInput Range: {msg['inputRange']}\nOutput Range: {msg['outputRange']}\nDescription: {msg['description']}"
    result = handle_formula_pbe(msg)
    conversation_history.append((request_summary, result))
    return {"message": "Formula PBE processed", "result": result}

class CreateVisualRequest(BaseModel):
    inputRange: str
    inputData: List[List[Any]]
    outputRange: Optional[str] = None
    outputData: Optional[List[List[Any]]] = None
    description: str

@app.post("/create_visual")
async def create_visual_route(request: CreateVisualRequest):
    msg = request.dict()
    print(f"Received create visual request: {msg}")
    request_summary = f"Action: create_visual\nInput Range: {msg['inputRange']}\nDescription: {msg['description']}"
    result = handle_create_visual(msg)
    conversation_history.append((request_summary, result))
    return {"message": "Create visual processed", "result": result}

class FormulaChkRequest(BaseModel):
    inputRange: str
    inputData: List[List[Any]]
    outputRange: Optional[str] = None
    outputData: Optional[List[List[Any]]] = None
    description: str

@app.post("/formula_chk")
async def formula_chk_route(request: FormulaChkRequest):
    msg = request.dict()
    print(f"Received formula check request: {msg}")
    request_summary = f"Action: formula_chk\nInput Range: {msg['inputRange']}\nOutput Range: {msg['outputRange']}\nDescription: {msg['description']}"
    result = handle_formula_chk(msg)
    conversation_history.append((request_summary, result))
    return {"message": "Formula check processed", "result": result}

@app.get("/history")
async def history_route():
    return {"history": conversation_history}
