import re
import json
from processors.matcher import Analysis, cellPattern, formulaList
from processors.context import ContextManager
from processors.dspy_config import setup_dspy, DSPyLLM

PLATFORM = "libreoffice"

lm = setup_dspy()
llm = DSPyLLM(lm=lm)
context_manager = ContextManager()

def _parse_json(reply: str):
    try:
        clean_reply = re.sub(r'^```json\s*', '', reply, flags=re.MULTILINE)
        clean_reply = re.sub(r'^```\s*', '', clean_reply, flags=re.MULTILINE)
        clean_reply = re.sub(r'\s*```$', '', clean_reply, flags=re.MULTILINE)
        return json.loads(clean_reply)
    except Exception as e:
        print(f"JSON Parse Error: {e} for reply: {reply}")
        return None

def _flatten_input(data):
    flat_data = []
    if isinstance(data, dict):
        for val in data.values():
            if isinstance(val, list):
                data = val
                break
    if isinstance(data, list):
        if data and isinstance(data[0], list):
            for row in data:
                flat_data.extend(row)
        else:
            flat_data = data
    return flat_data

def apply_reply(analysis, reply: str, forceFormula: bool = False, target: str = 'output'):
    data = _parse_json(reply)
    section = analysis.outputSection if target == 'output' else analysis.inputSection
    if section is None:
        return []

    cell_contents = _flatten_input(data)

    index = 0
    for r in range(len(section.data)):
        for c in range(len(section.data[r])):
            if index >= len(cell_contents):
                break
            
            val = cell_contents[index]
            content = str(val).strip() if val is not None else ""

            has_formula = any(substr in cell for cell in analysis.inputSection.data[r] for substr in formulaList) if r < len(analysis.inputSection.data) else False
            if forceFormula or has_formula:
                if content and content[0] != '=':
                    content = f"={content}"

            section.data[r][c] = content
            is_formula_like = False
            if isinstance(content, str):
                if cellPattern.search(content):
                    is_formula_like = True
                else:
                    upper_content = content.upper()
                    for f in formulaList:
                        if f + "(" in upper_content:
                            is_formula_like = True
                            break
            if (forceFormula or has_formula or is_formula_like) and isinstance(content, str) and content and not content.startswith('='):
                content = f"={content}"
            if isinstance(content, str) and content.startswith('=') and PLATFORM == "libreoffice":
                content = re.sub(r',(?=(?:[^"]*"[^"]*")*[^"]*$)', ';', content)
            section.data[r][c] = content
            index += 1
    return section.data

def apply_formula_chk(reply: str):
    data = _parse_json(reply)
    warns = []
    passes = []
    if data and isinstance(data, dict):
        warns = data.get("issues", [])
        if data.get("passed", False):
            passes = ["No compatibility issues found."]
    return warns, passes

def apply_create_visual(reply: str):
    data = _parse_json(reply)
    title = "Chart"
    cType = "Line"
    if data and isinstance(data, dict):
        title = data.get("title", title)
        cType = data.get("type", cType)
    return title, cType

def apply_colors(reply: str):
    data = _parse_json(reply)
    if isinstance(data, dict) and "colors" in data:
        data = data["colors"]
    flat_data = _flatten_input(data)
    return [str(c).strip().lower() for c in flat_data]

def apply_summary(reply: str):
    data = _parse_json(reply)
    if data and isinstance(data, dict):
        return data.get("summary", reply)
    return reply

def apply_explanation(reply: str):
    data = _parse_json(reply)
    if data and isinstance(data, dict):
        return data.get("explanation", reply)
    return reply

def handle_autofill(msg):
    analysis = Analysis(msg)
    context = llm.getContext()
    context_manager.set_last_context(context)
    context_manager.set_last_analysis(analysis)
    reply = analysis.run_query()
    cell_candidate = apply_reply(analysis, reply)
    
    reply = {
        "status": "ok",
        "range": analysis.outputSection.range,
        "candidate": cell_candidate,
    }

    print(reply)
    return reply

def handle_feedback(msg):
    context = context_manager.get_last_context()
    analysis = context_manager.get_last_analysis()
    if not context or not analysis:
        return {
            "status": "error",
            "message": "Feedback can only be provided after an operation that supports it. Please perform an operation first. "
                       "Note: The 'rangesel' and 'batchproc' operations do not support feedback."
        }

    if not getattr(analysis, 'outputSection', None):
        return {
            "status": "error",
            "message": "Feedback is not supported for the last operation (no output range)."
        }

    try:
        analysis.feedback = msg["feedbackMsg"]
        reply = analysis.run_query()
        print(reply)
        cell_candidate = apply_reply(analysis, reply)

        reply = {
            "status": "ok",
            "range": analysis.outputSection.range,
            "candidate": cell_candidate,
        }

        print(reply)
        return reply
    except Exception as e:
        print(f"Exception in handle_feedback: {e}")
        return {
            "status": "error",
            "message": f"An error occurred while processing feedback: {str(e)}"
        }

def handle_rangesel(msg):
    analysis = Analysis(msg)
    context = llm.getContext()
    context_manager.set_last_context(context)
    context_manager.set_last_analysis(analysis)
    reply = analysis.run_range_sel_query()
    colors = apply_colors(reply)
    reply = {
        "status": "ok",
        "range": analysis.inputSection.range,
        "colors": colors,
    }
    print(reply)
    return reply

def handle_summary(msg):
    analysis = Analysis(msg)
    context = llm.getContext()
    context_manager.set_last_context(context)
    context_manager.set_last_analysis(analysis)
    reply = analysis.run_summary_query()
    summary_text = apply_summary(reply)
    reply = {
        "status": "ok",
        "reply": summary_text
    }

    print(reply)
    return reply

def handle_formula_exp(msg):
    print(f"DEBUG: handle_formula_exp received msg: {msg}")
    analysis = Analysis(msg)
    context = llm.getContext()
    context_manager.set_last_context(context)
    context_manager.set_last_analysis(analysis)
    reply = analysis.run_exp_explain_query()
    explanation_text = apply_explanation(reply)
    reply = {
        "status": "ok",
        "reply": explanation_text
    }

    print(reply)
    return reply

def handle_batchproc(msg):
    analysis = Analysis(msg)
    context = llm.getContext()
    context_manager.set_last_context(context)
    context_manager.set_last_analysis(analysis)
    reply = analysis.run_batchproc_query()
    cell_candidate = apply_reply(analysis, reply, target='input')
    reply = {
        "status": "ok",
        "range": analysis.inputSection.range,
        "candidate": cell_candidate,
    }
    print(reply)
    return reply

def handle_formula_pbe(msg):
    try:
        analysis = Analysis(msg)
        context = llm.getContext()
        context_manager.set_last_context(context)
        context_manager.set_last_analysis(analysis)
        reply = analysis.run_formula_pbe_query()
        cell_candidate = apply_reply(analysis, reply)

        reply = {
            "status": "ok",
            "range": analysis.outputSection.range,
            "candidate": cell_candidate,
        }

        print(reply)
        return reply
    except Exception as e:
        print(f"Exception in handle_formula_pbe: {e}")
        return {
            "status": "error",
            "message": f"An error occurred while processing formula by example: {str(e)}"
        }

def handle_create_visual(msg):
    context = llm.getContext()
    analysis = Analysis(msg)
    context_manager.set_last_context(context)
    context_manager.set_last_analysis(analysis)
    reply = analysis.run_create_visual_query()
    title, type = apply_create_visual(reply)
    reply = {
        "type": "create_visual",
        "status": "ok",
        "range": analysis.inputSection.range,
        "title": title,
        "chart_type": type
    }
    print(reply)
    return reply

def handle_formula_chk(msg):
    context = llm.getContext()
    analysis = Analysis(msg)
    context_manager.set_last_context(context)
    context_manager.set_last_analysis(analysis)
    reply = analysis.run_formula_chk_query()
    warns, passes = apply_formula_chk(reply)
    infos = []
    for warn in warns:
        infos.append({"intent": "warning", "info": warn})
    for pass_ in passes:
        infos.append({"intent": "success", "info": pass_})
    if not infos:
        infos.append({"intent": "info", "info": reply})
    print(reply)
    infos.append({"intent": "info", "info": "No issues found."})
    reply = {
        "type": "formula_chk",
        "status": "ok",
        "info": infos,
    }
    print(reply)
    return reply


if __name__ == "__main__":
    msg = {
        "inputRange": "Sheet1!A1:B2",
        "inputData": [[1, 2], [3, 4]],
        "outputRange": "Sheet1!C1:C2",
        "outputData": [[3], [7]],
        "description": "fill column C with the sum of columns A and B",
    }
    handle_autofill(msg)