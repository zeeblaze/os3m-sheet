import re
import json
import dspy

cellPattern = re.compile(r'([A-Za-z]+)(\d+)')
formulaList = ["SUM", "AVERAGE", "COUNT", "SUBTOTAL", "MODULUS", "POWER", "CEILING", "FLOOR", "CONCATENATE", "LEN",
               "REPLACE", "SUBSTITUTE", "LEFT", "RIGHT", "MID", "UPPER", "LOWER", "PROPER", "TIME", "VLOOKUP",
               "COUNTIF", "SUMIF"]


def column_to_num(col: str):
    num = 0
    for char in col:
        num = num * 26 + (ord(char.upper()) - ord('A') + 1)
    return num


class Cell:
    def __init__(self, input: str):
        match = cellPattern.match(input)
        if not match:
            raise ValueError(f"Invalid cell reference: {input}")
        self.col, self.row = match.group(1), int(match.group(2))

    def __sub__(self, other):
        assert (isinstance(other, Cell))
        return column_to_num(self.col) - column_to_num(other.col) + 1, self.row - other.row + 1

    def get_index_str(self):
        return f"{self.col}{self.row}"


class Section:
    def __init__(self, sheet: str, cellL: Cell, cellR: Cell, data: list):
        self.sheet, self.cellL, self.cellR, self.data = sheet, cellL, cellR, data
        self.width, self.height = self.cellR - self.cellL
        self.range = f"{self.sheet}!{self.cellL.get_index_str()}:{self.cellR.get_index_str()}"
        pass


def getSection(input: str, data: list):
    if "!" in input:
        sheet, r = input.rsplit("!", 1)
    else:
        sheet = "Sheet1"
        r = input
    cells = [Cell(x) for x in r.split(":")]
    if len(cells) == 1:
        cellL = cellR = cells[0]
    else:
        cellL, cellR = cells
    return Section(sheet, cellL, cellR, data)


class GenerateFormulas(dspy.Signature):
    """Using the input and output data in row-major order, generate the necessary formulas to achieve the desired output.
    Think step by step.
    If the first row of the input data contains headers, generate a corresponding header for the output data instead of a formula.
    If feedback is provided, use it to refine the result.
    The output formulas should be a JSON 2D array (list of lists) containing the result formulas or values."""
    input_data = dspy.InputField(desc="Input data")
    input_range = dspy.InputField()
    output_range = dspy.InputField()
    goal = dspy.InputField()
    feedback = dspy.InputField(desc="User feedback")
    formulas = dspy.OutputField(desc="JSON 2D array")

class SummarizeData(dspy.Signature):
    """Summarize based on the description. Output JSON: {"summary": "..."}"""
    data = dspy.InputField()
    goal = dspy.InputField()
    summary = dspy.OutputField(desc="JSON object")

class ExplainFormulas(dspy.Signature):
    """Explain formulas concisely based on the description. Output JSON: {"explanation": "..."}"""
    formulas = dspy.InputField()
    goal = dspy.InputField()
    explanation = dspy.OutputField(desc="JSON object")

class GenerateFormulasPBE(dspy.Signature):
    """Think step by step to generate formulas based on the provided input and output data examples in row-major order.
    The output formulas should be a JSON 2D array (list of lists) containing the inferred formulas."""
    input_data = dspy.InputField()
    output_example = dspy.InputField()
    output_range = dspy.InputField()
    goal = dspy.InputField()
    formulas = dspy.OutputField(desc="JSON 2D array")

class SelectCells(dspy.Signature):
    """Think step by step to select cells based on the provided input data.
    The output colors should be a JSON 2D array of strings ('green' or 'white') matching the input data shape."""
    data = dspy.InputField()
    goal = dspy.InputField()
    colors = dspy.OutputField(desc="JSON 2D array of colors")

class TransformData(dspy.Signature):
    """Think step by step to transform data based on the provided input data in row-major order.
    The output transformed_data should be a JSON 2D array."""
    data = dspy.InputField()
    goal = dspy.InputField()
    transformed_data = dspy.OutputField(desc="JSON 2D array")

class CheckCompatibility(dspy.Signature):
    """Think step by step to check for formula compatibility issues (Excel/LibreOffice) based on the provided input data.
    The output issues should be a JSON object with keys "issues" and "passed"."""
    formulas = dspy.InputField()
    issues = dspy.OutputField(desc="JSON object")

class CreateChart(dspy.Signature):
    """Think step by step to create a chart based on the provided input data. Select the most suitable chart type from Line, Pie, Bar, Area, Column."""
    data = dspy.InputField()
    goal = dspy.InputField()
    chart_config = dspy.OutputField(desc="JSON object with keys 'title' and 'type'")

class Analysis:
    def __init__(self, msg: dict):
        self.inputSection = getSection(msg['inputRange'], msg['inputData'])
        self.outputSection = None
        if "outputRange" in msg.keys() and msg['outputRange']:
            try:
                self.outputSection = getSection(msg['outputRange'], msg['outputData'])
            except ValueError:
                pass
        self.desc = msg['description']
        self.feedback = msg.get('feedbackMsg', "")
        pass

    def run_query(self):
        goal = self.desc if self.desc else "Autofill the remaining cells based on the pattern"
        try:
            pred = dspy.ChainOfThought(GenerateFormulas)(
                input_data=str(self.inputSection.data),
                input_range=self.inputSection.range,
                output_range=self.outputSection.range,
                goal=goal,
                feedback=self.feedback
            )
            print(pred)
            return pred.formulas
        except Exception as e:
            print(f"Error in run_query with ChainOfThought: {e}")
            try:
                pred = dspy.Predict(GenerateFormulas)(
                    input_data=str(self.inputSection.data),
                    input_range=self.inputSection.range,
                    output_range=self.outputSection.range,
                    goal=goal,
                    feedback=self.feedback
                )
                print(pred)
                return pred.formulas
            except Exception as e2:
                print(f"Error in run_query with Predict: {e2}")
                return "[]"

    def run_summary_query(self):
        pred = dspy.Predict(SummarizeData)(data=str(self.inputSection.data), goal=self.desc)
        print(pred)
        return pred.summary

    def run_exp_explain_query(self):
        pred = dspy.Predict(ExplainFormulas)(formulas=str(self.inputSection.data), goal=self.desc)
        print(pred)
        return pred.explanation

    def run_formula_pbe_query(self):
        goal = self.desc if self.desc else "Infer the pattern from the examples"
        try:
            pred = dspy.ChainOfThought(GenerateFormulasPBE)(
                input_data=str(self.inputSection.data),
                output_example=str(self.outputSection.data),
                output_range=self.outputSection.range,
                goal=goal
            )
            print(pred)
            return pred.formulas
        except Exception as e:
            print(f"Error in run_formula_pbe_query with ChainOfThought: {e}")
            try:
                pred = dspy.Predict(GenerateFormulasPBE)(
                    input_data=str(self.inputSection.data),
                    output_example=str(self.outputSection.data),
                    output_range=self.outputSection.range,
                    goal=goal
                )
                print(pred)
                return pred.formulas
            except Exception as e2:
                print(f"Error in run_formula_pbe_query with Predict: {e2}")
                return "[]"

    def run_range_sel_query(self):
        pred = dspy.ChainOfThought(SelectCells)(data=str(self.inputSection.data), goal=self.desc)
        print(pred)
        colors = pred.colors
        if isinstance(colors, str):
            try:
                parsed = json.loads(colors.replace("'", '"'))
                return json.dumps(parsed)
            except Exception:
                pass
        return colors

    def run_batchproc_query(self):
        try:
            pred = dspy.ChainOfThought(TransformData)(data=str(self.inputSection.data), goal=self.desc)
            print(pred)
            return pred.transformed_data
        except Exception as e:
            print(f"Error in run_batchproc_query with ChainOfThought: {e}")
            try:
                pred = dspy.Predict(TransformData)(data=str(self.inputSection.data), goal=self.desc)
                print(pred)
                return pred.transformed_data
            except Exception as e2:
                print(f"Error in run_batchproc_query with Predict: {e2}")
                return "[]"

    def run_formula_chk_query(self):
        pred = dspy.ChainOfThought(CheckCompatibility)(formulas=str(self.inputSection.data))
        print(pred)
        return pred.issues

    def run_create_visual_query(self):
        def capitalize_type(config):
            try:
                if isinstance(config, str):
                    try:
                        data = json.loads(config)
                    except json.JSONDecodeError:
                        data = json.loads(config.replace("'", '"'))
                else:
                    data = config

                if isinstance(data, dict) and 'type' in data:
                    data['type'] = data['type'].capitalize()

                if isinstance(config, str):
                    return json.dumps(data)
                return data
            except:
                return config

        try:
            pred = dspy.ChainOfThought(CreateChart)(data=str(self.inputSection.data), goal=self.desc)
            print(pred)
            return capitalize_type(pred.chart_config)
        except Exception as e:
            print(f"Error in run_create_visual_query with ChainOfThought: {e}")
            try:
                pred = dspy.Predict(CreateChart)(data=str(self.inputSection.data), goal=self.desc)
                print(pred)
                return capitalize_type(pred.chart_config)
            except Exception as e2:
                print(f"Error in run_create_visual_query with Predict: {e2}")
                return "{}"