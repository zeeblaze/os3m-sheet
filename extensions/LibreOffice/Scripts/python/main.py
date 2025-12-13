import re
import urllib.request
import json
import traceback
import uno
import unohelper
from com.sun.star.awt import Rectangle
from com.sun.star.awt import XActionListener, XItemListener
from com.sun.star.awt.MessageBoxType import INFOBOX
from com.sun.star.task import XJobExecutor
from com.sun.star.lang import XServiceInfo

class ActionComboListener(unohelper.Base, XItemListener):
    def __init__(self, controller):
        self.controller = controller

    def itemStateChanged(self, event):
        self.controller.toggle_visibility(event.Source.Text)

    def disposing(self, event):
        pass

class CloseButtonListener(unohelper.Base, XActionListener):
    def __init__(self, dialog):
        self.dialog = dialog

    def actionPerformed(self, event):
        self.dialog.endExecute()

    def disposing(self, event):
        pass

class ExecuteButtonListener(unohelper.Base, XActionListener):
    def __init__(self, controller):
        self.controller = controller

    def actionPerformed(self, event):
        self.controller.execute_operation()

    def disposing(self, event):
        pass

class Os3mSheetController:
    def __init__(self, ctx):
        self.ctx = ctx
        self.smgr = ctx.ServiceManager
        self.desktop = self.smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        self.model = self.desktop.getCurrentComponent()
        self.sheet = None
        self.window = None
        
        if self.model and hasattr(self.model, "CurrentController") and self.model.CurrentController:
            if hasattr(self.model.CurrentController, "ActiveSheet"):
                self.sheet = self.model.CurrentController.ActiveSheet
            if hasattr(self.model.CurrentController, "Frame") and self.model.CurrentController.Frame:
                self.window = self.model.CurrentController.Frame.ContainerWindow
        self.dialog = None

    def run(self):
        if not self.sheet:
            self.show_message("Error", "No active spreadsheet found.")
            return

        # Get selected cell values and depict selection
        selection = self.model.CurrentController.Selection
        cell_addresses = []
        try:
            for row in range(selection.RangeAddress.StartRow, selection.RangeAddress.EndRow + 1):
                for col in range(selection.RangeAddress.StartColumn, selection.RangeAddress.EndColumn + 1):
                    # Convert col/row to address
                    col_addr = ""
                    col_num = col
                    while col_num >= 0:
                        col_addr = chr(ord('A') + (col_num % 26)) + col_addr
                        col_num = col_num // 26 - 1
                    cell_addresses.append(f"{col_addr}{row+1}")
            # Check if selection is sequential
            start_col = selection.RangeAddress.StartColumn
            end_col = selection.RangeAddress.EndColumn
            start_row = selection.RangeAddress.StartRow
            end_row = selection.RangeAddress.EndRow
            if (end_col > start_col or end_row > start_row):
                # Range depiction
                col_addr_start = ""
                col_num = start_col
                while col_num >= 0:
                    col_addr_start = chr(ord('A') + (col_num % 26)) + col_addr_start
                    col_num = col_num // 26 - 1
                col_addr_end = ""
                col_num = end_col
                while col_num >= 0:
                    col_addr_end = chr(ord('A') + (col_num % 26)) + col_addr_end
                    col_num = col_num // 26 - 1
                values_str = f"{self.sheet.Name}!{col_addr_start}{start_row+1}:{col_addr_end}{end_row+1}"
            else:
                values_str = ", ".join(cell_addresses)
        except Exception:
            values_str = "No selection or not a spreadsheet."

        # Create dialog model
        dialog_model = self.smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", self.ctx)
        dialog_model.PositionX = 100
        dialog_model.PositionY = 100
        dialog_model.Width = 340
        dialog_model.Height = 290
        dialog_model.Title = "OS3M-Sheet Spreadsheet Operations"

        # Input/Output Ranges and Description (common for many ops)
        y_pos = 10
        label_width = 85
        edit_x = 10 + label_width + 5
        edit_width = dialog_model.Width - edit_x - 10
        row_height = 22
        
        # Input Range Label and Text
        input_range_label_model = dialog_model.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        input_range_label_model.PositionX = 10
        input_range_label_model.PositionY = y_pos + 4 # Align text vertically
        input_range_label_model.Width = label_width
        input_range_label_model.Height = 15
        input_range_label_model.Name = "InputRangeLabel"
        input_range_label_model.Label = "Input Range:"
        dialog_model.insertByName("InputRangeLabel", input_range_label_model)
        input_range_edit_model = dialog_model.createInstance("com.sun.star.awt.UnoControlEditModel")
        input_range_edit_model.PositionX = edit_x
        input_range_edit_model.PositionY = y_pos
        input_range_edit_model.Width = edit_width
        input_range_edit_model.Height = 20
        input_range_edit_model.Name = "InputRangeEdit"
        input_range_edit_model.Text = values_str # Pre-fill with current selection
        dialog_model.insertByName("InputRangeEdit", input_range_edit_model)
        y_pos += row_height
        
        # Output Range Label and Text
        output_range_label_model = dialog_model.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        output_range_label_model.PositionX = 10
        output_range_label_model.PositionY = y_pos + 4
        output_range_label_model.Width = label_width
        output_range_label_model.Height = 15
        output_range_label_model.Name = "OutputRangeLabel"
        output_range_label_model.Label = "Output Range:"
        dialog_model.insertByName("OutputRangeLabel", output_range_label_model)
        output_range_edit_model = dialog_model.createInstance("com.sun.star.awt.UnoControlEditModel")
        output_range_edit_model.PositionX = edit_x
        output_range_edit_model.PositionY = y_pos
        output_range_edit_model.Width = edit_width
        output_range_edit_model.Height = 20
        output_range_edit_model.Name = "OutputRangeEdit"
        output_range_edit_model.Text = "" # User fills this
        dialog_model.insertByName("OutputRangeEdit", output_range_edit_model)
        y_pos += row_height
        
        # Description Label and Text
        description_label_model = dialog_model.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        description_label_model.PositionX = 10
        description_label_model.PositionY = y_pos + 4
        description_label_model.Width = label_width
        description_label_model.Height = 15
        description_label_model.Name = "DescriptionLabel"
        description_label_model.Label = "Description:"
        dialog_model.insertByName("DescriptionLabel", description_label_model)
        description_edit_model = dialog_model.createInstance("com.sun.star.awt.UnoControlEditModel")
        description_edit_model.PositionX = edit_x
        description_edit_model.PositionY = y_pos
        description_edit_model.Width = edit_width
        description_edit_model.Height = 20
        description_edit_model.Name = "DescriptionEdit"
        description_edit_model.Text = "" # User fills this
        dialog_model.insertByName("DescriptionEdit", description_edit_model)
        y_pos += row_height
        
        # Feedback Message Label and Text (for feedback only)
        feedback_label_model = dialog_model.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        feedback_label_model.PositionX = 10
        feedback_label_model.PositionY = y_pos + 4
        feedback_label_model.Width = label_width
        feedback_label_model.Height = 15
        feedback_label_model.Name = "FeedbackLabel"
        feedback_label_model.Label = "Feedback:"
        dialog_model.insertByName("FeedbackLabel", feedback_label_model)
        feedback_edit_model = dialog_model.createInstance("com.sun.star.awt.UnoControlEditModel")
        feedback_edit_model.PositionX = edit_x
        feedback_edit_model.PositionY = y_pos
        feedback_edit_model.Width = edit_width
        feedback_edit_model.Height = 20
        feedback_edit_model.Name = "FeedbackEdit"
        feedback_edit_model.Text = "" # User fills this
        dialog_model.insertByName("FeedbackEdit", feedback_edit_model)
        y_pos += row_height # Extra space before action
        
        # Action selection via ComboBox
        action_label_model = dialog_model.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        action_label_model.PositionX = 10
        action_label_model.PositionY = y_pos + 4
        action_label_model.Width = label_width
        action_label_model.Height = 15
        action_label_model.Name = "ActionLabel"
        action_label_model.Label = "Action:"
        dialog_model.insertByName("ActionLabel", action_label_model)

        action_combo_model = dialog_model.createInstance("com.sun.star.awt.UnoControlComboBoxModel")
        action_combo_model.PositionX = edit_x
        action_combo_model.PositionY = y_pos
        action_combo_model.Width = edit_width
        action_combo_model.Height = 20
        action_combo_model.Name = "ActionComboBox"
        action_combo_model.Dropdown = True
        action_combo_model.StringItemList = (
            "autofill", "feedback", "rangesel", "summary", "formula_exp", 
            "batchproc", "formula_pbe", "create_visual", "formula_chk"
        )
        action_combo_model.Text = "autofill" # Default selection
        dialog_model.insertByName("ActionComboBox", action_combo_model)
        y_pos += row_height

        # Action Description Label
        action_desc_label_model = dialog_model.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        action_desc_label_model.PositionX = 10
        action_desc_label_model.PositionY = y_pos
        action_desc_label_model.Width = dialog_model.Width - 20
        action_desc_label_model.Height = 30 # Allow for two lines
        action_desc_label_model.Name = "ActionDescriptionLabel"
        action_desc_label_model.Label = "" # Will be set by listener
        dialog_model.insertByName("ActionDescriptionLabel", action_desc_label_model)
        y_pos += 30 # Space after description, matches height

        # History Text Area
        history_label_model = dialog_model.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        history_label_model.PositionX = 10
        history_label_model.PositionY = y_pos
        history_label_model.Width = 100
        history_label_model.Height = 15
        history_label_model.Name = "HistoryLabel"
        history_label_model.Label = "Conversation History:"
        dialog_model.insertByName("HistoryLabel", history_label_model)
        y_pos += 18

        history_text_area_model = dialog_model.createInstance("com.sun.star.awt.UnoControlEditModel")
        history_text_area_model.PositionX = 10
        history_text_area_model.PositionY = y_pos
        history_text_area_model.Width = dialog_model.Width - 20
        history_text_area_model.Height = 60
        history_text_area_model.Name = "HistoryTextArea"
        history_text_area_model.Text = "" # Will be populated later
        history_text_area_model.ReadOnly = True
        history_text_area_model.MultiLine = True
        history_text_area_model.VScroll = True
        dialog_model.insertByName("HistoryTextArea", history_text_area_model)
        y_pos += history_text_area_model.Height + 10 # Space before buttons
        
        # --- Buttons ---
        button_width = 80
        button_spacing = 10
        total_buttons_width = (button_width * 2) + button_spacing # For two buttons
        start_x = (dialog_model.Width - total_buttons_width) / 2
        
        # Execute button
        execute_button_model = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
        execute_button_model.PositionX = start_x
        execute_button_model.PositionY = y_pos # Placeholder Y
        execute_button_model.Width = button_width
        execute_button_model.Height = 25
        execute_button_model.Name = "ExecuteButton"
        execute_button_model.Label = "Execute"
        execute_button_model.DefaultButton = True
        dialog_model.insertByName("ExecuteButton", execute_button_model)

        # Close button at the bottom
        close_button_model = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
        close_button_model.PositionX = start_x + button_width + button_spacing
        close_button_model.PositionY = y_pos
        close_button_model.Width = button_width
        close_button_model.Height = 25
        close_button_model.Name = "CloseButton"
        close_button_model.Label = "Close"
        dialog_model.insertByName("CloseButton", close_button_model)

        # Create the dialog control
        self.dialog = self.smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", self.ctx)
        self.dialog.setModel(dialog_model)
        toolkit = self.smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)
        self.dialog.createPeer(toolkit, None)

        # Attach listener to the ComboBox
        action_combo_control = self.dialog.getControl("ActionComboBox")
        action_listener = ActionComboListener(self)
        action_combo_control.addItemListener(action_listener)

        # Set initial visibility
        self.toggle_visibility(action_combo_control.Text)
        self.update_history_display(force_refresh=True)

        # Add event listener to Close button
        close_button = self.dialog.getControl("CloseButton")
        close_listener = CloseButtonListener(self.dialog)
        close_button.addActionListener(close_listener)

        # Attach listener to Execute button
        execute_button = self.dialog.getControl("ExecuteButton")
        execute_button.addActionListener(ExecuteButtonListener(self))

        # Show dialog
        self.dialog.execute()
        self.dialog.dispose()

    def toggle_visibility(self, action):
        action_descriptions = {
            "autofill": "Fills a range based on examples in the input and output ranges.",
            "feedback": "Provide feedback on the last operation to improve future results.",
            "rangesel": "Selects a data range based on your text description.",
            "summary": "Generates a text summary of the data in the input range.",
            "formula_exp": "Explains the formula(s) found in the input range.",
            "batchproc": "Performs a batch operation on the input range based on a description.",
            "formula_pbe": "Generates a formula by example, using input and output ranges.",
            "create_visual": "Creates a chart from the data in the input range.",
            "formula_chk": "Checks the formula in the input range for correctness and provides info."
        }
        # Set the description label text and control visibility
        desc_label = self.dialog.getControl("ActionDescriptionLabel")
        desc_label.setText(action_descriptions.get(action, "Select an action to see its description."))

        is_feedback = (action == "feedback")
        needs_input = action in ["autofill", "rangesel", "summary", "formula_exp", "batchproc", "formula_pbe", "create_visual", "formula_chk"]
        needs_output = action in ["autofill", "formula_pbe"]
        needs_description = action in ["autofill", "rangesel", "summary", "formula_exp", "batchproc", "formula_pbe", "create_visual", "formula_chk"]

        # Toggle visibility based on the needs of the selected action
        self.dialog.getControl("InputRangeLabel").setVisible(needs_input)
        self.dialog.getControl("InputRangeEdit").setVisible(needs_input)

        self.dialog.getControl("OutputRangeLabel").setVisible(needs_output)
        self.dialog.getControl("OutputRangeEdit").setVisible(needs_output)

        self.dialog.getControl("DescriptionLabel").setVisible(needs_description)
        self.dialog.getControl("DescriptionEdit").setVisible(needs_description)

        # Controls for feedback operation
        self.dialog.getControl("FeedbackLabel").setVisible(is_feedback)
        self.dialog.getControl("FeedbackEdit").setVisible(is_feedback)
        
        # Start with the Y position of the history text area and add its height
        button_y_pos = self.dialog.getControl("HistoryTextArea").getPosSize().Y + self.dialog.getControl("HistoryTextArea").getPosSize().Height + 15
        
        self.dialog.getControl("ExecuteButton").setPosSize(self.dialog.getControl("ExecuteButton").getPosSize().X, button_y_pos, 0, 0, 1) # 1 = PosSize.Y
        self.dialog.getControl("CloseButton").setPosSize(self.dialog.getControl("CloseButton").getPosSize().X, button_y_pos, 0, 0, 1) # 1 = PosSize.Y
        
        # The new height is the button's Y position plus its height and a margin
        final_height = button_y_pos + self.dialog.getControl("ExecuteButton").getPosSize().Height + 20
        
        self.dialog.setPosSize(self.dialog.getPosSize().X, self.dialog.getPosSize().Y, self.dialog.getPosSize().Width, final_height, 15) # 15 = SizeFlags.HEIGHT

    def update_history_display(self, force_refresh=False):
        try:
            url = "http://127.0.0.1:8000/history"
            req = urllib.request.Request(url, headers={"Accept": "application/json"})
            with urllib.request.urlopen(req) as response:
                response_data = json.loads(response.read().decode("utf-8"))
            
            history = response_data.get("history", [])
            history_text = ""
            if not history:
                history_text = "No conversation history yet."
            else:
                for i, (request_str, response_obj) in enumerate(history):
                    response_str = json.dumps(response_obj, indent=2)
                    history_text += f"--- Interaction {i+1} ---\nUSER:\n{request_str}\n\nLLM:\n{response_str}\n\n"
            
            self.dialog.getControl("HistoryTextArea").setText(history_text)
        except Exception as e:
            error_text = f"Failed to fetch conversation history: {e}"
            self.dialog.getControl("HistoryTextArea").setText(error_text)
            print(error_text)

    def show_output_dialog(self, result):
        dialog_model = self.smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", self.ctx)
        dialog_model.PositionX = 100
        dialog_model.PositionY = 100
        dialog_model.Width = 300
        dialog_model.Height = 210
        dialog_model.Title = "Os3m-Server Output"

        text_area_model = dialog_model.createInstance("com.sun.star.awt.UnoControlEditModel")
        text_area_model.PositionX = 10
        text_area_model.PositionY = 10
        text_area_model.Width = 280 # Dialog width - 2*margin
        text_area_model.Height = 150
        text_area_model.Name = "OutputTextArea"
        text_area_model.Text = str(result)
        text_area_model.ReadOnly = True
        text_area_model.MultiLine = True
        text_area_model.VScroll = True # Enable vertical scrollbar
        dialog_model.insertByName("OutputTextArea", text_area_model)

        close_button_model = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
        close_button_model.PositionX = 110 # (300 - 80) / 2
        close_button_model.PositionY = 170 # 10 (margin) + 150 (text area) + 10 (spacing)
        close_button_model.Width = 80
        close_button_model.Height = 25
        close_button_model.Name = "CloseButton"
        close_button_model.Label = "Close"
        dialog_model.insertByName("CloseButton", close_button_model)

        dialog = self.smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", self.ctx)
        dialog.setModel(dialog_model)
        toolkit = self.smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)
        dialog.createPeer(toolkit, None)

        close_button = dialog.getControl("CloseButton")
        close_listener = CloseButtonListener(dialog)
        close_button.addActionListener(close_listener)

        dialog.execute()
        dialog.dispose()

    def create_chart(self, sheet, chart_range_address_str, title, chart_type_str):
        try:
            # Define the chart position and size
            charts = sheet.getCharts()
            count = len(charts.getElementNames())
            chart_name = f"Chart_{count + 1}"

            chart_rect = Rectangle()
            chart_rect.X = 10000
            chart_rect.Y = 5000 + (count * 1000)
            chart_rect.Width = 15000
            chart_rect.Height = 10000

            # Parse range string to handle sheet names
            if "!" in chart_range_address_str:
                sheet_name, range_addr = chart_range_address_str.rsplit("!", 1)
                data_sheet = self.model.getSheets().getByName(sheet_name)
            else:
                data_sheet = sheet
                range_addr = chart_range_address_str

            # Get the CellRangeAddress object from the string
            chart_range = data_sheet.getCellRangeByName(range_addr)
            chart_range_address = chart_range.getRangeAddress()

            # Determine if first column should be labels. If only 1 column, it is data.
            has_row_headers = (chart_range_address.EndColumn > chart_range_address.StartColumn)
            # Add a new chart to the sheet
            charts.addNewByName(chart_name, chart_rect, (chart_range_address,), True, has_row_headers)
            
            chart_doc = charts.getByName(chart_name).getEmbeddedObject()
            chart_doc.HasMainTitle = True
            chart_doc.Title.String = title

            # Set the chart type
            if chart_type_str == "Bar":
                diagram = chart_doc.createInstance("com.sun.star.chart.BarDiagram")
                diagram.Vertical = True
            elif chart_type_str == "Column":
                diagram = chart_doc.createInstance("com.sun.star.chart.BarDiagram")
                diagram.Vertical = False
            else:
                diagram = chart_doc.createInstance(f"com.sun.star.chart.{chart_type_str}Diagram")

            if not diagram:
                self.show_message("Chart Error", f"Failed to create a {chart_type_str} diagram. The chart type may be invalid or unsupported.")
                return
            
            chart_doc.setDiagram(diagram)

            # Set data source
            data_row_source = uno.Enum("com.sun.star.chart.ChartDataRowSource", "COLUMNS")
            diagram.DataRowSource = data_row_source

            self.show_message("Chart Creation", f"Successfully created a {chart_type_str} chart: {title}")

        except Exception as e:
            self.show_message("Chart Error", f"Failed to create chart: {e}\nType: {type(e)}\nArgs: {e.args}")

    def call_api(self, endpoint, request_data):
        url = f"http://127.0.0.1:8000/{endpoint}"
        try:
            data = json.dumps(request_data).encode("utf-8")
            req = urllib.request.Request(
                url,
                data=data,
                headers={"Content-Type": "application/json"}
            )
            with urllib.request.urlopen(req) as response:
                response_data = json.loads(response.read().decode("utf-8"))
            return response_data
        except Exception as e:
            self.show_message("API Error", f"Could not call API endpoint {endpoint}: {e}")
            return None

    def show_message(self, title, message):
        toolkit = self.smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)
        msgbox = toolkit.createMessageBox(self.window, INFOBOX, 1, title, message)
        msgbox.execute()

    def execute_operation(self):
        op_type = self.dialog.getControl("ActionComboBox").Text.strip()

        input_range = self.dialog.getControl("InputRangeEdit").Text.strip()
        output_range = self.dialog.getControl("OutputRangeEdit").Text.strip()
        description = self.dialog.getControl("DescriptionEdit").Text.strip()
        feedback_msg = self.dialog.getControl("FeedbackEdit").Text.strip() # Only for feedback

        # Helper to get data from a range
        def get_data_from_range(range_str):
            if not range_str:
                return []
            try:
                # Basic parsing to get sheet and cell range
                parts = range_str.rsplit('!', 1)
                if len(parts) > 1:
                    sheet_name = parts[0]
                    cell_range_str = parts[1]
                    current_sheet = self.model.getSheets().getByName(sheet_name)
                else:
                    current_sheet = self.sheet # Use the current active sheet
                    cell_range_str = parts[0]

                cell_range = current_sheet.getCellRangeByName(cell_range_str)
                range_address = cell_range.getRangeAddress()
                
                data = []
                for r in range(range_address.StartRow, range_address.EndRow + 1):
                    row_data = []
                    for c in range(range_address.StartColumn, range_address.EndColumn + 1):
                        cell = current_sheet.getCellByPosition(c, r)
                        if op_type in ["formula_exp", "formula_chk", "formula_pbe"]:
                            row_data.append(cell.getFormula())
                        else:
                            row_data.append(cell.getString())
                    data.append(row_data)
                return data
            except Exception as e:
                self.show_message("Error", f"Could not get data from range {range_str}: {e}")
                return []

        request_data = {}
        if op_type == "feedback":
            request_data = {"feedbackMsg": feedback_msg}
        # All other operations need input range, data, and description
        else:
            input_data = get_data_from_range(input_range)
            # For batchproc, the output range is the same as the input range
            request_data = {
                "inputRange": input_range,
                "inputData": input_data,
                "description": description
            }
            if output_range: # outputRange and outputData are optional for Analysis constructor
                output_data = get_data_from_range(output_range)
                request_data["outputRange"] = output_range
                request_data["outputData"] = output_data

        result = self.call_api(op_type, request_data)

        # Update the history display on the dialog
        if result:
            self.update_history_display(force_refresh=True) # Update history for all other ops
            if op_type in ["autofill", "formula_pbe", "feedback", "batchproc"]:
                # Assume result has {"candidate": data_to_write, "range": target_range}
                candidate_data = result.get("result", {}).get("candidate") # Assuming result from api.py is {"message": "...", "result": actual_result}
                target_range_str = result.get("result", {}).get("range")
                
                # Prefer user input for range to preserve sheet name
                range_to_use = target_range_str
                if op_type in ["autofill", "formula_pbe"] and output_range:
                    range_to_use = output_range
                elif op_type == "batchproc" and input_range:
                    range_to_use = input_range

                if candidate_data and range_to_use:
                    # Write data back to sheet
                    self.show_message("Autofill Result", f"Writing to {range_to_use}: {candidate_data}")
                    try:
                        parts = range_to_use.rsplit('!', 1)
                        if len(parts) > 1:
                            sheet_name = parts[0]
                            cell_range_str = parts[1]
                        else:
                            sheet_name = self.sheet.Name
                            cell_range_str = parts[0]

                        target_sheet = self.model.getSheets().getByName(sheet_name)
                        target_range_address = target_sheet.getCellRangeByName(cell_range_str).getRangeAddress()

                        row_offset = target_range_address.StartRow
                        col_offset = target_range_address.StartColumn

                        for r_idx, row_list in enumerate(candidate_data):
                            for c_idx, cell_val in enumerate(row_list):
                                cell = target_sheet.getCellByPosition(col_offset + c_idx, row_offset + r_idx)
                                # Heuristic to check if it's a formula: starts with '='
                                if isinstance(cell_val, str) and cell_val.startswith('='):
                                    cell.setFormula(cell_val)
                                else:
                                    cell.setString(str(cell_val))
                        self.show_message("Operation Success", f"{op_type} data written to {range_to_use}.")
                    except Exception as e:
                        self.show_message("Write Error", f"Failed to write autofill/formula_pbe data: {e}")
                else:
                    self.show_message("Error", f"Invalid result for {op_type}: {result}")
            elif op_type == "summary" or op_type == "formula_exp":
                summary_text = result.get("result", {}).get("reply", "No reply received.")
                self.show_output_dialog(summary_text)
            elif op_type == "rangesel":
                colors = result.get("result", {}).get("colors", [])
                colors = [c for c in colors if c]
                target_range_str = result.get("result", {}).get("range")
                if not colors or not target_range_str:
                    self.show_message("Range Select Error", "Did not receive color data from the API.")
                    return
                try:
                    parts = target_range_str.rsplit('!', 1)
                    if len(parts) > 1:
                        sheet_name = parts[0]
                        cell_range_str = parts[1]
                    else:
                        sheet_name = self.sheet.Name
                        cell_range_str = parts[0]
                    target_sheet = self.model.getSheets().getByName(sheet_name)
                    target_range_address = target_sheet.getCellRangeByName(cell_range_str).getRangeAddress()

                    color_map = {'green': 0x00FF00, 'yellow': 0xFFFF00, 'red': 0xFF0000, 'white': 0xFFFFFF}
                    color_counts = {'green': 0, 'yellow': 0, 'red': 0, 'white': 0}
                    color_idx = 0

                    for r in range(target_range_address.StartRow, target_range_address.EndRow + 1):
                        for c in range(target_range_address.StartColumn, target_range_address.EndColumn + 1):
                            if color_idx < len(colors):
                                color_name = str(colors[color_idx]).lower().strip()
                                cell = target_sheet.getCellByPosition(c, r)
                                cell.CellBackColor = color_map.get(color_name, 0xFFFFFF)
                                color_counts[color_name] = color_counts.get(color_name, 0) + 1
                                color_idx += 1
                    self.show_message("Range Select Complete", f"Highlighting complete. Counts: {color_counts}")
                except Exception as e:
                    self.show_message("Range Select Error", f"Failed to apply highlighting: {e}")
            elif op_type == "create_visual":
                title = result.get("result", {}).get("title", "No Title")
                chart_type = result.get("result", {}).get("chart_type", "Unknown")
                target_range_str = result.get("result", {}).get("range") # This is inputSection range
                
                # Use input_range if available to preserve sheet info
                range_to_use = input_range if input_range else target_range_str

                if range_to_use:
                    self.create_chart(self.sheet, range_to_use, title, chart_type)
                else:
                    self.show_message("Chart Error", "No data range provided for chart creation.")
            elif op_type == "formula_chk":
                infos = result.get("result", {}).get("info", [])
                info_text = "\n".join([f"[{i['intent'].upper()}] {i['info']}" for i in infos])
                self.show_output_dialog(f"Formula Check Results:\n{info_text}")
            elif op_type == "feedback":
                self.show_message("Feedback", result.get("message", "Feedback sent."))
            else:
                self.show_message("API Response", str(result))
        else:
            self.show_message("API Call Failed", f"No response from {op_type} API.")

class Os3mSheetJob(unohelper.Base, XJobExecutor, XServiceInfo):
    def __init__(self, ctx):
        self.ctx = ctx

    def trigger(self, args):
        try:
            controller = Os3mSheetController(self.ctx)
            controller.run()
        except Exception:
            error_message = traceback.format_exc()
            try:
                smgr = self.ctx.ServiceManager
                toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)
                desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
                frame = desktop.getCurrentFrame()
                window = frame.getContainerWindow() if frame else None
                msgbox = toolkit.createMessageBox(window, INFOBOX, 1, "Os3mSheet Error", f"Failed to launch:\n{error_message}")
                msgbox.execute()
            except:
                print(f"Os3mSheet Error: {error_message}")

    def getImplementationName(self):
        return "org.os3msheet.Os3mSheetJob"

    def getSupportedServiceNames(self):
        return ("com.sun.star.task.Job", "org.os3msheet.Os3mSheetJob")

    def supportsService(self, serviceName):
        return serviceName in self.getSupportedServiceNames()

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(Os3mSheetJob, "org.os3msheet.Os3mSheetJob", ("com.sun.star.task.Job", "org.os3msheet.Os3mSheetJob"))