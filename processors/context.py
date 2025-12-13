class ContextManager:
    def __init__(self):
        self.context = None
        self.analysis = None

    def set_last_context(self, context):
        self.context = context

    def set_last_analysis(self, analysis):
        self.analysis = analysis

    def get_last_context(self):
        return self.context

    def get_last_analysis(self):
        return self.analysis