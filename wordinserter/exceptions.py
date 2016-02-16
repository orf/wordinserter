
class InsertError(RuntimeError):
    def __init__(self, operation):
        self.operation = operation
