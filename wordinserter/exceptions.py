
class InsertError(RuntimeError):
    def __init__(self, operation, cause):
        self.operation = operation
        self.cause = cause
