import abc
import functools
import inspect
import contextlib
from ..operations import ChildlessOperation, IgnoredOperation, Group
import sys


def renders(*operations):
    def _wrapper(func):
        func.renders_operations = operations

        @functools.wraps(func)
        def _inner(*args, **kwargs):
            return func(*args, **kwargs)

        if any(isinstance(op, ChildlessOperation) for op in operations):
            if not all(isinstance(op, ChildlessOperation) for op in operations):
                raise Exception("Cannot mix ChildlessOperations and normal Operations")

            return func

        return contextlib.contextmanager(_inner)

    return _wrapper


class BaseRenderer(abc.ABC):
    def __init__(self, debug=False):
        self.debug = debug
        self.render_methods = {}
        for name, method in inspect.getmembers(self, inspect.ismethod):
            if hasattr(method, "renders_operations"):
                for op in method.renders_operations:
                    self.render_methods[op] = method

    @renders(IgnoredOperation, Group)
    def ignored_element(self, op):
        yield

    def render(self, operations):
        for op in operations:
            op.set_parents()

        self._render(operations)

    def _render(self, operations, args=None, indent=0):
        for operation in operations:
            method = self.render_methods.get(operation.__class__, None)
            if method is None:
                raise NotImplementedError("Operation {0} not supported by this renderer".format(operation.__class__.__name__))

            if operation.format is not None and operation.format.__class__ in self.render_methods:
                format_func = self.render_methods[operation.format.__class__]
            else:
                format_func = self.ignored_element

            with format_func(operation.format):
                if isinstance(operation, ChildlessOperation):

                    if self.debug:
                        print((" " * indent) + operation.__class__.__name__)

                    method(operation, *args or [])

                else:
                    if self.debug:
                        method = debug_method(method, indent)

                    with method(operation, *args or []) as new_args:
                        self._render(operation.children, new_args, indent+1)


import contextlib


class debug_method(object):
    def __init__(self, method, indent):
        self.method = method
        self.indent = indent
        self.operation = None

    def __call__(self, operation, *args):
        self.operation = operation
        self.inner_manager = self.method(operation, *args)
        return self

    def __enter__(self):
        print((" " * self.indent) + self.operation.__class__.__name__)
        return self.inner_manager.__enter__()

    def __exit__(self, *args):
        print((" " * self.indent) + "/" + self.operation.__class__.__name__)
        return self.inner_manager.__exit__(*args)



from .com import COMRenderer