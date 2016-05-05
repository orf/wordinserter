import abc
import inspect
from wordinserter.operations import ChildlessOperation, IgnoredOperation, Group, Text
from wordinserter.exceptions import InsertError
import contextlib
from collections.abc import Iterable


def renders(*operations):
    def _wrapper(func):
        func.renders_operations = operations

        if any(isinstance(op, ChildlessOperation) for op in operations):
            if not all(isinstance(op, ChildlessOperation) for op in operations):
                raise Exception("Cannot mix ChildlessOperations and normal Operations")

            return func

        return contextlib.contextmanager(func)

    return _wrapper


class NewOperations(object):
    def __init__(self, ops):
        self.ops = ops


class BaseRenderer(abc.ABC):
    def __init__(self, debug=False, hooks=None):
        self.debug = debug
        self.render_methods = {}
        self.hooks = hooks or {}

        for name, method in inspect.getmembers(self, inspect.ismethod):
            if hasattr(method, "renders_operations"):
                for op in method.renders_operations:
                    self.render_methods[op] = method

    def _call_hook(self, key, operation):
        cls = operation.__class__
        if key in self.hooks and cls in self.hooks[key]:
            hooks = self.hooks[key][cls]
            if not isinstance(hooks, Iterable):
                hooks = [hooks]

            for hook in hooks:
                hook(operation, self)

    def new_operations(self, operations):
        return NewOperations(operations)

    @contextlib.contextmanager
    def with_hooks(self, operation):
        self._call_hook("pre", operation)
        yield
        self._call_hook("post", operation)

    @renders(IgnoredOperation, Group)
    def ignored_element(self, *args, **kwargs):
        yield

    def render(self, operations, args=None, indent=0):
        for operation in operations:
            method = self.render_methods.get(operation.__class__, None)
            if method is None:
                raise NotImplementedError(
                    "Operation {0} not supported by this renderer".format(operation.__class__.__name__))

            if operation.format is not None \
                    and operation.format.has_format() \
                    and operation.format.__class__ in self.render_methods:
                format_func = self.render_methods[operation.format.__class__]
            else:
                format_func = self.ignored_element

            try:
                with self.with_hooks(operation), format_func(operation.format, operation):
                    if isinstance(operation, ChildlessOperation):

                        if self.debug:
                            output = operation.__class__.__name__ \
                                if not isinstance(operation, Text) else operation.short_text

                            output = output.encode(errors="replace")

                            print((" " * indent) + str(output))

                        method(operation, *args or [])

                    else:
                        if self.debug:
                            method = DebugMethod(method, indent)

                        with method(operation, *args or []) as new_args:
                            if isinstance(new_args, NewOperations):
                                self.render(new_args.ops, None, indent + 1)
                            else:
                                self.render(operation.children, new_args, indent + 1)
            except InsertError:
                raise
            except Exception as e:
                raise InsertError(operation) from e


class DebugMethod(object):
    def __init__(self, method, indent):
        self.method = method
        self.indent = indent
        self.operation = None

    def __call__(self, operation, *args):
        self.operation = operation
        self.inner_manager = self.method(operation, *args)
        return self

    def __enter__(self):
        print((" " * self.indent) + self.operation.__class__.__name__
              + " " + (str(self.operation.format) if self.operation.format is not None else ""))
        return self.inner_manager.__enter__()

    def __exit__(self, *args):
        print((" " * self.indent) + "/" + self.operation.__class__.__name__
              + " " + (str(self.operation.format) if self.operation.format is not None else ""))
        return self.inner_manager.__exit__(*args)


from .com import COMRenderer
