from wordinserter.operations import Text, Operation, CodeBlock, Paragraph
import re

_COLLAPSE_REGEX = re.compile(r'\s+')


def correct_whitespace(tokens):
    remove_arbitrary_newlines(tokens)
    remove_paragraph_whitespace(tokens)
    return


def remove_paragraph_whitespace(parent_token: Operation):
    for token in list(parent_token):
        if isinstance(token, Paragraph):
            _inner_remove_paragraph_whitespace(token)
        else:
            remove_paragraph_whitespace(token)


def _inner_remove_paragraph_whitespace(parent_operation):
    if isinstance(parent_operation, Text):
        if parent_operation.text.startswith(" "):
            parent_operation.text = parent_operation.text.lstrip()
    elif not parent_operation.has_children:
        return
    else:
        target = parent_operation[0]
        _inner_remove_paragraph_whitespace(target)


def remove_arbitrary_newlines(parent_token: Operation):
    for token in list(parent_token):
        if isinstance(token, Text):
            if token.text.isspace() and not token.has_parent(Paragraph):
                parent_token.remove_child(token)
            else:
                token.text = _COLLAPSE_REGEX.sub(' ', token.text)
        elif isinstance(token, CodeBlock):
            # Ignore CodeBlocks
            continue
        else:
            remove_arbitrary_newlines(token)
