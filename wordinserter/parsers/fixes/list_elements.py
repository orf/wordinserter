from wordinserter.operations import BaseList, ListElement


def normalize_list_elements(tokens):
    for token in tokens:
        if isinstance(token, BaseList):
            normalize_list(token)
        else:
            normalize_list_elements(token.children)


def normalize_list(op: BaseList):
    # If there are > 1 lists to move out then we need to insert it after previously moved ones,
    # instead of before. `moved` tracks this.
    children = list(op)

    for child in children:
        if isinstance(child, ListElement):
            moved = 0
            for element_child in child:
                if isinstance(element_child, BaseList):
                    moved += 1
                    # Move the list outside of the ListElement
                    child_index = op.child_index(child)
                    op.insert_child(child_index + moved, element_child)
                    child.remove_child(element_child)
                    normalize_list(element_child)
