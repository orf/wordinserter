from wordinserter.operations import Table


def normalize_table_colspans(tokens):
    for token in tokens:
        if isinstance(token, Table):
            normalize_table(token)
            token.update_child_widths()
        else:
            normalize_table_colspans(token.children)


def normalize_table(table: Table):
    if any(any(cell.rowspan > 1 for cell in row.children) for row in table.children):
        # Cannot normalize tables with rowspans just yet.
        return

    max_table_cells = max(len(row.children) for row in table.children)
    for row_idx, row in enumerate(table.children):
        children_with_colspan = [child for child in row.children if child.colspan > 1]
        colspan_left = max_table_cells - (len(row.children) - len(children_with_colspan))

        # Travel up and find any previous cells with a 'rowspan' that can impact us
        for idx, child in enumerate(children_with_colspan):
            child_wants_colspan = child.colspan
            if child_wants_colspan >= colspan_left:
                child.colspan = colspan_left - len(children_with_colspan[idx+1:])
            elif child == children_with_colspan[-1]:
                child.colspan = colspan_left

            colspan_left -= child.colspan
