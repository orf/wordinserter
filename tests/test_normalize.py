from wordinserter.operations import Table, TableRow, TableCell
from wordinserter.parsers.fixes import table_colspans
import pytest


class TestNormalizeTable:
    # Lists of (table_length, given, expected) table colspans
    TABLES = [
        (8, (5, 1, 3), (5, 1, 2)),  # Reduce last column by 1
        (8, (3, 1, 4, 4), (3, 1, 3, 1)),  # Reduce 3rd column by one, last column to 1
        (8, (3, 2, 4, 4), (3, 2, 2, 1)),  # Reduce 2nd column, third and fourth
        (8, (3, 1, 3), (3, 1, 4)),  # Expand fourth column,
        *[(i, (8,), (i,)) for i in range(1, 7)]  # Reduce single column colspans to single values
    ]

    @pytest.mark.parametrize('table_length, given_spans,expected_spans', TABLES)
    def test_normalize_rowspans(self, table_length, given_spans, expected_spans):
        given_row = TableRow(
            *(TableCell(colspan=span, rowspan=1) for span in given_spans)
        )

        eight_column_table = Table(
            TableRow(
                [TableCell(colspan=1, rowspan=1) for _ in range(table_length)]
            ),
            given_row
        )

        table_colspans.normalize_table(eight_column_table)

        assert tuple(c.colspan for c in given_row.children) == expected_spans
