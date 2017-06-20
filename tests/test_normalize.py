from wordinserter.operations import Table, TableRow, TableCell
from wordinserter.parsers.fixes import table_colspans
import pytest


@pytest.fixture
def four_col_table():
    return Table(
        TableRow(
            [TableCell(colspan=1) for _ in range(4)]
        )
    )


@pytest.fixture
def eight_col_table():
    return Table(
        TableRow(
            [TableCell(colspan=1) for _ in range(8)]
        )
    )


class TestNormalizeTable:
    def test_rowspan_invalid_one(self, four_col_table: Table):
        invalid_row = TableRow(
                TableCell(colspan=4),
                TableCell(colspan=1),
                TableCell(colspan=2)
            )

        four_col_table.add_child(invalid_row)

        table_colspans.normalize_table(four_col_table)

        assert [c.colspan for c in invalid_row.children] == [2, 1, 1]

    def test_rowspan_invalid_two(self, eight_col_table: Table):
        invalid_row = TableRow(
            TableCell(colspan=5),
            TableCell(colspan=1),
            TableCell(colspan=3)
        )

        eight_col_table.add_child(invalid_row)

        table_colspans.normalize_table(eight_col_table)

        assert [c.colspan for c in invalid_row.children] == [5, 1, 2]

        assert invalid_row.children[0].colspan == 5
        assert invalid_row.children[1].colspan == 1
        assert invalid_row.children[2].colspan == 2

    def test_rowspan_invalid_three(self, eight_col_table: Table):
        invalid_row = TableRow(
            TableCell(colspan=3),
            TableCell(colspan=1),
            TableCell(colspan=4),
            TableCell(colspan=4)
        )

        eight_col_table.add_child(invalid_row)

        table_colspans.normalize_table(eight_col_table)

        assert [c.colspan for c in invalid_row.children] == [3, 1, 3, 1]

    def test_rowspan_invalid_four(self, eight_col_table: Table):
        invalid_row = TableRow(
            TableCell(colspan=3),
            TableCell(colspan=2),
            TableCell(colspan=4),
            TableCell(colspan=4)
        )

        eight_col_table.add_child(invalid_row)

        table_colspans.normalize_table(eight_col_table)

        assert [c.colspan for c in invalid_row.children] == [3, 2, 2, 1]

    def test_rowspan_invalid_end(self, eight_col_table: Table):
        invalid_row = TableRow(
            TableCell(colspan=3),
            TableCell(colspan=1),
            TableCell(colspan=3)
        )

        eight_col_table.add_child(invalid_row)

        table_colspans.normalize_table(eight_col_table)

        assert [c.colspan for c in invalid_row.children] == [3, 1, 4]
