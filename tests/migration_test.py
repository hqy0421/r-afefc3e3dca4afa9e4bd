import unittest

from gspread.exceptions import InvalidInputValue
from gspread.migration import (
    CELL_TYPE_BLANK,
    CELL_TYPE_ERROR,
    CELL_TYPE_FORMULA,
    CELL_TYPE_FORMULA_LIKE_TEXT,
    CELL_TYPE_VALUE,
    MigrationPlanError,
    build_migration_preview,
    classify_cell,
    normalize_column_mapping,
    parse_finite_a1_range,
    rewrite_formula,
)


def entered(**kwargs):
    return {"userEnteredValue": kwargs}


class NormalizeColumnMappingTest(unittest.TestCase):
    def test_letters_and_indexes(self):
        self.assertEqual(
            normalize_column_mapping({"A": 3, "B": "D"}),
            {1: 3, 2: 4},
        )
        self.assertEqual(
            normalize_column_mapping({1: 10, 2: 11}),
            {1: 10, 2: 11},
        )

    def test_empty_mapping_rejected(self):
        with self.assertRaises(InvalidInputValue):
            normalize_column_mapping({})

    def test_duplicate_destination_rejected(self):
        with self.assertRaises(InvalidInputValue):
            normalize_column_mapping({"A": "Z", "B": "Z"})

    def test_invalid_column_rejected(self):
        with self.assertRaises(InvalidInputValue):
            normalize_column_mapping({"nope": "A"})
        with self.assertRaises(InvalidInputValue):
            normalize_column_mapping({"A": 0})


class ParseFiniteRangeTest(unittest.TestCase):
    def test_finite_range(self):
        self.assertEqual(parse_finite_a1_range("A1:C4"), (1, 4, 1, 3))

    def test_single_cell(self):
        self.assertEqual(parse_finite_a1_range("B2"), (2, 2, 2, 2))

    def test_reversed_input_is_normalized(self):
        self.assertEqual(parse_finite_a1_range("C4:A1"), (1, 4, 1, 3))

    def test_full_columns_rejected(self):
        with self.assertRaises(InvalidInputValue):
            parse_finite_a1_range("A:B")

    def test_open_ended_rejected(self):
        with self.assertRaises(InvalidInputValue):
            parse_finite_a1_range("A1:B")


class RewriteFormulaTest(unittest.TestCase):
    MAPPING = {1: 5, 2: 6, 3: 7}  # A->E, B->F, C->G

    def rewrite(self, formula):
        return rewrite_formula(formula, "Z99", self.MAPPING, "Old", "New")

    def test_unqualified_cells_are_mapped(self):
        rewritten, issues, _ = self.rewrite("=A1+B2")
        self.assertEqual(rewritten, "=E1+F2")
        self.assertEqual(issues, [])

    def test_dollar_markers_are_preserved(self):
        rewritten, _, _ = self.rewrite("=$A$1*B$2")
        self.assertEqual(rewritten, "=$E$1*F$2")

    def test_finite_rectangle_stays_contiguous_and_ordered(self):
        rewritten, issues, _ = self.rewrite("=A1:C10")
        self.assertEqual(rewritten, "=E1:G10")
        self.assertEqual(issues, [])

    def test_rectangle_keeps_dollar_markers(self):
        rewritten, _, _ = self.rewrite("=$A$1:$C$10")
        self.assertEqual(rewritten, "=$E$1:$G$10")

    def test_explicit_old_sheet_reference_is_retargeted(self):
        rewritten, issues, _ = self.rewrite("='Old'!A1")
        self.assertEqual(rewritten, "='New'!E1")
        self.assertEqual(issues, [])

    def test_bare_old_sheet_title_is_retargeted_and_quoted(self):
        rewritten, _, _ = self.rewrite("=SUM(Old!A1)")
        self.assertEqual(rewritten, "=SUM('New'!E1)")

    def test_qualified_rectangle_is_retargeted(self):
        rewritten, _, _ = self.rewrite("='Old'!A1:C5")
        self.assertEqual(rewritten, "='New'!E1:G5")

    def test_other_sheet_references_are_preserved(self):
        rewritten, issues, _ = self.rewrite("=Sheet2!A1:C5+A1")
        self.assertEqual(rewritten, "=Sheet2!A1:C5+E1")
        self.assertEqual(issues, [])

    def test_named_ranges_are_preserved(self):
        rewritten, _, _ = self.rewrite("=MYRANGE+A1")
        self.assertEqual(rewritten, "=MYRANGE+E1")

    def test_qualified_named_range_is_preserved(self):
        rewritten, _, _ = self.rewrite("='Old'!MyRange")
        self.assertEqual(rewritten, "='Old'!MyRange")

    def test_text_inside_string_literal_is_preserved(self):
        rewritten, _, _ = self.rewrite('=A1&"A1 stays"')
        self.assertEqual(rewritten, '=E1&"A1 stays"')

    def test_dynamic_function_inside_string_is_preserved(self):
        rewritten, issues, _ = self.rewrite('="INDIRECT(A1)"')
        self.assertEqual(rewritten, '="INDIRECT(A1)"')
        self.assertEqual(issues, [])

    def test_full_row_references_are_preserved(self):
        rewritten, _, _ = self.rewrite("=1:1+A2")
        self.assertEqual(rewritten, "=1:1+E2")

    def test_numeric_literals_are_not_treated_as_references(self):
        rewritten, _, refs = self.rewrite("=B2*2+1.5")
        self.assertEqual(rewritten, "=F2*2+1.5")
        originals = [reference.original for reference in refs]
        self.assertEqual(originals, ["B2"])
        self.assertNotIn("2", originals)
        self.assertNotIn("1.5", originals)

    def test_function_sharing_a_cell_prefix_is_not_a_reference(self):
        rewritten, _, _ = self.rewrite("=LOG10(A1)")
        self.assertEqual(rewritten, "=LOG10(E1)")

    # -- failure modes -----------------------------------------------------

    def assert_issue(self, formula, code):
        rewritten, issues, _ = self.rewrite(formula)
        self.assertIsNone(rewritten)
        self.assertTrue(any(issue.code == code for issue in issues))
        self.assertTrue(all(issue.cell == "Z99" for issue in issues))

    def test_dynamic_indirect_aborts(self):
        self.assert_issue('=INDIRECT("A1")', "dynamic_reference")

    def test_dynamic_offset_aborts(self):
        self.assert_issue("=OFFSET(A1,1,2)", "dynamic_reference")

    def test_reversed_range_aborts(self):
        self.assert_issue("=C1:A1", "reversed_range")

    def test_unmapped_column_in_cell_aborts(self):
        self.assert_issue("=D1", "unmapped_column")

    def test_unmapped_column_in_range_aborts(self):
        self.assert_issue("=A1:D1", "unmapped_column")

    def test_full_column_aborts(self):
        self.assert_issue("=SUM(A:A)", "unbounded_reference")

    def test_open_ended_range_aborts(self):
        self.assert_issue("=A1:B", "unbounded_reference")

    def test_split_mapping_with_gap_is_allowed(self):
        # Destination columns may have gaps; only ordering must be preserved.
        rewritten, issues, _ = rewrite_formula(
            "=A1:B2", "K3", {1: 5, 2: 8}, "Old", "New"
        )
        self.assertEqual(rewritten, "=E1:H2")
        self.assertEqual(issues, [])

    def test_reordered_mapping_aborts(self):
        rewritten, issues, _ = rewrite_formula(
            "=A1:B2", "K3", {1: 7, 2: 6}, "Old", "New"
        )
        self.assertIsNone(rewritten)
        self.assertEqual(issues[0].code, "split_or_reordered_range")


class ClassifyCellTest(unittest.TestCase):
    def test_blank(self):
        self.assertEqual(classify_cell({}), (CELL_TYPE_BLANK, ""))
        self.assertEqual(
            classify_cell({"userEnteredValue": {}}),
            (CELL_TYPE_BLANK, ""),
        )

    def test_number_bool_string(self):
        self.assertEqual(
            classify_cell(entered(numberValue=3.5)),
            (CELL_TYPE_VALUE, 3.5),
        )
        self.assertEqual(
            classify_cell(entered(boolValue=True)),
            (CELL_TYPE_VALUE, True),
        )
        self.assertEqual(
            classify_cell(entered(stringValue="hello")),
            (CELL_TYPE_VALUE, "hello"),
        )

    def test_formula(self):
        self.assertEqual(
            classify_cell(entered(formulaValue="=A1")),
            (CELL_TYPE_FORMULA, "=A1"),
        )

    def test_formula_like_text_stays_distinct(self):
        self.assertEqual(
            classify_cell(entered(stringValue="=not a formula")),
            (CELL_TYPE_FORMULA_LIKE_TEXT, "=not a formula"),
        )

    def test_error_value(self):
        cell_type, value = classify_cell(
            entered(errorValue={"type": "REF", "message": "#REF!"})
        )
        self.assertEqual(cell_type, CELL_TYPE_ERROR)
        self.assertEqual(value, "#REF!")


class FakeSpreadsheet:
    def __init__(self, grid_metadata, titles):
        self._grid_metadata = grid_metadata
        self._titles = titles
        self.metadata_calls = []

    def fetch_sheet_metadata(self, params=None):
        self.metadata_calls.append(params)
        if params and params.get("includeGridData") == "true":
            return self._grid_metadata
        return {"sheets": [{"properties": {"title": title}} for title in self._titles]}


class FakeWorksheet:
    id = 7
    title = "Old"

    def __init__(self, grid_metadata, titles):
        self.spreadsheet = FakeSpreadsheet(grid_metadata, titles)


def build_grid(rows):
    return {
        "sheets": [
            {
                "properties": {"sheetId": 7, "title": "Old"},
                "data": [
                    {
                        "startRow": 0,
                        "startColumn": 0,
                        "rowData": [
                            {"values": [entered(**cell) for cell in row]}
                            for row in rows
                        ],
                    }
                ],
            }
        ]
    }


class BuildMigrationPreviewTest(unittest.TestCase):
    MAPPING = {"A": "X", "B": "Y", "C": "Z"}

    def setUp(self):
        self.grid = build_grid(
            [
                [
                    {"stringValue": "name"},
                    {"numberValue": 42},
                    {"formulaValue": "=B2*2"},
                ],
                [
                    {"stringValue": "=not a formula"},
                    {"boolValue": True},
                    {"formulaValue": "=SUM(B1:B2)"},
                ],
                [
                    {},
                    {"numberValue": 1.5},
                    {"errorValue": {"type": "REF", "message": "#REF!"}},
                ],
            ]
        )

    def test_plan_is_read_only_and_keeps_rows(self):
        worksheet = FakeWorksheet(self.grid, ("Old", "New"))
        plan = build_migration_preview(worksheet, "New", self.MAPPING, "A1:C3")
        self.assertEqual(plan.row_start, 1)
        self.assertEqual(plan.row_end, 3)
        # only metadata reads happened, nothing write-shaped
        for call in worksheet.spreadsheet.metadata_calls:
            self.assertNotIn("valueInputOption", call or {})

    def test_cell_types_are_preserved(self):
        worksheet = FakeWorksheet(self.grid, ("Old", "New"))
        plan = build_migration_preview(worksheet, "New", self.MAPPING, "A1:C3")
        by_address = {entry.cell: entry for entry in plan.cells}

        self.assertEqual(by_address["A1"].cell_type, CELL_TYPE_VALUE)
        self.assertEqual(by_address["A1"].value, "name")
        self.assertEqual(by_address["B1"].value, 42)
        self.assertEqual(by_address["B2"].value, True)
        self.assertEqual(by_address["B3"].value, 1.5)
        self.assertEqual(by_address["A3"].cell_type, CELL_TYPE_BLANK)
        self.assertEqual(by_address["C3"].cell_type, CELL_TYPE_ERROR)

        formula_cell = by_address["C1"]
        self.assertEqual(formula_cell.cell_type, CELL_TYPE_FORMULA)
        self.assertEqual(formula_cell.formula, "=Y2*2")
        self.assertEqual(formula_cell.value, "=Y2*2")

        text_cell = by_address["A2"]
        self.assertEqual(text_cell.cell_type, CELL_TYPE_FORMULA_LIKE_TEXT)
        self.assertIsNone(text_cell.formula)
        self.assertEqual(text_cell.value, "=not a formula")

    def test_payloads_use_two_input_options_and_correct_ranges(self):
        worksheet = FakeWorksheet(self.grid, ("Old", "New"))
        plan = build_migration_preview(worksheet, "New", self.MAPPING, "A1:C3")
        groups = {group.value_input_option: group.data for group in plan.batch_groups}
        self.assertIn("USER_ENTERED", groups)
        self.assertIn("RAW", groups)

        entered_ranges = {item["range"] for item in groups["USER_ENTERED"]}
        self.assertIn("'New'!X2:X2", entered_ranges)
        self.assertIn("'New'!Z1:Z3", entered_ranges)

        entered_values = [
            value for item in groups["USER_ENTERED"] for value in item["values"]
        ]
        # formula-like text is apostrophe-prefixed so it stays text
        self.assertIn(["'=not a formula"], entered_values)
        # formulas keep their leading "=" without an apostrophe
        self.assertIn(["=Y2*2"], entered_values)
        # error value is reproduced as an error under USER_ENTERED
        self.assertIn(["#REF!"], entered_values)

        raw_ranges = {item["range"] for item in groups["RAW"]}
        self.assertIn("'New'!Y1:Y3", raw_ranges)
        raw_values = [value for item in groups["RAW"] for value in item["values"]]
        self.assertIn([42], raw_values)
        self.assertIn([True], raw_values)
        self.assertIn([1.5], raw_values)

    def test_unmapped_source_columns_are_skipped(self):
        # A grid without cross-column formulas so skipping B is unambiguous.
        grid = build_grid(
            [
                [
                    {"stringValue": "a1"},
                    {"numberValue": 99},
                    {"numberValue": 7},
                ],
                [
                    {"stringValue": "a2"},
                    {"numberValue": 100},
                    {"numberValue": 8},
                ],
            ]
        )
        worksheet = FakeWorksheet(grid, ("Old", "New"))
        plan = build_migration_preview(worksheet, "New", {"A": "X", "C": "Z"}, "A1:C2")
        # B is not in the mapping so its cells never enter the plan
        self.assertTrue(all(entry.col != 2 for entry in plan.cells))
        self.assertEqual(plan.column_mapping, {1: 24, 3: 26})
        addresses = {entry.cell for entry in plan.cells}
        self.assertEqual(addresses, {"A1", "A2", "C1", "C2"})

    def test_skipping_a_referenced_column_aborts(self):
        # Here a migrated formula references skipped column B, which must
        # abort the whole batch instead of silently producing a bad formula.
        worksheet = FakeWorksheet(self.grid, ("Old", "New"))
        with self.assertRaises(MigrationPlanError) as context:
            build_migration_preview(worksheet, "New", {"A": "X", "C": "Z"}, "A1:C3")
        self.assertTrue(
            all(issue.code == "unmapped_column" for issue in context.exception.issues)
        )

    def test_missing_destination_sheet_warns(self):
        worksheet = FakeWorksheet(self.grid, ("Old",))
        plan = build_migration_preview(worksheet, "Ghost", self.MAPPING, "A1:C3")
        self.assertEqual(len(plan.warnings), 1)
        self.assertIn("Ghost", plan.warnings[0])

    def test_empty_middle_row_keeps_its_position(self):
        # An empty row is a positional {} placeholder: rows below it must not
        # shift up in the destination payload.
        grid = build_grid(
            [
                [{"stringValue": "top"}],
                [{}],
                [{"stringValue": "bottom"}],
            ]
        )
        worksheet = FakeWorksheet(grid, ("Old", "New"))
        plan = build_migration_preview(worksheet, "New", {"A": "X"}, "A1:A3")
        by_row = {entry.row: entry for entry in plan.cells}
        self.assertEqual(by_row[1].value, "top")
        self.assertEqual(by_row[2].cell_type, CELL_TYPE_BLANK)
        self.assertEqual(by_row[3].value, "bottom")

        raw = next(
            group.data
            for group in plan.batch_groups
            if group.value_input_option == "RAW"
        )
        self.assertEqual(raw[0]["range"], "'New'!X1:X3")
        self.assertEqual(raw[0]["values"], [["top"], [""], ["bottom"]])

    def test_formula_errors_abort_the_whole_batch_with_cell_locations(self):
        grid = build_grid(
            [
                [
                    {"formulaValue": '=INDIRECT("A1")'},
                    {"formulaValue": "=A1:D1"},
                ]
            ]
        )
        worksheet = FakeWorksheet(grid, ("Old", "New"))
        with self.assertRaises(MigrationPlanError) as context:
            build_migration_preview(worksheet, "New", self.MAPPING, "A1:B1")
        issues = context.exception.issues
        codes = {(issue.cell, issue.code) for issue in issues}
        self.assertIn(("A1", "dynamic_reference"), codes)
        self.assertIn(("B1", "unmapped_column"), codes)
        # every issue points at a concrete source formula cell
        self.assertTrue(all(issue.cell for issue in issues))

    def test_to_dict_is_json_serializable(self):
        import json

        worksheet = FakeWorksheet(self.grid, ("Old", "New"))
        plan = build_migration_preview(worksheet, "New", self.MAPPING, "A1:C3")
        encoded = json.dumps(plan.to_dict())
        decoded = json.loads(encoded)
        self.assertEqual(decoded["source_sheet"], "Old")
        self.assertEqual(decoded["column_mapping"], {"A": "X", "B": "Y", "C": "Z"})
        self.assertEqual(len(decoded["cells"]), 9)


if __name__ == "__main__":
    unittest.main()
