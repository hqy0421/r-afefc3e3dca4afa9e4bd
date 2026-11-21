"""
gspread.migration
~~~~~~~~~~~~~~~~~

Build *migration previews* for moving columns from one worksheet of a
spreadsheet to another worksheet of the **same** spreadsheet, keeping the row
numbers unchanged.

A preview only computes a plan. It never writes, clears or deletes anything.
The returned :class:`MigrationPlan` holds:

* every cell of the source range with its *original* type preserved (text,
  formula, number, boolean, error, blank), including literal text that starts
  with ``=`` (which is deliberately kept distinct from real formulas);
* the column mapping and ready-to-send payload groups for
  :meth:`gspread.worksheet.Worksheet.batch_update`;
* every formula reference that was rewritten, and why.

Formula rewriting rules:

* unqualified A1 references (``A1``) and references explicitly pointing at the
  source sheet (``'Old'!A1``) are mapped with the column mapping and the
  destination sheet is used;
* ``$`` absolute markers are preserved;
* references to other sheets, named ranges and text inside string literals are
  left untouched;
* a finite rectangular reference can only be migrated when every column it
  covers is present in the mapping and the mapped columns stay contiguous and
  in the same order. Missing columns, split ranges, reversed ranges,
  open-ended references or dynamic references (``INDIRECT`` / ``OFFSET``) abort
  the whole batch and report the offending source formula cell.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import (
    TYPE_CHECKING,
    Any,
    Dict,
    List,
    Mapping,
    MutableMapping,
    Optional,
    Sequence,
    Tuple,
)

from .exceptions import GSpreadException, InvalidInputValue
from .utils import (
    ValueInputOption,
    _a1_to_rowcol_unbounded,
    column_letter_to_index,
    rowcol_to_a1,
)

if TYPE_CHECKING:
    from .worksheet import Worksheet


# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------


class MigrationPlanError(GSpreadException):
    """Raised when a migration preview cannot be built safely.

    All the problems detected while scanning the source range are collected in
    :attr:`issues` instead of failing on the first one, so a single run reports
    every offending cell.
    """

    def __init__(self, issues: Sequence["MigrationIssue"]):
        self.issues: List[MigrationIssue] = list(issues)
        summary = "; ".join(str(issue) for issue in self.issues)
        super().__init__(
            "migration preview aborted with {} issue(s): {}".format(
                len(self.issues), summary
            )
        )


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

#: A cell holds a plain, non-formula value.
CELL_TYPE_VALUE = "value"
#: A cell holds a formula (its stored value starts with ``=``).
CELL_TYPE_FORMULA = "formula"
#: A cell holds literal text that starts with ``=`` but is not a formula.
CELL_TYPE_FORMULA_LIKE_TEXT = "formula_like_text"
#: An empty cell.
CELL_TYPE_BLANK = "blank"
#: A cell that holds one of the spreadsheet error values (``#REF!`` ...).
CELL_TYPE_ERROR = "error"

_DYNAMIC_FUNCTIONS = frozenset({"INDIRECT", "OFFSET"})

#: A sheet title prefix, quoted (``'My Sheet'!``) or bare (``Sheet1!``).
#: Bare titles are restricted to word characters so a single-letter title
#: (``O!A1``) stays distinguishable from a cell reference followed by ``!``.
_SHEET_PREFIX_RE = re.compile(r"(?:'(?:[^']|'')+'|[\w.]+)!")

#: A function call, e.g. ``SUM(`` or ``OFFSET (``.
_FUNCTION_CALL_RE = re.compile(r"([A-Za-z_]\w*(?:\.[A-Za-z_]\w*)*)\s*\(")

_COL = r"\$?[A-Za-z]{1,3}"
_ROW = r"\$?[1-9]\d*"
_CELL = r"(?:\$?[A-Za-z]{1,3})(?:\$?[1-9]\d*)"

# A valid A1 token ends at a formula boundary: an operator, parenthesis,
# comma, semicolon, whitespace, another reference prefix ("!" / ":") or the
# end of the formula. A word character means we only matched a prefix of a
# longer identifier (e.g. the "MYR" inside "MYRANGE"), which is a named range
# and must be preserved.
_BOUNDARY = r"(?![A-Za-z0-9_])"
# The token may not start in the middle of a longer identifier either. The
# optional "$" marker makes the lookbehind fixed-width.
_START_BOUNDARY = r"(?<![\w$])"

# Ordered A1 component patterns. The first match wins, so open-ended forms
# must be listed before the plain cell pattern.
_COMPONENT_PATTERNS: List[Tuple[re.Pattern, str]] = [
    (
        re.compile(_START_BOUNDARY + r"{}:{}{}".format(_CELL, _CELL, _BOUNDARY)),
        "rectangle",
    ),
    (
        re.compile(_START_BOUNDARY + r"{}:{}{}".format(_COL, _COL, _BOUNDARY)),
        "column_range",
    ),
    (
        re.compile(_START_BOUNDARY + r"{}:{}{}".format(_ROW, _ROW, _BOUNDARY)),
        "row_range",
    ),
    # Open-ended mixed references (one endpoint names only a column or row).
    (
        re.compile(_START_BOUNDARY + r"{}:{}{}".format(_CELL, _COL, _BOUNDARY)),
        "open_ended",
    ),
    (
        re.compile(_START_BOUNDARY + r"{}:{}{}".format(_COL, _CELL, _BOUNDARY)),
        "open_ended",
    ),
    (
        re.compile(_START_BOUNDARY + r"{}:{}{}".format(_CELL, _ROW, _BOUNDARY)),
        "open_ended",
    ),
    (
        re.compile(_START_BOUNDARY + r"{}:{}{}".format(_ROW, _CELL, _BOUNDARY)),
        "open_ended",
    ),
    (re.compile(_START_BOUNDARY + _CELL + _BOUNDARY), "cell"),
    (re.compile(_START_BOUNDARY + _COL + _BOUNDARY), "bare_column"),
    # Note: there is deliberately no "bare row" (a lone number) pattern. A
    # standalone number in a formula is a numeric literal, not a reference;
    # full-row references are only the explicit "1:1" / "2:5" row_range form.
]


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------


@dataclass
class MigrationIssue:
    """A single reason why the migration preview cannot be completed.

    :ivar str code: machine-readable issue code
    :ivar str cell: A1 address of the *source formula cell* the issue was found
        in (or the offending reference itself when no cell is involved)
    :ivar str reference: the A1 token that could not be mapped
    :ivar str message: human-readable explanation
    """

    code: str
    cell: str
    reference: str
    message: str

    def __str__(self) -> str:
        return "[{}] {} (reference: {!r})".format(self.code, self.cell, self.reference)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "code": self.code,
            "cell": self.cell,
            "reference": self.reference,
            "message": self.message,
        }


@dataclass
class CellEntry:
    """One source cell, typed.

    :ivar str cell: A1 address in the source sheet
    :ivar int row: 1-based row
    :ivar int col: 1-based source column
    :ivar str cell_type: one of the ``CELL_TYPE_*`` constants
    :ivar value: the value as stored in the source cell; formulas keep their
        leading ``=`` and formula-like text keeps its literal text
    :ivar Optional[str] formula: the rewritten formula when
        ``cell_type == CELL_TYPE_FORMULA``, otherwise ``None``
    """

    cell: str
    row: int
    col: int
    cell_type: str
    value: Any
    formula: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            "cell": self.cell,
            "row": self.row,
            "col": self.col,
            "cell_type": self.cell_type,
            "value": self.value,
            "formula": self.formula,
        }


@dataclass
class BatchGroup:
    """A ready-to-send group for
    :meth:`gspread.worksheet.Worksheet.batch_update`.

    All ranges inside a group must be written with the same
    ``value_input_option``.

    :ivar str value_input_option: ``RAW`` or ``USER_ENTERED``
    :ivar list data: ``[{"range": ..., "values": [[...]]}, ...]`` with ranges
        already qualified by the destination sheet title
    """

    value_input_option: str
    data: List[MutableMapping[str, Any]] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "value_input_option": self.value_input_option,
            "data": self.data,
        }


@dataclass
class RewrittenReference:
    """Record of one formula reference rewrite."""

    cell: str
    original: str
    rewritten: str
    kind: str

    def to_dict(self) -> Dict[str, Any]:
        return {
            "cell": self.cell,
            "original": self.original,
            "rewritten": self.rewritten,
            "kind": self.kind,
        }


@dataclass
class MigrationPlan:
    """The complete, inert migration plan.

    Nothing in this object performs a write. The payloads in
    :attr:`batch_groups` can be handed to
    :meth:`gspread.worksheet.Worksheet.batch_update` once the caller has
    reviewed the plan.
    """

    source_sheet: str
    source_range: str
    destination_sheet: str
    column_mapping: Dict[int, int]
    row_start: int
    row_end: int
    cells: List[CellEntry]
    batch_groups: List[BatchGroup]
    rewritten_references: List[RewrittenReference]
    warnings: List[str]

    def to_dict(self) -> Dict[str, Any]:
        return {
            "source_sheet": self.source_sheet,
            "source_range": self.source_range,
            "destination_sheet": self.destination_sheet,
            "column_mapping": {
                _column_letter(source_col): _column_letter(destination_col)
                for source_col, destination_col in sorted(self.column_mapping.items())
            },
            "row_start": self.row_start,
            "row_end": self.row_end,
            "cells": [entry.to_dict() for entry in self.cells],
            "batch_groups": [group.to_dict() for group in self.batch_groups],
            "rewritten_references": [
                reference.to_dict() for reference in self.rewritten_references
            ],
            "warnings": list(self.warnings),
        }


# ---------------------------------------------------------------------------
# Mapping helpers
# ---------------------------------------------------------------------------


def normalize_column_mapping(mapping: Mapping[Any, Any]) -> Dict[int, int]:
    """Normalize a user-provided column mapping to ``{source_col: dest_col}``
    with 1-based integer column indexes.

    Keys and values may be column letters (``"C"``) or 1-based indexes
    (``3``).
    """
    if not mapping:
        raise InvalidInputValue("column mapping must not be empty")

    normalized: Dict[int, int] = {}
    for source, destination in mapping.items():
        source_col = _column_to_index(source)
        destination_col = _column_to_index(destination)
        if source_col in normalized:
            raise InvalidInputValue(
                "duplicate source column in mapping: {!r}".format(source)
            )
        normalized[source_col] = destination_col

    if len(set(normalized.values())) != len(normalized):
        raise InvalidInputValue(
            "two source columns are mapped to the same destination column"
        )

    return normalized


def _column_to_index(value: Any) -> int:
    if isinstance(value, bool):
        raise InvalidInputValue(
            "column must be a letter or an index: {!r}".format(value)
        )
    if isinstance(value, int):
        index = value
    elif isinstance(value, str):
        text = value.strip()
        # Spreadsheet column letters are 1 to 3 letters long; a longer word
        # is a named range or a typo, not a column.
        if not re.fullmatch(r"[A-Za-z]{1,3}", text):
            raise InvalidInputValue("invalid column in mapping: {!r}".format(value))
        try:
            index = column_letter_to_index(text)
        except InvalidInputValue:
            raise InvalidInputValue("invalid column in mapping: {!r}".format(value))
    else:
        raise InvalidInputValue(
            "column must be a letter or an index: {!r}".format(value)
        )

    if index < 1:
        raise InvalidInputValue("column indexes start at 1: {!r}".format(value))
    return index


def parse_finite_a1_range(range_name: str) -> Tuple[int, int, int, int]:
    """Parse a finite rectangular A1 range.

    :returns: ``(row_start, row_end, col_start, col_end)``, 1-based and
        normalized so each pair is ascending.
    :raises InvalidInputValue: if the range is missing or unbounded (full
        columns, full rows, open-ended ranges).
    """
    start_label, separator, end_label = range_name.partition(":")
    if not separator:
        end_label = start_label

    try:
        start_row, start_col = _a1_to_rowcol_unbounded(start_label)
        end_row, end_col = _a1_to_rowcol_unbounded(end_label)
    except Exception:
        raise InvalidInputValue(
            "source range must be a finite A1 range: {!r}".format(range_name)
        )

    if not (
        isinstance(start_row, int)
        and isinstance(start_col, int)
        and isinstance(end_row, int)
        and isinstance(end_col, int)
    ):
        raise InvalidInputValue(
            "source range must be a finite rectangle; unbounded references "
            "(full columns/rows or open-ended ranges) are not supported: "
            "{!r}".format(range_name)
        )

    return (
        min(start_row, end_row),
        max(start_row, end_row),
        min(start_col, end_col),
        max(start_col, end_col),
    )


# ---------------------------------------------------------------------------
# Formula tokenizing / rewriting
# ---------------------------------------------------------------------------


def _column_letter(col_index: int) -> str:
    return rowcol_to_a1(1, col_index)[:-1]


def _format_ref(
    col_dollar: bool,
    col: int,
    row_dollar: bool,
    row: int,
) -> str:
    return "{}{}{}{}".format(
        "$" if col_dollar else "",
        _column_letter(col),
        "$" if row_dollar else "",
        row,
    )


def _unquote_sheet_title(raw_title: str) -> str:
    if raw_title.startswith("'") and raw_title.endswith("'"):
        return raw_title[1:-1].replace("''", "'")
    return raw_title


def _quote_sheet_title(title: str) -> str:
    return "'{}'".format(title.replace("'", "''"))


def _match_component(text: str, cursor: int) -> Optional[Tuple[re.Match, str]]:
    for pattern, kind in _COMPONENT_PATTERNS:
        match = pattern.match(text, cursor)
        if match:
            return match, kind
    return None


def _consume_string_literal(formula: str, cursor: int) -> str:
    """Return the full quoted string starting at ``cursor`` (``"`` included).

    Doubled quotes (``""`` inside a literal) are treated as escapes.
    """
    end = cursor + 1
    while end < len(formula):
        if formula[end] != '"':
            end += 1
        elif end + 1 < len(formula) and formula[end + 1] == '"':
            end += 2
        else:
            return formula[cursor : end + 1]
    return formula[cursor:end]


def _consume_function_call(
    formula: str,
    cursor: int,
    cell_address: str,
    issues: List[MigrationIssue],
) -> Optional[re.Match]:
    """If a function call starts at ``cursor``, consume it and record an issue
    for dynamic reference functions. Returns the match (or ``None``)."""
    function_match = _FUNCTION_CALL_RE.match(formula, cursor)
    if not function_match:
        return None

    function_name = function_match.group(1).split(".")[-1].upper()
    if function_name in _DYNAMIC_FUNCTIONS:
        issues.append(
            MigrationIssue(
                code="dynamic_reference",
                cell=cell_address,
                reference=function_match.group(0),
                message=(
                    "dynamic reference function {}() cannot be remapped "
                    "statically".format(function_name)
                ),
            )
        )
    return function_match


def rewrite_formula(
    formula: str,
    cell_address: str,
    mapping: Mapping[int, int],
    source_sheet: str,
    destination_sheet: str,
) -> Tuple[Optional[str], List[MigrationIssue], List[RewrittenReference]]:
    """Rewrite the A1 references of one formula.

    :returns: ``(rewritten_formula, issues, rewritten)``. ``rewritten_formula``
        is ``None`` when at least one issue was found; in that case the whole
        migration must be aborted.
    """
    issues: List[MigrationIssue] = []
    rewritten: List[RewrittenReference] = []
    output: List[str] = []
    cursor = 0

    while cursor < len(formula):
        if formula[cursor] == '"':
            token = _consume_string_literal(formula, cursor)
            output.append(token)
            cursor += len(token)
            continue

        function_match = _consume_function_call(formula, cursor, cell_address, issues)
        if function_match:
            output.append(function_match.group(0))
            cursor = function_match.end()
            continue

        prefix_match = _SHEET_PREFIX_RE.match(formula, cursor)
        if prefix_match:
            cursor = _consume_qualified_reference(
                formula,
                cursor,
                prefix_match,
                cell_address,
                mapping,
                source_sheet,
                destination_sheet,
                issues,
                rewritten,
                output,
            )
            continue

        component = _match_component(formula, cursor)
        if component:
            component_match, kind = component
            output.append(
                _rewrite_local_component(
                    token=component_match.group(0),
                    kind=kind,
                    cell_address=cell_address,
                    mapping=mapping,
                    issues=issues,
                    rewritten=rewritten,
                )
            )
            cursor = component_match.end()
            continue

        output.append(formula[cursor])
        cursor += 1

    if issues:
        return None, issues, rewritten
    return "".join(output), issues, rewritten


def _consume_qualified_reference(
    formula: str,
    cursor: int,
    prefix_match: re.Match,
    cell_address: str,
    mapping: Mapping[int, int],
    source_sheet: str,
    destination_sheet: str,
    issues: List[MigrationIssue],
    rewritten: List[RewrittenReference],
    output: List[str],
) -> int:
    """Handle a sheet-prefixed token and return the new cursor position."""
    component = _match_component(formula, prefix_match.end())
    if component is None:
        # Named range or anything unparsable after the prefix: preserve the
        # prefix alone and continue.
        output.append(prefix_match.group(0))
        return prefix_match.end()

    component_match, kind = component
    token = formula[cursor : component_match.end()]
    output.append(
        _rewrite_qualified_component(
            token=token,
            prefix=prefix_match.group(0),
            component_text=component_match.group(0),
            kind=kind,
            cell_address=cell_address,
            mapping=mapping,
            source_sheet=source_sheet,
            destination_sheet=destination_sheet,
            issues=issues,
            rewritten=rewritten,
        )
    )
    return component_match.end()


def _rewrite_qualified_component(
    token: str,
    prefix: str,
    component_text: str,
    kind: str,
    cell_address: str,
    mapping: Mapping[int, int],
    source_sheet: str,
    destination_sheet: str,
    issues: List[MigrationIssue],
    rewritten: List[RewrittenReference],
) -> str:
    # ``prefix`` includes the trailing "!"; the title is everything before it.
    sheet_title = _unquote_sheet_title(prefix[:-1])

    if sheet_title != source_sheet:
        # References to other worksheets are preserved verbatim.
        rewritten.append(
            RewrittenReference(cell_address, token, token, "other_sheet_preserved")
        )
        return token

    replacement, local_kind = _map_component(
        component_text, kind, cell_address, mapping, issues
    )
    if (
        issues
        and issues[-1].cell == cell_address
        and issues[-1].reference
        in (
            component_text,
            token,
        )
    ):
        return token

    new_token = "{}!{}".format(_quote_sheet_title(destination_sheet), replacement)
    rewritten.append(RewrittenReference(cell_address, token, new_token, local_kind))
    return new_token


def _rewrite_local_component(
    token: str,
    kind: str,
    cell_address: str,
    mapping: Mapping[int, int],
    issues: List[MigrationIssue],
    rewritten: List[RewrittenReference],
) -> str:
    replacement, replacement_kind = _map_component(
        token, kind, cell_address, mapping, issues
    )
    if not (
        issues and issues[-1].cell == cell_address and issues[-1].reference == token
    ):
        rewritten.append(
            RewrittenReference(cell_address, token, replacement, replacement_kind)
        )
    return replacement


def _map_component(
    token: str,
    kind: str,
    cell_address: str,
    mapping: Mapping[int, int],
    issues: List[MigrationIssue],
) -> Tuple[str, str]:
    """Map one A1 component (without sheet prefix).

    Returns ``(replacement_text, kind)``. Appends an issue and returns the
    original token when the component cannot be mapped.
    """
    if kind in ("column_range", "bare_column", "open_ended"):
        issues.append(
            MigrationIssue(
                code="unbounded_reference",
                cell=cell_address,
                reference=token,
                message=(
                    "unbounded or open-ended reference is not a finite "
                    "rectangle and cannot be remapped"
                ),
            )
        )
        return token, kind

    if kind == "row_range":
        # Full-row references name no column, so there is nothing to remap.
        return token, "row_reference_preserved"

    if kind == "cell":
        return _map_single_cell(token, cell_address, mapping, issues)

    if kind == "rectangle":
        return _map_rectangle(token, cell_address, mapping, issues)

    # Unknown component kinds are preserved.
    return token, "preserved"


def _map_single_cell(
    token: str,
    cell_address: str,
    mapping: Mapping[int, int],
    issues: List[MigrationIssue],
) -> Tuple[str, str]:
    match = re.fullmatch(r"(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)", token)
    assert match is not None
    col_abs, col_letter, row_abs, row_label = match.groups()
    col_index = column_letter_to_index(col_letter)

    if col_index not in mapping:
        issues.append(
            MigrationIssue(
                code="unmapped_column",
                cell=cell_address,
                reference=token,
                message=(
                    "column {} is not part of the column mapping".format(
                        col_letter.upper()
                    )
                ),
            )
        )
        return token, "unmapped"

    replacement = _format_ref(
        bool(col_abs),
        mapping[col_index],
        bool(row_abs),
        int(row_label),
    )
    return replacement, "cell"


def _map_rectangle(
    token: str,
    cell_address: str,
    mapping: Mapping[int, int],
    issues: List[MigrationIssue],
) -> Tuple[str, str]:
    match = re.fullmatch(
        r"(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*):" r"(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)",
        token,
    )
    assert match is not None
    (
        start_col_abs,
        start_col_letter,
        start_row_abs,
        start_row_label,
        end_col_abs,
        end_col_letter,
        end_row_abs,
        end_row_label,
    ) = match.groups()

    start_col = column_letter_to_index(start_col_letter)
    end_col = column_letter_to_index(end_col_letter)
    start_row = int(start_row_label)
    end_row = int(end_row_label)

    if start_col > end_col or start_row > end_row:
        issues.append(
            MigrationIssue(
                code="reversed_range",
                cell=cell_address,
                reference=token,
                message=(
                    "the range is written in reverse order and cannot be "
                    "remapped while preserving order"
                ),
            )
        )
        return token, "reversed_range"

    columns = list(range(start_col, end_col + 1))
    if any(column not in mapping for column in columns):
        issues.append(
            MigrationIssue(
                code="unmapped_column",
                cell=cell_address,
                reference=token,
                message=(
                    "the range covers a column that is not part of the "
                    "column mapping"
                ),
            )
        )
        return token, "unmapped_column"

    mapped_columns = [mapping[column] for column in columns]
    if any(
        mapped_columns[index] >= mapped_columns[index + 1]
        for index in range(len(mapped_columns) - 1)
    ):
        issues.append(
            MigrationIssue(
                code="split_or_reordered_range",
                cell=cell_address,
                reference=token,
                message=(
                    "the mapped columns would no longer be contiguous and "
                    "in the same order"
                ),
            )
        )
        return token, "split_or_reordered_range"

    replacement = "{}:{}".format(
        _format_ref(
            bool(start_col_abs),
            mapped_columns[0],
            bool(start_row_abs),
            start_row,
        ),
        _format_ref(
            bool(end_col_abs),
            mapped_columns[-1],
            bool(end_row_abs),
            end_row,
        ),
    )
    return replacement, "rectangle"


# ---------------------------------------------------------------------------
# Raw cell reading
# ---------------------------------------------------------------------------


def _extract_grid(metadata: Mapping[str, Any], sheet_id: int) -> Mapping[str, Any]:
    for sheet in metadata.get("sheets", []):
        properties = sheet.get("properties", {})
        if properties.get("sheetId") == sheet_id:
            data = sheet.get("data", [])
            if not data:
                raise GSpreadException(
                    "no grid data returned for sheet id {}; the metadata "
                    "request must include grid data".format(sheet_id)
                )
            return data[0]
    raise GSpreadException("sheet id {} not found in metadata".format(sheet_id))


def _iter_cell_data(
    grid: Mapping[str, Any],
    row_start: int,
    row_end: int,
    col_start: int,
    col_end: int,
):
    """Yield ``(row, col, cell_data)`` from grid data, 1-based coordinates.

    The API keeps positional ``{}`` placeholders for empty rows and cells
    inside a requested range and only omits trailing empties, so enumeration
    order plus ``startRow`` / ``startColumn`` gives the coordinates.
    """
    grid_start_row = grid.get("startRow", 0)
    grid_start_col = grid.get("startColumn", 0)

    for relative_row, row_data in enumerate(grid.get("rowData", [])):
        absolute_row = grid_start_row + relative_row + 1
        if not (row_start <= absolute_row <= row_end):
            continue
        for relative_col, cell_data in enumerate(row_data.get("values", [])):
            absolute_col = grid_start_col + relative_col + 1
            if col_start <= absolute_col <= col_end:
                yield absolute_row, absolute_col, cell_data


def classify_cell(cell_data: Mapping[str, Any]) -> Tuple[str, Any]:
    """Classify one raw ``CellData`` object.

    :returns: ``(cell_type, value)`` where ``value`` is the value as stored in
        the source cell (formulas keep their leading ``=``).
    """
    entered = cell_data.get("userEnteredValue", {})
    if not entered:
        return CELL_TYPE_BLANK, ""

    if "formulaValue" in entered:
        return CELL_TYPE_FORMULA, entered["formulaValue"]
    if "stringValue" in entered:
        text = entered["stringValue"]
        if isinstance(text, str) and text.startswith("="):
            return CELL_TYPE_FORMULA_LIKE_TEXT, text
        return CELL_TYPE_VALUE, text
    if "numberValue" in entered:
        return CELL_TYPE_VALUE, entered["numberValue"]
    if "boolValue" in entered:
        return CELL_TYPE_VALUE, entered["boolValue"]
    if "errorValue" in entered:
        error = entered["errorValue"]
        return CELL_TYPE_ERROR, error.get("message", "#ERROR!")

    # Unknown userEnteredValue shape: treat as blank rather than dropping it.
    return CELL_TYPE_BLANK, ""


# ---------------------------------------------------------------------------
# Plan builder
# ---------------------------------------------------------------------------


def build_migration_preview(
    worksheet: "Worksheet",
    destination_title: str,
    column_mapping: Mapping[Any, Any],
    source_range: str,
) -> MigrationPlan:
    """Build a :class:`MigrationPlan` for one worksheet.

    This performs read-only API calls only; nothing is written.

    :param worksheet: the source :class:`gspread.worksheet.Worksheet`
    :param str destination_title: title of the destination worksheet in the
        same spreadsheet
    :param dict column_mapping: source column to destination column, letters
        (``"C"``) or 1-based indexes (``3``)
    :param str source_range: finite A1 range to migrate, e.g. ``"A1:F20"``
    :raises MigrationPlanError: if any formula cannot be remapped safely
    :raises InvalidInputValue: if the mapping or the source range is invalid
    """
    mapping = normalize_column_mapping(column_mapping)
    row_start, row_end, col_start, col_end = parse_finite_a1_range(source_range)

    # Lightweight call to check the destination sheet exists.
    titles_metadata = worksheet.spreadsheet.fetch_sheet_metadata(
        params={"fields": "sheets(properties(title))"}
    )
    sheet_titles = {
        sheet.get("properties", {}).get("title")
        for sheet in titles_metadata.get("sheets", [])
    }
    warnings: List[str] = []
    if destination_title not in sheet_titles:
        warnings.append(
            "destination worksheet {!r} does not exist yet; it must be "
            "created before applying the plan".format(destination_title)
        )

    metadata = worksheet.spreadsheet.fetch_sheet_metadata(
        params={
            "includeGridData": "true",
            "ranges": [
                "{}!{}".format(_quote_sheet_title(worksheet.title), source_range)
            ],
            "fields": (
                "sheets(properties(sheetId,title),"
                "data(startRow,startColumn,"
                "rowData(values(userEnteredValue))))"
            ),
        }
    )

    grid = _extract_grid(metadata, worksheet.id)

    table: Dict[Tuple[int, int], Mapping[str, Any]] = {}
    for row, col, cell_data in _iter_cell_data(
        grid, row_start, row_end, col_start, col_end
    ):
        table[(row, col)] = cell_data

    entries: List[CellEntry] = []
    all_issues: List[MigrationIssue] = []
    all_rewritten: List[RewrittenReference] = []

    for row in range(row_start, row_end + 1):
        for col in range(col_start, col_end + 1):
            if col not in mapping:
                continue
            cell_data = table.get((row, col), {})
            cell_type, value = classify_cell(cell_data)
            address = rowcol_to_a1(row, col)

            formula = None
            if cell_type == CELL_TYPE_FORMULA:
                formula, issues, rewritten = rewrite_formula(
                    value,
                    address,
                    mapping,
                    worksheet.title,
                    destination_title,
                )
                all_issues.extend(issues)
                all_rewritten.extend(rewritten)

            entries.append(
                CellEntry(
                    cell=address,
                    row=row,
                    col=col,
                    cell_type=cell_type,
                    value=formula if formula is not None else value,
                    formula=formula,
                )
            )

    if all_issues:
        raise MigrationPlanError(all_issues)

    plan = MigrationPlan(
        source_sheet=worksheet.title,
        source_range=source_range,
        destination_sheet=destination_title,
        column_mapping=dict(sorted(mapping.items())),
        row_start=row_start,
        row_end=row_end,
        cells=entries,
        batch_groups=[],
        rewritten_references=all_rewritten,
        warnings=warnings,
    )
    plan.batch_groups = _build_batch_groups(plan)
    return plan


_USER_ENTERED_TYPES = frozenset(
    {CELL_TYPE_FORMULA, CELL_TYPE_FORMULA_LIKE_TEXT, CELL_TYPE_ERROR}
)


def _entry_input_option(entry: CellEntry) -> str:
    if entry.cell_type in _USER_ENTERED_TYPES:
        return "USER_ENTERED"
    return "RAW"


def _encode_entry_value(entry: CellEntry) -> Any:
    # A leading apostrophe forces literal text under USER_ENTERED so a
    # formula-like string never becomes a formula.
    if entry.cell_type == CELL_TYPE_FORMULA_LIKE_TEXT:
        return "'" + entry.value
    return entry.value


def _split_segments(
    column_entries: List[CellEntry],
) -> List[Tuple[str, List[CellEntry]]]:
    """Group a column's rows into contiguous same-input-option segments."""
    segments: List[Tuple[str, List[CellEntry]]] = []
    for entry in column_entries:
        kind = _entry_input_option(entry)
        if segments and segments[-1][0] == kind:
            segments[-1][1].append(entry)
        else:
            segments.append((kind, [entry]))
    return segments


def _build_column_segments(
    source_col: int,
    column_entries: List[CellEntry],
    plan: MigrationPlan,
) -> List[Tuple[str, MutableMapping[str, Any]]]:
    column_entries.sort(key=lambda entry: entry.row)
    destination_letter = _column_letter(plan.column_mapping[source_col])
    prefix = _quote_sheet_title(plan.destination_sheet)
    payloads: List[Tuple[str, MutableMapping[str, Any]]] = []

    for kind, segment_entries in _split_segments(column_entries):
        values = [[_encode_entry_value(entry)] for entry in segment_entries]
        range_name = "{}!{}{}:{}{}".format(
            prefix,
            destination_letter,
            segment_entries[0].row,
            destination_letter,
            segment_entries[-1].row,
        )
        payloads.append((kind, {"range": range_name, "values": values}))
    return payloads


def _build_batch_groups(plan: MigrationPlan) -> List[BatchGroup]:
    """Split the plan cells into the two ``value_input_option`` groups.

    * ``USER_ENTERED``: formulas (so they are evaluated as formulas), literal
      text starting with ``=`` (prefixed with ``'`` so it stays text) and
      error cells (so their error value is reproduced rather than stored as
      plain text);
    * ``RAW``: every other plain value, stored without parsing.
    """
    by_column: Dict[int, List[CellEntry]] = {
        source_col: [] for source_col in plan.column_mapping
    }
    for entry in plan.cells:
        by_column[entry.col].append(entry)

    grouped: Dict[str, List[MutableMapping[str, Any]]] = {
        "USER_ENTERED": [],
        "RAW": [],
    }
    for source_col, column_entries in by_column.items():
        for kind, payload in _build_column_segments(source_col, column_entries, plan):
            grouped[kind].append(payload)

    groups: List[BatchGroup] = []
    if grouped["USER_ENTERED"]:
        groups.append(
            BatchGroup(
                value_input_option=ValueInputOption.user_entered,
                data=grouped["USER_ENTERED"],
            )
        )
    if grouped["RAW"]:
        groups.append(
            BatchGroup(
                value_input_option=ValueInputOption.raw,
                data=grouped["RAW"],
            )
        )
    return groups
