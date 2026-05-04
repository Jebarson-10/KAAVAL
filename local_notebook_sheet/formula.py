from __future__ import annotations

import ast
import operator
import re
from dataclasses import dataclass
from typing import Callable, Iterable


CellGetter = Callable[[int, int], str]

CELL_RE = re.compile(r"\b([A-Z]{1,3})([1-9][0-9]*)\b")
RANGE_RE = re.compile(r"\b([A-Z]{1,3}[1-9][0-9]*):([A-Z]{1,3}[1-9][0-9]*)\b")
FUNC_RE = re.compile(r"\b(SUM|AVERAGE|AVG|MIN|MAX|COUNT)\(([^()]*)\)", re.IGNORECASE)


def column_name(index: int) -> str:
    index += 1
    name = ""
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def column_index(name: str) -> int:
    total = 0
    for char in name.upper():
        if not ("A" <= char <= "Z"):
            raise ValueError(f"Invalid column name: {name}")
        total = total * 26 + (ord(char) - 64)
    return total - 1


def parse_cell(ref: str) -> tuple[int, int]:
    match = CELL_RE.fullmatch(ref.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {ref}")
    col, row = match.groups()
    return int(row) - 1, column_index(col)


def iter_range(start_ref: str, end_ref: str) -> Iterable[tuple[int, int]]:
    start_row, start_col = parse_cell(start_ref)
    end_row, end_col = parse_cell(end_ref)
    row_min, row_max = sorted((start_row, end_row))
    col_min, col_max = sorted((start_col, end_col))
    for row in range(row_min, row_max + 1):
        for col in range(col_min, col_max + 1):
            yield row, col


def as_number(value: object) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def format_result(value: float) -> str:
    if abs(value - round(value)) < 0.0000001:
        return str(int(round(value)))
    return f"{value:.8f}".rstrip("0").rstrip(".")


@dataclass
class FormulaEngine:
    getter: CellGetter

    def evaluate(self, raw: str, stack: set[tuple[int, int]] | None = None) -> str:
        if not raw.startswith("="):
            return raw
        stack = stack or set()
        expression = raw[1:].strip()
        if not expression:
            return ""
        try:
            value = self._evaluate_expression(expression, stack)
        except Exception:
            return "#ERR"
        return format_result(value)

    def _evaluate_cell(self, row: int, col: int, stack: set[tuple[int, int]]) -> float:
        key = (row, col)
        if key in stack:
            raise ValueError("Circular formula reference")
        raw = self.getter(row, col) or ""
        if raw.startswith("="):
            text = self.evaluate(raw, stack | {key})
            return as_number(text)
        return as_number(raw)

    def _range_values(self, start: str, end: str, stack: set[tuple[int, int]]) -> list[float]:
        return [self._evaluate_cell(row, col, stack) for row, col in iter_range(start, end)]

    def _evaluate_function(self, name: str, argument_text: str, stack: set[tuple[int, int]]) -> float:
        values: list[float] = []
        for argument in _split_args(argument_text):
            argument = argument.strip()
            if not argument:
                continue
            range_match = RANGE_RE.fullmatch(argument.upper())
            cell_match = CELL_RE.fullmatch(argument.upper())
            if range_match:
                values.extend(self._range_values(range_match.group(1), range_match.group(2), stack))
            elif cell_match:
                row, col = parse_cell(argument)
                values.append(self._evaluate_cell(row, col, stack))
            else:
                values.append(self._safe_eval(self._replace_cells(argument, stack)))

        upper = name.upper()
        if upper == "SUM":
            return sum(values)
        if upper in {"AVERAGE", "AVG"}:
            return sum(values) / len(values) if values else 0.0
        if upper == "MIN":
            return min(values) if values else 0.0
        if upper == "MAX":
            return max(values) if values else 0.0
        if upper == "COUNT":
            return float(len([value for value in values if value != 0]))
        raise ValueError(f"Unsupported function: {name}")

    def _evaluate_expression(self, expression: str, stack: set[tuple[int, int]]) -> float:
        working = expression
        while True:
            match = FUNC_RE.search(working)
            if not match:
                break
            result = self._evaluate_function(match.group(1), match.group(2), stack)
            working = working[: match.start()] + format_result(result) + working[match.end() :]
        working = self._replace_cells(working, stack)
        if RANGE_RE.search(working):
            raise ValueError("Ranges must be inside a supported function")
        return self._safe_eval(working)

    def _replace_cells(self, expression: str, stack: set[tuple[int, int]]) -> str:
        def replace(match: re.Match[str]) -> str:
            row, col = parse_cell(match.group(0))
            return format_result(self._evaluate_cell(row, col, stack))

        return CELL_RE.sub(replace, expression.upper())

    def _safe_eval(self, expression: str) -> float:
        tree = ast.parse(expression, mode="eval")
        return float(_eval_ast(tree.body))


def _split_args(text: str) -> list[str]:
    args: list[str] = []
    current: list[str] = []
    depth = 0
    for char in text:
        if char == "," and depth == 0:
            args.append("".join(current))
            current = []
            continue
        if char == "(":
            depth += 1
        elif char == ")":
            depth = max(0, depth - 1)
        current.append(char)
    args.append("".join(current))
    return args


_BINARY_OPS = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
    ast.Pow: operator.pow,
    ast.Mod: operator.mod,
}

_UNARY_OPS = {
    ast.UAdd: operator.pos,
    ast.USub: operator.neg,
}


def _eval_ast(node: ast.AST) -> float:
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
        return float(node.value)
    if isinstance(node, ast.BinOp) and type(node.op) in _BINARY_OPS:
        return _BINARY_OPS[type(node.op)](_eval_ast(node.left), _eval_ast(node.right))
    if isinstance(node, ast.UnaryOp) and type(node.op) in _UNARY_OPS:
        return _UNARY_OPS[type(node.op)](_eval_ast(node.operand))
    raise ValueError("Unsupported formula expression")
