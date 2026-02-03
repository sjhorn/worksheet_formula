import 'dart:math' as math;

import '../ast/nodes.dart';
import '../evaluation/context.dart';
import '../evaluation/errors.dart';
import '../evaluation/value.dart';
import 'function.dart';
import 'registry.dart';

/// Register all dynamic array functions.
void registerArrayFunctions(FunctionRegistry registry) {
  registry.registerAll([
    // Wave A — Generators
    SequenceFunction(),
    RandArrayFunction(),
    // Wave B — Flatten/Reshape
    ToColFunction(),
    ToRowFunction(),
    WrapRowsFunction(),
    WrapColsFunction(),
    // Wave C — Slice/Select
    ChooseRowsFunction(),
    ChooseColsFunction(),
    DropFunction(),
    TakeFunction(),
    ExpandFunction(),
    // Wave D — Concatenation
    HStackFunction(),
    VStackFunction(),
    // Wave E — Filter/Sort/Unique
    FilterFunction(),
    UniqueFunction(),
    SortFunction(),
    SortByFunction(),
    MunitFunction(),
    MmultFunction(),
    MdetermFunction(),
    MinverseFunction(),
  ]);
}

// ─── Shared Helpers ─────────────────────────────────────────

/// Convert any FormulaValue to a 2D matrix (scalars become 1x1).
List<List<FormulaValue>> _toMatrix(FormulaValue value) {
  if (value is RangeValue) return value.values;
  return [
    [value],
  ];
}

/// Flatten a 2D matrix to a 1D list.
List<FormulaValue> _flatten(
  List<List<FormulaValue>> matrix, {
  bool byColumn = false,
}) {
  if (byColumn) {
    final result = <FormulaValue>[];
    if (matrix.isEmpty) return result;
    final rows = matrix.length;
    final cols = matrix[0].length;
    for (var c = 0; c < cols; c++) {
      for (var r = 0; r < rows; r++) {
        result.add(matrix[r][c]);
      }
    }
    return result;
  }
  return matrix.expand((row) => row).toList();
}

/// Check if two rows are equal (case-insensitive for text).
bool _rowsEqual(List<FormulaValue> a, List<FormulaValue> b) {
  if (a.length != b.length) return false;
  for (var i = 0; i < a.length; i++) {
    if (!_valuesEqual(a[i], b[i])) return false;
  }
  return true;
}

/// Check if two values are equal (case-insensitive for text).
bool _valuesEqual(FormulaValue a, FormulaValue b) {
  if (a is TextValue && b is TextValue) {
    return a.value.toLowerCase() == b.value.toLowerCase();
  }
  if (a is NumberValue && b is NumberValue) return a.value == b.value;
  if (a is BooleanValue && b is BooleanValue) return a.value == b.value;
  if (a is ErrorValue && b is ErrorValue) return a.error == b.error;
  if (a is EmptyValue && b is EmptyValue) return true;
  return false;
}

/// Compare two values for sorting: numbers < text < booleans < errors.
int _compareValues(FormulaValue a, FormulaValue b) {
  final orderA = _typeOrder(a);
  final orderB = _typeOrder(b);
  if (orderA != orderB) return orderA.compareTo(orderB);

  if (a is NumberValue && b is NumberValue) {
    return a.value.compareTo(b.value);
  }
  if (a is TextValue && b is TextValue) {
    return a.value.toLowerCase().compareTo(b.value.toLowerCase());
  }
  if (a is BooleanValue && b is BooleanValue) {
    return (a.value ? 1 : 0).compareTo(b.value ? 1 : 0);
  }
  return 0;
}

int _typeOrder(FormulaValue v) {
  if (v is EmptyValue) return 0;
  if (v is NumberValue) return 1;
  if (v is TextValue) return 2;
  if (v is BooleanValue) return 3;
  if (v is ErrorValue) return 4;
  return 5;
}

/// Resolve a 1-based index (negative = from end). Returns null if invalid.
int? _resolveIndex(int index, int length) {
  if (index == 0) return null;
  if (index > 0) {
    if (index > length) return null;
    return index - 1;
  }
  // negative
  final resolved = length + index;
  if (resolved < 0) return null;
  return resolved;
}

// ─── Wave A — Generators ────────────────────────────────────

/// SEQUENCE(rows, [cols=1], [start=1], [step=1])
class SequenceFunction extends FormulaFunction {
  @override
  String get name => 'SEQUENCE';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    for (final v in values) {
      if (v.isError) return v;
    }

    final rows = values[0].toNumber()?.toInt();
    if (rows == null) return const FormulaValue.error(FormulaError.value);
    if (rows < 1) return const FormulaValue.error(FormulaError.value);

    final cols =
        values.length > 1 ? values[1].toNumber()?.toInt() ?? 1 : 1;
    if (cols < 1) return const FormulaValue.error(FormulaError.value);

    final start = values.length > 2 ? values[2].toNumber() ?? 1 : 1;
    final step = values.length > 3 ? values[3].toNumber() ?? 1 : 1;

    final result = <List<FormulaValue>>[];
    var current = start;
    for (var r = 0; r < rows; r++) {
      final row = <FormulaValue>[];
      for (var c = 0; c < cols; c++) {
        row.add(FormulaValue.number(current));
        current += step;
      }
      result.add(row);
    }
    return RangeValue(result);
  }
}

/// RANDARRAY([rows=1], [cols=1], [min=0], [max=1], [whole=FALSE])
class RandArrayFunction extends FormulaFunction {
  @override
  String get name => 'RANDARRAY';
  @override
  int get minArgs => 0;
  @override
  int get maxArgs => 5;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    for (final v in values) {
      if (v.isError) return v;
    }

    final rows =
        values.isNotEmpty ? values[0].toNumber()?.toInt() ?? 1 : 1;
    if (rows < 1) return const FormulaValue.error(FormulaError.value);

    final cols = values.length > 1 ? values[1].toNumber()?.toInt() ?? 1 : 1;
    if (cols < 1) return const FormulaValue.error(FormulaError.value);

    final minVal = values.length > 2 ? values[2].toNumber() ?? 0 : 0;
    final maxVal = values.length > 3 ? values[3].toNumber() ?? 1 : 1;
    if (minVal > maxVal) {
      return const FormulaValue.error(FormulaError.value);
    }

    final whole = values.length > 4 ? values[4].toBool() : false;
    final rng = math.Random();

    final result = <List<FormulaValue>>[];
    for (var r = 0; r < rows; r++) {
      final row = <FormulaValue>[];
      for (var c = 0; c < cols; c++) {
        if (whole) {
          final min = minVal.toInt();
          final max = maxVal.toInt();
          row.add(FormulaValue.number(min + rng.nextInt(max - min + 1)));
        } else {
          row.add(
            FormulaValue.number(minVal + rng.nextDouble() * (maxVal - minVal)),
          );
        }
      }
      result.add(row);
    }
    return RangeValue(result);
  }
}

// ─── Wave B — Flatten/Reshape ───────────────────────────────

/// TOCOL(array, [ignore=0], [scan_by_col=FALSE])
class ToColFunction extends FormulaFunction {
  @override
  String get name => 'TOCOL';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);

    final matrix = _toMatrix(values[0]);
    final ignore = values.length > 1 ? values[1].toNumber()?.toInt() ?? 0 : 0;
    final byCol = values.length > 2 ? values[2].toBool() : false;

    var flat = _flatten(matrix, byColumn: byCol);
    flat = _filterValues(flat, ignore);

    if (flat.isEmpty) return const FormulaValue.error(FormulaError.calc);

    return RangeValue(flat.map((v) => [v]).toList());
  }
}

/// TOROW(array, [ignore=0], [scan_by_col=FALSE])
class ToRowFunction extends FormulaFunction {
  @override
  String get name => 'TOROW';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);

    final matrix = _toMatrix(values[0]);
    final ignore = values.length > 1 ? values[1].toNumber()?.toInt() ?? 0 : 0;
    final byCol = values.length > 2 ? values[2].toBool() : false;

    var flat = _flatten(matrix, byColumn: byCol);
    flat = _filterValues(flat, ignore);

    if (flat.isEmpty) return const FormulaValue.error(FormulaError.calc);

    return RangeValue([flat]);
  }
}

/// Filter values based on ignore parameter: 0=none, 1=blanks, 2=errors, 3=both.
List<FormulaValue> _filterValues(List<FormulaValue> values, int ignore) {
  if (ignore == 0) return values;
  return values.where((v) {
    if (ignore == 1 || ignore == 3) {
      if (v is EmptyValue) return false;
    }
    if (ignore == 2 || ignore == 3) {
      if (v is ErrorValue) return false;
    }
    return true;
  }).toList();
}

/// WRAPROWS(vector, wrap_count, [pad=#N/A])
class WrapRowsFunction extends FormulaFunction {
  @override
  String get name => 'WRAPROWS';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);

    final matrix = _toMatrix(values[0]);
    final wrapCount = values[1].toNumber()?.toInt();
    if (wrapCount == null || wrapCount < 1) {
      return const FormulaValue.error(FormulaError.value);
    }
    final pad = values.length > 2
        ? values[2]
        : const FormulaValue.error(FormulaError.na);

    final flat = _flatten(matrix);
    final result = <List<FormulaValue>>[];

    for (var i = 0; i < flat.length; i += wrapCount) {
      final row = <FormulaValue>[];
      for (var j = 0; j < wrapCount; j++) {
        final idx = i + j;
        row.add(idx < flat.length ? flat[idx] : pad);
      }
      result.add(row);
    }

    return RangeValue(result);
  }
}

/// WRAPCOLS(vector, wrap_count, [pad=#N/A])
class WrapColsFunction extends FormulaFunction {
  @override
  String get name => 'WRAPCOLS';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);

    final matrix = _toMatrix(values[0]);
    final wrapCount = values[1].toNumber()?.toInt();
    if (wrapCount == null || wrapCount < 1) {
      return const FormulaValue.error(FormulaError.value);
    }
    final pad = values.length > 2
        ? values[2]
        : const FormulaValue.error(FormulaError.na);

    final flat = _flatten(matrix);
    final numCols = (flat.length / wrapCount).ceil();

    final result = <List<FormulaValue>>[];
    for (var r = 0; r < wrapCount; r++) {
      final row = <FormulaValue>[];
      for (var c = 0; c < numCols; c++) {
        final idx = c * wrapCount + r;
        row.add(idx < flat.length ? flat[idx] : pad);
      }
      result.add(row);
    }

    return RangeValue(result);
  }
}

// ─── Wave C — Slice/Select ──────────────────────────────────

/// CHOOSEROWS(array, row_num1, ...)
class ChooseRowsFunction extends FormulaFunction {
  @override
  String get name => 'CHOOSEROWS';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arrayVal = args[0].evaluate(context);
    if (arrayVal.isError) return arrayVal;
    final matrix = _toMatrix(arrayVal);
    final rowCount = matrix.length;

    final result = <List<FormulaValue>>[];
    for (var i = 1; i < args.length; i++) {
      final idxVal = args[i].evaluate(context);
      if (idxVal.isError) return idxVal;
      final idx = idxVal.toNumber()?.toInt();
      if (idx == null) return const FormulaValue.error(FormulaError.value);
      final resolved = _resolveIndex(idx, rowCount);
      if (resolved == null) {
        return const FormulaValue.error(FormulaError.value);
      }
      result.add(List.of(matrix[resolved]));
    }

    return RangeValue(result);
  }
}

/// CHOOSECOLS(array, col_num1, ...)
class ChooseColsFunction extends FormulaFunction {
  @override
  String get name => 'CHOOSECOLS';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arrayVal = args[0].evaluate(context);
    if (arrayVal.isError) return arrayVal;
    final matrix = _toMatrix(arrayVal);
    final colCount = matrix.isEmpty ? 0 : matrix[0].length;

    final indices = <int>[];
    for (var i = 1; i < args.length; i++) {
      final idxVal = args[i].evaluate(context);
      if (idxVal.isError) return idxVal;
      final idx = idxVal.toNumber()?.toInt();
      if (idx == null) return const FormulaValue.error(FormulaError.value);
      final resolved = _resolveIndex(idx, colCount);
      if (resolved == null) {
        return const FormulaValue.error(FormulaError.value);
      }
      indices.add(resolved);
    }

    final result = <List<FormulaValue>>[];
    for (final row in matrix) {
      result.add(indices.map((c) => row[c]).toList());
    }

    return RangeValue(result);
  }
}

/// DROP(array, rows, [cols=0])
class DropFunction extends FormulaFunction {
  @override
  String get name => 'DROP';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arrayVal = args[0].evaluate(context);
    if (arrayVal.isError) return arrayVal;
    final matrix = _toMatrix(arrayVal);
    final rowCount = matrix.length;
    final colCount = matrix.isEmpty ? 0 : matrix[0].length;

    final dropRowsVal = args[1].evaluate(context);
    if (dropRowsVal.isError) return dropRowsVal;
    final dropRows = dropRowsVal.toNumber()?.toInt() ?? 0;

    final dropCols = args.length > 2
        ? (args[2].evaluate(context).toNumber()?.toInt() ?? 0)
        : 0;

    // Calculate remaining rows
    int startRow, endRow;
    if (dropRows >= 0) {
      startRow = dropRows;
      endRow = rowCount;
    } else {
      startRow = 0;
      endRow = rowCount + dropRows;
    }

    // Calculate remaining columns
    int startCol, endCol;
    if (dropCols >= 0) {
      startCol = dropCols;
      endCol = colCount;
    } else {
      startCol = 0;
      endCol = colCount + dropCols;
    }

    if (startRow >= endRow || startCol >= endCol) {
      return const FormulaValue.error(FormulaError.calc);
    }

    final result = <List<FormulaValue>>[];
    for (var r = startRow; r < endRow; r++) {
      result.add(matrix[r].sublist(startCol, endCol));
    }

    return RangeValue(result);
  }
}

/// TAKE(array, rows, [cols=0])
class TakeFunction extends FormulaFunction {
  @override
  String get name => 'TAKE';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arrayVal = args[0].evaluate(context);
    if (arrayVal.isError) return arrayVal;
    final matrix = _toMatrix(arrayVal);
    final rowCount = matrix.length;
    final colCount = matrix.isEmpty ? 0 : matrix[0].length;

    final takeRowsVal = args[1].evaluate(context);
    if (takeRowsVal.isError) return takeRowsVal;
    final takeRows = takeRowsVal.toNumber()?.toInt() ?? 0;
    if (takeRows == 0) return const FormulaValue.error(FormulaError.calc);

    final takeCols = args.length > 2
        ? (args[2].evaluate(context).toNumber()?.toInt() ?? 0)
        : 0;

    // Calculate row range
    int startRow, endRow;
    if (takeRows > 0) {
      startRow = 0;
      endRow = math.min(takeRows, rowCount);
    } else {
      startRow = math.max(0, rowCount + takeRows);
      endRow = rowCount;
    }

    // Calculate column range: 0 means take all columns
    int startCol, endCol;
    if (takeCols == 0) {
      startCol = 0;
      endCol = colCount;
    } else if (takeCols > 0) {
      startCol = 0;
      endCol = math.min(takeCols, colCount);
    } else {
      startCol = math.max(0, colCount + takeCols);
      endCol = colCount;
    }

    final result = <List<FormulaValue>>[];
    for (var r = startRow; r < endRow; r++) {
      result.add(matrix[r].sublist(startCol, endCol));
    }

    return RangeValue(result);
  }
}

/// EXPAND(array, rows, [cols=current], [pad=#N/A])
class ExpandFunction extends FormulaFunction {
  @override
  String get name => 'EXPAND';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arrayVal = args[0].evaluate(context);
    if (arrayVal.isError) return arrayVal;
    final matrix = _toMatrix(arrayVal);
    final curRows = matrix.length;
    final curCols = matrix.isEmpty ? 0 : matrix[0].length;

    final newRowsVal = args[1].evaluate(context);
    if (newRowsVal.isError) return newRowsVal;
    final newRows = newRowsVal.toNumber()?.toInt() ?? curRows;

    final newCols = args.length > 2
        ? (args[2].evaluate(context).toNumber()?.toInt() ?? curCols)
        : curCols;

    final pad = args.length > 3
        ? args[3].evaluate(context)
        : const FormulaValue.error(FormulaError.na);

    if (newRows < curRows || newCols < curCols) {
      return const FormulaValue.error(FormulaError.value);
    }

    final result = <List<FormulaValue>>[];
    for (var r = 0; r < newRows; r++) {
      final row = <FormulaValue>[];
      for (var c = 0; c < newCols; c++) {
        if (r < curRows && c < curCols) {
          row.add(matrix[r][c]);
        } else {
          row.add(pad);
        }
      }
      result.add(row);
    }

    return RangeValue(result);
  }
}

// ─── Wave D — Concatenation ─────────────────────────────────

/// HSTACK(array1, ...)
class HStackFunction extends FormulaFunction {
  @override
  String get name => 'HSTACK';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final matrices = <List<List<FormulaValue>>>[];
    for (final arg in args) {
      final val = arg.evaluate(context);
      if (val.isError) return val;
      matrices.add(_toMatrix(val));
    }

    final maxRows =
        matrices.map((m) => m.length).reduce((a, b) => math.max(a, b));

    final result = <List<FormulaValue>>[];
    for (var r = 0; r < maxRows; r++) {
      final row = <FormulaValue>[];
      for (final m in matrices) {
        if (r < m.length) {
          row.addAll(m[r]);
        } else {
          // Pad with #N/A for the width of this matrix
          row.addAll(List.filled(
            m[0].length,
            const FormulaValue.error(FormulaError.na),
          ));
        }
      }
      result.add(row);
    }

    return RangeValue(result);
  }
}

/// VSTACK(array1, ...)
class VStackFunction extends FormulaFunction {
  @override
  String get name => 'VSTACK';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final matrices = <List<List<FormulaValue>>>[];
    for (final arg in args) {
      final val = arg.evaluate(context);
      if (val.isError) return val;
      matrices.add(_toMatrix(val));
    }

    final maxCols = matrices
        .map((m) => m.isEmpty ? 0 : m[0].length)
        .reduce((a, b) => math.max(a, b));

    final result = <List<FormulaValue>>[];
    for (final m in matrices) {
      for (final row in m) {
        if (row.length < maxCols) {
          result.add([
            ...row,
            ...List.filled(
              maxCols - row.length,
              const FormulaValue.error(FormulaError.na),
            ),
          ]);
        } else {
          result.add(row);
        }
      }
    }

    return RangeValue(result);
  }
}

// ─── Wave E — Filter/Sort/Unique ────────────────────────────

/// FILTER(array, include, [if_empty])
class FilterFunction extends FormulaFunction {
  @override
  String get name => 'FILTER';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arrayVal = args[0].evaluate(context);
    if (arrayVal.isError) return arrayVal;
    final matrix = _toMatrix(arrayVal);
    final rowCount = matrix.length;
    final colCount = matrix.isEmpty ? 0 : matrix[0].length;

    final includeVal = args[1].evaluate(context);
    if (includeVal.isError) return includeVal;
    final includeMatrix = _toMatrix(includeVal);
    final incRows = includeMatrix.length;
    final incCols = includeMatrix.isEmpty ? 0 : includeMatrix[0].length;

    // Determine if filtering rows or columns
    if (incRows == rowCount && incCols == 1) {
      // Row filter: include is a column vector matching array rows
      final result = <List<FormulaValue>>[];
      for (var r = 0; r < rowCount; r++) {
        if (includeMatrix[r][0].isTruthy) {
          result.add(List.of(matrix[r]));
        }
      }
      if (result.isEmpty) {
        if (args.length > 2) return args[2].evaluate(context);
        return const FormulaValue.error(FormulaError.calc);
      }
      return RangeValue(result);
    } else if (incRows == 1 && incCols == colCount) {
      // Column filter: include is a row vector matching array columns
      final keepCols = <int>[];
      for (var c = 0; c < colCount; c++) {
        if (includeMatrix[0][c].isTruthy) {
          keepCols.add(c);
        }
      }
      if (keepCols.isEmpty) {
        if (args.length > 2) return args[2].evaluate(context);
        return const FormulaValue.error(FormulaError.calc);
      }
      final result = <List<FormulaValue>>[];
      for (final row in matrix) {
        result.add(keepCols.map((c) => row[c]).toList());
      }
      return RangeValue(result);
    } else {
      return const FormulaValue.error(FormulaError.value);
    }
  }
}

/// UNIQUE(array, [by_col=FALSE], [exactly_once=FALSE])
class UniqueFunction extends FormulaFunction {
  @override
  String get name => 'UNIQUE';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arrayVal = args[0].evaluate(context);
    if (arrayVal.isError) return arrayVal;
    final matrix = _toMatrix(arrayVal);
    final byCol =
        args.length > 1 ? args[1].evaluate(context).toBool() : false;
    final exactlyOnce =
        args.length > 2 ? args[2].evaluate(context).toBool() : false;

    if (byCol) {
      return _uniqueColumns(matrix, exactlyOnce);
    } else {
      return _uniqueRows(matrix, exactlyOnce);
    }
  }

  FormulaValue _uniqueRows(
    List<List<FormulaValue>> matrix,
    bool exactlyOnce,
  ) {
    if (exactlyOnce) {
      // Keep only rows that appear exactly once
      final result = <List<FormulaValue>>[];
      for (var i = 0; i < matrix.length; i++) {
        var count = 0;
        for (var j = 0; j < matrix.length; j++) {
          if (_rowsEqual(matrix[i], matrix[j])) count++;
        }
        if (count == 1) result.add(matrix[i]);
      }
      if (result.isEmpty) {
        return const FormulaValue.error(FormulaError.calc);
      }
      return RangeValue(result);
    }

    // Keep first occurrence of each unique row
    final result = <List<FormulaValue>>[];
    for (final row in matrix) {
      var found = false;
      for (final existing in result) {
        if (_rowsEqual(row, existing)) {
          found = true;
          break;
        }
      }
      if (!found) result.add(row);
    }
    return RangeValue(result);
  }

  FormulaValue _uniqueColumns(
    List<List<FormulaValue>> matrix,
    bool exactlyOnce,
  ) {
    if (matrix.isEmpty) return RangeValue(matrix);
    final colCount = matrix[0].length;
    final rowCount = matrix.length;

    // Extract columns as lists
    List<FormulaValue> getCol(int c) =>
        [for (var r = 0; r < rowCount; r++) matrix[r][c]];

    if (exactlyOnce) {
      final keepCols = <int>[];
      for (var i = 0; i < colCount; i++) {
        var count = 0;
        final colI = getCol(i);
        for (var j = 0; j < colCount; j++) {
          if (_rowsEqual(colI, getCol(j))) count++;
        }
        if (count == 1) keepCols.add(i);
      }
      if (keepCols.isEmpty) {
        return const FormulaValue.error(FormulaError.calc);
      }
      return RangeValue([
        for (var r = 0; r < rowCount; r++)
          [for (final c in keepCols) matrix[r][c]],
      ]);
    }

    // Keep first occurrence of each unique column
    final keepCols = <int>[];
    for (var i = 0; i < colCount; i++) {
      final colI = getCol(i);
      var found = false;
      for (final existing in keepCols) {
        if (_rowsEqual(colI, getCol(existing))) {
          found = true;
          break;
        }
      }
      if (!found) keepCols.add(i);
    }

    return RangeValue([
      for (var r = 0; r < rowCount; r++)
        [for (final c in keepCols) matrix[r][c]],
    ]);
  }
}

/// SORT(array, [sort_index=1], [sort_order=1], [by_col=FALSE])
class SortFunction extends FormulaFunction {
  @override
  String get name => 'SORT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arrayVal = args[0].evaluate(context);
    if (arrayVal.isError) return arrayVal;
    final matrix = _toMatrix(arrayVal);

    final sortIndex =
        args.length > 1 ? args[1].evaluate(context).toNumber()?.toInt() ?? 1 : 1;
    final sortOrder =
        args.length > 2 ? args[2].evaluate(context).toNumber()?.toInt() ?? 1 : 1;
    final byCol =
        args.length > 3 ? args[3].evaluate(context).toBool() : false;

    if (byCol) {
      return _sortByCol(matrix, sortIndex, sortOrder);
    } else {
      return _sortByRow(matrix, sortIndex, sortOrder);
    }
  }

  FormulaValue _sortByRow(
    List<List<FormulaValue>> matrix,
    int sortIndex,
    int sortOrder,
  ) {
    if (matrix.isEmpty) return RangeValue(matrix);
    final colCount = matrix[0].length;
    if (sortIndex < 1 || sortIndex > colCount) {
      return const FormulaValue.error(FormulaError.value);
    }
    final colIdx = sortIndex - 1;

    final sorted = List.of(matrix);
    sorted.sort((a, b) {
      final cmp = _compareValues(a[colIdx], b[colIdx]);
      return sortOrder >= 0 ? cmp : -cmp;
    });

    return RangeValue(sorted);
  }

  FormulaValue _sortByCol(
    List<List<FormulaValue>> matrix,
    int sortIndex,
    int sortOrder,
  ) {
    if (matrix.isEmpty) return RangeValue(matrix);
    final rowCount = matrix.length;
    final colCount = matrix[0].length;
    if (sortIndex < 1 || sortIndex > rowCount) {
      return const FormulaValue.error(FormulaError.value);
    }
    final rowIdx = sortIndex - 1;

    // Create column index list and sort by the specified row
    final colIndices = List.generate(colCount, (i) => i);
    colIndices.sort((a, b) {
      final cmp = _compareValues(matrix[rowIdx][a], matrix[rowIdx][b]);
      return sortOrder >= 0 ? cmp : -cmp;
    });

    // Rebuild matrix with sorted columns
    return RangeValue([
      for (var r = 0; r < rowCount; r++)
        [for (final c in colIndices) matrix[r][c]],
    ]);
  }
}

/// SORTBY(array, by_array1, [sort_order1], ...)
class SortByFunction extends FormulaFunction {
  @override
  String get name => 'SORTBY';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arrayVal = args[0].evaluate(context);
    if (arrayVal.isError) return arrayVal;
    final matrix = _toMatrix(arrayVal);
    final rowCount = matrix.length;

    // Parse (by_array, sort_order) pairs from remaining args
    final keys = <List<List<FormulaValue>>>[];
    final orders = <int>[];

    var i = 1;
    while (i < args.length) {
      final byVal = args[i].evaluate(context);
      if (byVal.isError) return byVal;
      final byMatrix = _toMatrix(byVal);

      if (byMatrix.length != rowCount) {
        return const FormulaValue.error(FormulaError.value);
      }

      keys.add(byMatrix);
      i++;

      if (i < args.length) {
        // Check if next arg is a sort_order (number) or another by_array
        final nextVal = args[i].evaluate(context);
        if (nextVal.isError) return nextVal;
        final nextNum = nextVal.toNumber()?.toInt();
        if (nextNum != null && (nextNum == 1 || nextNum == -1)) {
          orders.add(nextNum);
          i++;
        } else {
          orders.add(1);
        }
      } else {
        orders.add(1);
      }
    }

    // Create index list and perform stable multi-key sort
    final indices = List.generate(rowCount, (i) => i);
    indices.sort((a, b) {
      for (var k = 0; k < keys.length; k++) {
        // Use first column of each by_array for comparison
        final valA = keys[k][a][0];
        final valB = keys[k][b][0];
        final cmp = _compareValues(valA, valB);
        if (cmp != 0) return orders[k] > 0 ? cmp : -cmp;
      }
      return 0;
    });

    return RangeValue([
      for (final idx in indices) List.of(matrix[idx]),
    ]);
  }
}

// ─── Matrix Functions ───────────────────────────────────────

/// MUNIT(dimension) - Returns the identity matrix of size N×N.
class MunitFunction extends FormulaFunction {
  @override
  String get name => 'MUNIT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber()?.toInt();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    if (n < 1) return const FormulaValue.error(FormulaError.value);

    final result = <List<FormulaValue>>[];
    for (var r = 0; r < n; r++) {
      final row = <FormulaValue>[];
      for (var c = 0; c < n; c++) {
        row.add(FormulaValue.number(r == c ? 1 : 0));
      }
      result.add(row);
    }
    return RangeValue(result);
  }
}

/// MMULT(array1, array2) - Matrix multiplication.
class MmultFunction extends FormulaFunction {
  @override
  String get name => 'MMULT';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final val1 = args[0].evaluate(context);
    final val2 = args[1].evaluate(context);
    if (val1.isError) return val1;
    if (val2.isError) return val2;

    final a = _toMatrix(val1);
    final b = _toMatrix(val2);

    final rowsA = a.length;
    final colsA = a.isEmpty ? 0 : a[0].length;
    final rowsB = b.length;
    final colsB = b.isEmpty ? 0 : b[0].length;

    if (colsA != rowsB) {
      return const FormulaValue.error(FormulaError.value);
    }

    // Convert to numeric matrices
    final numA = <List<double>>[];
    for (final row in a) {
      final numRow = <double>[];
      for (final cell in row) {
        final n = cell.toNumber()?.toDouble();
        if (n == null) return const FormulaValue.error(FormulaError.value);
        numRow.add(n);
      }
      numA.add(numRow);
    }

    final numB = <List<double>>[];
    for (final row in b) {
      final numRow = <double>[];
      for (final cell in row) {
        final n = cell.toNumber()?.toDouble();
        if (n == null) return const FormulaValue.error(FormulaError.value);
        numRow.add(n);
      }
      numB.add(numRow);
    }

    final result = <List<FormulaValue>>[];
    for (var r = 0; r < rowsA; r++) {
      final row = <FormulaValue>[];
      for (var c = 0; c < colsB; c++) {
        var sum = 0.0;
        for (var k = 0; k < colsA; k++) {
          sum += numA[r][k] * numB[k][c];
        }
        row.add(FormulaValue.number(sum));
      }
      result.add(row);
    }
    return RangeValue(result);
  }
}

/// Convert a FormulaValue matrix to a list of lists of doubles.
/// Returns null if any value is non-numeric.
List<List<double>>? _toNumericMatrix(List<List<FormulaValue>> matrix) {
  final result = <List<double>>[];
  for (final row in matrix) {
    final numRow = <double>[];
    for (final cell in row) {
      final n = cell.toNumber()?.toDouble();
      if (n == null) return null;
      numRow.add(n);
    }
    result.add(numRow);
  }
  return result;
}

/// MDETERM(array) - Returns the determinant of a square matrix.
/// Uses LU decomposition with partial pivoting.
class MdetermFunction extends FormulaFunction {
  @override
  String get name => 'MDETERM';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final val = args[0].evaluate(context);
    if (val.isError) return val;

    final matrix = _toMatrix(val);
    final rows = matrix.length;
    final cols = matrix.isEmpty ? 0 : matrix[0].length;
    if (rows != cols) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (rows == 0) return const FormulaValue.error(FormulaError.value);

    final numMatrix = _toNumericMatrix(matrix);
    if (numMatrix == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    final n = rows;

    // Make a mutable copy for LU decomposition
    final lu = [for (final row in numMatrix) List<double>.from(row)];

    var det = 1.0;
    for (var i = 0; i < n; i++) {
      // Partial pivoting
      var maxVal = lu[i][i].abs();
      var maxRow = i;
      for (var k = i + 1; k < n; k++) {
        if (lu[k][i].abs() > maxVal) {
          maxVal = lu[k][i].abs();
          maxRow = k;
        }
      }

      if (maxRow != i) {
        final tmp = lu[i];
        lu[i] = lu[maxRow];
        lu[maxRow] = tmp;
        det = -det; // Row swap flips sign
      }

      if (lu[i][i] == 0) return const FormulaValue.number(0);

      det *= lu[i][i];

      for (var k = i + 1; k < n; k++) {
        lu[k][i] /= lu[i][i];
        for (var j = i + 1; j < n; j++) {
          lu[k][j] -= lu[k][i] * lu[i][j];
        }
      }
    }

    // Round to avoid floating point noise for integer results
    if ((det - det.roundToDouble()).abs() < 1e-10) {
      det = det.roundToDouble();
    }

    return FormulaValue.number(det);
  }
}

/// MINVERSE(array) - Returns the inverse of a square matrix.
/// Uses Gauss-Jordan elimination.
class MinverseFunction extends FormulaFunction {
  @override
  String get name => 'MINVERSE';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final val = args[0].evaluate(context);
    if (val.isError) return val;

    final matrix = _toMatrix(val);
    final rows = matrix.length;
    final cols = matrix.isEmpty ? 0 : matrix[0].length;
    if (rows != cols) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (rows == 0) return const FormulaValue.error(FormulaError.value);

    final numMatrix = _toNumericMatrix(matrix);
    if (numMatrix == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    final n = rows;

    // Augment with identity matrix
    final aug = <List<double>>[];
    for (var r = 0; r < n; r++) {
      final row = List<double>.from(numMatrix[r]);
      for (var c = 0; c < n; c++) {
        row.add(r == c ? 1.0 : 0.0);
      }
      aug.add(row);
    }

    // Gauss-Jordan elimination with partial pivoting
    for (var i = 0; i < n; i++) {
      // Find pivot
      var maxVal = aug[i][i].abs();
      var maxRow = i;
      for (var k = i + 1; k < n; k++) {
        if (aug[k][i].abs() > maxVal) {
          maxVal = aug[k][i].abs();
          maxRow = k;
        }
      }

      if (maxRow != i) {
        final tmp = aug[i];
        aug[i] = aug[maxRow];
        aug[maxRow] = tmp;
      }

      final pivot = aug[i][i];
      if (pivot.abs() < 1e-12) {
        return const FormulaValue.error(FormulaError.num);
      }

      // Scale pivot row
      for (var j = 0; j < 2 * n; j++) {
        aug[i][j] /= pivot;
      }

      // Eliminate column in all other rows
      for (var k = 0; k < n; k++) {
        if (k == i) continue;
        final factor = aug[k][i];
        for (var j = 0; j < 2 * n; j++) {
          aug[k][j] -= factor * aug[i][j];
        }
      }
    }

    // Extract the inverse (right half of augmented matrix)
    final result = <List<FormulaValue>>[];
    for (var r = 0; r < n; r++) {
      final row = <FormulaValue>[];
      for (var c = n; c < 2 * n; c++) {
        var v = aug[r][c];
        // Round near-integer values
        if ((v - v.roundToDouble()).abs() < 1e-10) {
          v = v.roundToDouble();
        }
        row.add(FormulaValue.number(v));
      }
      result.add(row);
    }
    return RangeValue(result);
  }
}
