import 'dart:math' as math;

import '../ast/nodes.dart';
import '../evaluation/context.dart';
import '../evaluation/errors.dart';
import '../evaluation/value.dart';
import 'function.dart';
import 'registry.dart';

/// Register all database functions.
void registerDatabaseFunctions(FunctionRegistry registry) {
  registry.registerAll([
    DSumFunction(),
    DAverageFunction(),
    DCountFunction(),
    DCountAFunction(),
    DMaxFunction(),
    DMinFunction(),
    DGetFunction(),
    DProductFunction(),
    DStdevFunction(),
    DStdevPFunction(),
    DVarFunction(),
    DVarPFunction(),
  ]);
}

// -- Shared helpers -----------------------------------------------------------

/// Parse a criteria string into (operator, value) if it starts with an operator.
(String, String)? _parseCriteria(String text) {
  if (text.startsWith('>=')) return ('>=', text.substring(2));
  if (text.startsWith('<=')) return ('<=', text.substring(2));
  if (text.startsWith('<>')) return ('<>', text.substring(2));
  if (text.startsWith('>')) return ('>', text.substring(1));
  if (text.startsWith('<')) return ('<', text.substring(1));
  if (text.startsWith('=')) return ('=', text.substring(1));
  return null;
}

/// Check if a cell value matches the given criteria.
bool _matchesCriteria(FormulaValue cellValue, FormulaValue criteria) {
  final criteriaText = criteria.toText();

  final parsed = _parseCriteria(criteriaText);
  if (parsed != null) {
    final (op, compareValue) = parsed;
    final cellNum = cellValue.toNumber();
    final compareNum = num.tryParse(compareValue);

    if (cellNum != null && compareNum != null) {
      return switch (op) {
        '>' => cellNum > compareNum,
        '<' => cellNum < compareNum,
        '>=' => cellNum >= compareNum,
        '<=' => cellNum <= compareNum,
        '<>' => cellNum != compareNum,
        '=' => cellNum == compareNum,
        _ => false,
      };
    }

    // Text comparison for <> and =
    if (op == '<>') {
      return cellValue.toText().toLowerCase() !=
          compareValue.toLowerCase();
    }
    if (op == '=') {
      return cellValue.toText().toLowerCase() ==
          compareValue.toLowerCase();
    }

    return false;
  }

  // Numeric exact match
  final criteriaNum = criteria.toNumber();
  final cellNum = cellValue.toNumber();
  if (criteriaNum != null && cellNum != null && criteria is NumberValue) {
    return cellNum == criteriaNum;
  }

  // Text exact match (case-insensitive)
  return cellValue.toText().toLowerCase() == criteriaText.toLowerCase();
}

/// Resolve a field argument to a 0-based column index.
/// Returns -1 if invalid.
int _resolveField(RangeValue database, FormulaValue field) {
  if (field is TextValue) {
    final target = field.value.toLowerCase();
    final headers = database.values[0];
    for (var i = 0; i < headers.length; i++) {
      if (headers[i].toText().toLowerCase() == target) return i;
    }
    return -1;
  }
  if (field is NumberValue) {
    final col = field.value.toInt();
    if (col < 1 || col > database.columnCount) return -1;
    return col - 1;
  }
  return -1;
}

/// Get indices of data rows (0-based, relative to row 1+) that match criteria.
List<int> _getMatchingRows(RangeValue database, RangeValue criteria) {
  final dbHeaders = database.values[0];
  final critHeaders = criteria.values[0];
  final dataRowCount = database.rowCount - 1; // exclude header row

  // Map criteria column indices to database column indices
  final columnMap = <int, int>{}; // critCol -> dbCol
  for (var c = 0; c < critHeaders.length; c++) {
    final critHeader = critHeaders[c].toText().toLowerCase();
    for (var d = 0; d < dbHeaders.length; d++) {
      if (dbHeaders[d].toText().toLowerCase() == critHeader) {
        columnMap[c] = d;
        break;
      }
    }
    // If criteria header not found in database, skip it
  }

  final matching = <int>[];

  // Each criteria row (row 1+) is OR'd together
  final critRowCount = criteria.rowCount - 1;

  for (var dataRow = 0; dataRow < dataRowCount; dataRow++) {
    final dbRow = database.values[dataRow + 1]; // +1 to skip header

    var matchesAnyRow = false;
    for (var critRow = 0; critRow < critRowCount; critRow++) {
      final critValues = criteria.values[critRow + 1]; // +1 to skip header

      var matchesAllCols = true;
      for (final entry in columnMap.entries) {
        final critCol = entry.key;
        final dbCol = entry.value;

        if (critCol >= critValues.length) continue;
        final critValue = critValues[critCol];

        // Empty criteria cells are ignored (match all)
        if (critValue is EmptyValue) continue;
        if (critValue is TextValue && critValue.value.isEmpty) continue;

        if (!_matchesCriteria(dbRow[dbCol], critValue)) {
          matchesAllCols = false;
          break;
        }
      }

      if (matchesAllCols) {
        matchesAnyRow = true;
        break;
      }
    }

    if (matchesAnyRow) matching.add(dataRow);
  }

  return matching;
}

/// Extract and validate the common database function arguments.
/// Returns null and the error value if validation fails.
({RangeValue db, int fieldCol, List<int> rows, FormulaValue? error})
    _extractDbArgs(List<FormulaNode> args, EvaluationContext context) {
  final dbValue = args[0].evaluate(context);
  final fieldValue = args[1].evaluate(context);
  final critValue = args[2].evaluate(context);

  if (dbValue is! RangeValue) {
    return (
      db: const RangeValue([]),
      fieldCol: -1,
      rows: const [],
      error: const FormulaValue.error(FormulaError.value),
    );
  }
  if (critValue is! RangeValue) {
    return (
      db: const RangeValue([]),
      fieldCol: -1,
      rows: const [],
      error: const FormulaValue.error(FormulaError.value),
    );
  }

  final fieldCol = _resolveField(dbValue, fieldValue);
  if (fieldCol < 0) {
    return (
      db: const RangeValue([]),
      fieldCol: -1,
      rows: const [],
      error: const FormulaValue.error(FormulaError.value),
    );
  }

  final rows = _getMatchingRows(dbValue, critValue);
  return (db: dbValue, fieldCol: fieldCol, rows: rows, error: null);
}

/// Compute variance of a list of numbers.
double? _variance(List<num> numbers, {required bool sample}) {
  final n = numbers.length;
  if (sample && n < 2) return null;
  if (!sample && n < 1) return null;
  final mean = numbers.fold(0.0, (a, b) => a + b) / n;
  var sum = 0.0;
  for (final x in numbers) {
    final d = x - mean;
    sum += d * d;
  }
  return sum / (sample ? n - 1 : n);
}

/// Collect numeric values from the field column of matching rows.
List<num> _collectFieldNumbers(
    RangeValue db, int fieldCol, List<int> rows) {
  final numbers = <num>[];
  for (final row in rows) {
    final cell = db.values[row + 1][fieldCol]; // +1 to skip header
    if (cell is NumberValue) {
      numbers.add(cell.value);
    }
  }
  return numbers;
}

// -- Database functions -------------------------------------------------------

/// DSUM(database, field, criteria) - Sum of matching field values.
class DSumFunction extends FormulaFunction {
  @override
  String get name => 'DSUM';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final extracted = _extractDbArgs(args, context);
    if (extracted.error != null) return extracted.error!;

    final numbers = _collectFieldNumbers(
        extracted.db, extracted.fieldCol, extracted.rows);
    if (numbers.isEmpty) return const FormulaValue.number(0);
    return FormulaValue.number(numbers.fold(0.0, (a, b) => a + b));
  }
}

/// DAVERAGE(database, field, criteria) - Average of matching field values.
class DAverageFunction extends FormulaFunction {
  @override
  String get name => 'DAVERAGE';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final extracted = _extractDbArgs(args, context);
    if (extracted.error != null) return extracted.error!;

    final numbers = _collectFieldNumbers(
        extracted.db, extracted.fieldCol, extracted.rows);
    if (numbers.isEmpty) {
      return const FormulaValue.error(FormulaError.divZero);
    }
    final sum = numbers.fold(0.0, (a, b) => a + b);
    return FormulaValue.number(sum / numbers.length);
  }
}

/// DCOUNT(database, field, criteria) - Count numeric cells in matching rows.
class DCountFunction extends FormulaFunction {
  @override
  String get name => 'DCOUNT';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final extracted = _extractDbArgs(args, context);
    if (extracted.error != null) return extracted.error!;

    final numbers = _collectFieldNumbers(
        extracted.db, extracted.fieldCol, extracted.rows);
    return FormulaValue.number(numbers.length);
  }
}

/// DCOUNTA(database, field, criteria) - Count non-empty cells in matching rows.
class DCountAFunction extends FormulaFunction {
  @override
  String get name => 'DCOUNTA';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final extracted = _extractDbArgs(args, context);
    if (extracted.error != null) return extracted.error!;

    var count = 0;
    for (final row in extracted.rows) {
      final cell = extracted.db.values[row + 1][extracted.fieldCol];
      if (cell is! EmptyValue) count++;
    }
    return FormulaValue.number(count);
  }
}

/// DMAX(database, field, criteria) - Maximum of matching numeric field values.
class DMaxFunction extends FormulaFunction {
  @override
  String get name => 'DMAX';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final extracted = _extractDbArgs(args, context);
    if (extracted.error != null) return extracted.error!;

    final numbers = _collectFieldNumbers(
        extracted.db, extracted.fieldCol, extracted.rows);
    if (numbers.isEmpty) return const FormulaValue.number(0);
    return FormulaValue.number(numbers.reduce(math.max));
  }
}

/// DMIN(database, field, criteria) - Minimum of matching numeric field values.
class DMinFunction extends FormulaFunction {
  @override
  String get name => 'DMIN';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final extracted = _extractDbArgs(args, context);
    if (extracted.error != null) return extracted.error!;

    final numbers = _collectFieldNumbers(
        extracted.db, extracted.fieldCol, extracted.rows);
    if (numbers.isEmpty) return const FormulaValue.number(0);
    return FormulaValue.number(numbers.reduce(math.min));
  }
}

/// DGET(database, field, criteria) - Return value from single matching row.
class DGetFunction extends FormulaFunction {
  @override
  String get name => 'DGET';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final extracted = _extractDbArgs(args, context);
    if (extracted.error != null) return extracted.error!;

    if (extracted.rows.isEmpty) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (extracted.rows.length > 1) {
      return const FormulaValue.error(FormulaError.num);
    }
    return extracted.db.values[extracted.rows[0] + 1][extracted.fieldCol];
  }
}

/// DPRODUCT(database, field, criteria) - Product of matching numeric values.
class DProductFunction extends FormulaFunction {
  @override
  String get name => 'DPRODUCT';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final extracted = _extractDbArgs(args, context);
    if (extracted.error != null) return extracted.error!;

    final numbers = _collectFieldNumbers(
        extracted.db, extracted.fieldCol, extracted.rows);
    if (numbers.isEmpty) return const FormulaValue.number(0);
    return FormulaValue.number(numbers.fold(1.0, (a, b) => a * b));
  }
}

/// DSTDEV(database, field, criteria) - Sample standard deviation.
class DStdevFunction extends FormulaFunction {
  @override
  String get name => 'DSTDEV';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final extracted = _extractDbArgs(args, context);
    if (extracted.error != null) return extracted.error!;

    final numbers = _collectFieldNumbers(
        extracted.db, extracted.fieldCol, extracted.rows);
    final v = _variance(numbers, sample: true);
    if (v == null) return const FormulaValue.error(FormulaError.divZero);
    return FormulaValue.number(math.sqrt(v));
  }
}

/// DSTDEVP(database, field, criteria) - Population standard deviation.
class DStdevPFunction extends FormulaFunction {
  @override
  String get name => 'DSTDEVP';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final extracted = _extractDbArgs(args, context);
    if (extracted.error != null) return extracted.error!;

    final numbers = _collectFieldNumbers(
        extracted.db, extracted.fieldCol, extracted.rows);
    final v = _variance(numbers, sample: false);
    if (v == null) return const FormulaValue.error(FormulaError.divZero);
    return FormulaValue.number(math.sqrt(v));
  }
}

/// DVAR(database, field, criteria) - Sample variance.
class DVarFunction extends FormulaFunction {
  @override
  String get name => 'DVAR';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final extracted = _extractDbArgs(args, context);
    if (extracted.error != null) return extracted.error!;

    final numbers = _collectFieldNumbers(
        extracted.db, extracted.fieldCol, extracted.rows);
    final v = _variance(numbers, sample: true);
    if (v == null) return const FormulaValue.error(FormulaError.divZero);
    return FormulaValue.number(v);
  }
}

/// DVARP(database, field, criteria) - Population variance.
class DVarPFunction extends FormulaFunction {
  @override
  String get name => 'DVARP';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final extracted = _extractDbArgs(args, context);
    if (extracted.error != null) return extracted.error!;

    final numbers = _collectFieldNumbers(
        extracted.db, extracted.fieldCol, extracted.rows);
    final v = _variance(numbers, sample: false);
    if (v == null) return const FormulaValue.error(FormulaError.divZero);
    return FormulaValue.number(v);
  }
}
