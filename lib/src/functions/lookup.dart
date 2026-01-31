import '../ast/nodes.dart';
import '../evaluation/context.dart';
import '../evaluation/errors.dart';
import '../evaluation/value.dart';
import 'function.dart';
import 'registry.dart';

/// Register all lookup functions.
void registerLookupFunctions(FunctionRegistry registry) {
  registry.registerAll([
    VlookupFunction(),
    IndexFunction(),
    MatchFunction(),
  ]);
}

/// VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
class VlookupFunction extends FormulaFunction {
  @override
  String get name => 'VLOOKUP';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);

    final lookupValue = values[0];
    final tableValue = values[1];
    final colIndex = values[2].toNumber()?.toInt();
    final rangeLookup = args.length > 3 ? values[3].toBool() : true;

    if (tableValue is! RangeValue) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (colIndex == null || colIndex < 1) {
      return const FormulaValue.error(FormulaError.value);
    }

    final table = tableValue.values;
    if (table.isEmpty) return const FormulaValue.error(FormulaError.na);

    final numCols = table[0].length;
    if (colIndex > numCols) {
      return const FormulaValue.error(FormulaError.ref);
    }

    if (rangeLookup) {
      // Approximate match: find largest first-column value <= lookupValue
      return _approximateMatch(table, lookupValue, colIndex);
    } else {
      // Exact match
      return _exactMatch(table, lookupValue, colIndex);
    }
  }

  FormulaValue _exactMatch(
    List<List<FormulaValue>> table,
    FormulaValue lookupValue,
    int colIndex,
  ) {
    for (final row in table) {
      if (_valuesEqual(row[0], lookupValue)) {
        return row[colIndex - 1];
      }
    }
    return const FormulaValue.error(FormulaError.na);
  }

  FormulaValue _approximateMatch(
    List<List<FormulaValue>> table,
    FormulaValue lookupValue,
    int colIndex,
  ) {
    final lookupNum = lookupValue.toNumber();
    if (lookupNum == null) return const FormulaValue.error(FormulaError.na);

    int? bestRow;
    num? bestValue;

    for (var i = 0; i < table.length; i++) {
      final cellNum = table[i][0].toNumber();
      if (cellNum == null) continue;
      if (cellNum <= lookupNum) {
        if (bestValue == null || cellNum > bestValue) {
          bestValue = cellNum;
          bestRow = i;
        }
      }
    }

    if (bestRow == null) return const FormulaValue.error(FormulaError.na);
    return table[bestRow][colIndex - 1];
  }
}

/// INDEX(array, row_num, [col_num])
class IndexFunction extends FormulaFunction {
  @override
  String get name => 'INDEX';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);

    final arrayValue = values[0];
    final rowNum = values[1].toNumber()?.toInt();
    final colNum = args.length > 2 ? values[2].toNumber()?.toInt() : 1;

    if (arrayValue is! RangeValue) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (rowNum == null || rowNum < 1 || colNum == null || colNum < 1) {
      return const FormulaValue.error(FormulaError.value);
    }

    final array = arrayValue.values;
    if (rowNum > array.length) {
      return const FormulaValue.error(FormulaError.ref);
    }
    if (colNum > array[rowNum - 1].length) {
      return const FormulaValue.error(FormulaError.ref);
    }

    return array[rowNum - 1][colNum - 1];
  }
}

/// MATCH(lookup_value, lookup_array, [match_type])
class MatchFunction extends FormulaFunction {
  @override
  String get name => 'MATCH';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);

    final lookupValue = values[0];
    final arrayValue = values[1];
    final matchType = args.length > 2 ? values[2].toNumber()?.toInt() ?? 1 : 1;

    if (arrayValue is! RangeValue) {
      return const FormulaValue.error(FormulaError.value);
    }

    // Flatten to 1D: use first column if vertical, first row if horizontal
    final flat = _to1DList(arrayValue);

    if (matchType == 0) {
      // Exact match
      for (var i = 0; i < flat.length; i++) {
        if (_valuesEqual(flat[i], lookupValue)) {
          return FormulaValue.number(i + 1);
        }
      }
      return const FormulaValue.error(FormulaError.na);
    } else if (matchType == 1) {
      // Ascending: find largest value <= lookupValue
      final lookupNum = lookupValue.toNumber();
      if (lookupNum == null) return const FormulaValue.error(FormulaError.na);

      int? bestIndex;
      num? bestValue;
      for (var i = 0; i < flat.length; i++) {
        final cellNum = flat[i].toNumber();
        if (cellNum == null) continue;
        if (cellNum <= lookupNum) {
          if (bestValue == null || cellNum > bestValue) {
            bestValue = cellNum;
            bestIndex = i;
          }
        }
      }
      if (bestIndex == null) return const FormulaValue.error(FormulaError.na);
      return FormulaValue.number(bestIndex + 1);
    } else {
      // Descending (match_type = -1): find smallest value >= lookupValue
      final lookupNum = lookupValue.toNumber();
      if (lookupNum == null) return const FormulaValue.error(FormulaError.na);

      int? bestIndex;
      num? bestValue;
      for (var i = 0; i < flat.length; i++) {
        final cellNum = flat[i].toNumber();
        if (cellNum == null) continue;
        if (cellNum >= lookupNum) {
          if (bestValue == null || cellNum < bestValue) {
            bestValue = cellNum;
            bestIndex = i;
          }
        }
      }
      if (bestIndex == null) return const FormulaValue.error(FormulaError.na);
      return FormulaValue.number(bestIndex + 1);
    }
  }

  List<FormulaValue> _to1DList(RangeValue range) {
    final values = range.values;
    if (values.length == 1) {
      // Single row: use that row
      return values[0];
    }
    // Multiple rows: use first column
    return [for (final row in values) row[0]];
  }
}

// -- Shared helpers -----------------------------------------------------------

/// Compare two FormulaValues for equality (case-insensitive for text).
bool _valuesEqual(FormulaValue a, FormulaValue b) {
  final aNum = a.toNumber();
  final bNum = b.toNumber();
  if (aNum != null && bNum != null) return aNum == bNum;

  return a.toText().toLowerCase() == b.toText().toLowerCase();
}
