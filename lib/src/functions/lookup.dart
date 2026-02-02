import 'package:a1/a1.dart';

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
    HlookupFunction(),
    LookupFunction(),
    ChooseFunction(),
    XmatchFunction(),
    XlookupFunction(),
    RowFunction(),
    ColumnFunction(),
    RowsFunction(),
    ColumnsFunction(),
    AddressFunction(),
    IndirectFunction(),
    OffsetFunction(),
    TransposeFunction(),
    HyperlinkFunction(),
    AreasFunction(),
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

/// HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])
class HlookupFunction extends FormulaFunction {
  @override
  String get name => 'HLOOKUP';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);

    final lookupValue = values[0];
    final tableValue = values[1];
    final rowIndex = values[2].toNumber()?.toInt();
    final rangeLookup = args.length > 3 ? values[3].toBool() : true;

    if (tableValue is! RangeValue) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (rowIndex == null || rowIndex < 1) {
      return const FormulaValue.error(FormulaError.value);
    }

    final table = tableValue.values;
    if (table.isEmpty) return const FormulaValue.error(FormulaError.na);
    if (rowIndex > table.length) {
      return const FormulaValue.error(FormulaError.ref);
    }

    final firstRow = table[0];

    if (rangeLookup) {
      // Approximate match in first row
      final lookupNum = lookupValue.toNumber();
      if (lookupNum == null) return const FormulaValue.error(FormulaError.na);

      int? bestCol;
      num? bestValue;
      for (var i = 0; i < firstRow.length; i++) {
        final cellNum = firstRow[i].toNumber();
        if (cellNum == null) continue;
        if (cellNum <= lookupNum) {
          if (bestValue == null || cellNum > bestValue) {
            bestValue = cellNum;
            bestCol = i;
          }
        }
      }
      if (bestCol == null) return const FormulaValue.error(FormulaError.na);
      return table[rowIndex - 1][bestCol];
    } else {
      // Exact match in first row
      for (var i = 0; i < firstRow.length; i++) {
        if (_valuesEqual(firstRow[i], lookupValue)) {
          return table[rowIndex - 1][i];
        }
      }
      return const FormulaValue.error(FormulaError.na);
    }
  }
}

/// LOOKUP(lookup_value, lookup_vector, [result_vector])
class LookupFunction extends FormulaFunction {
  @override
  String get name => 'LOOKUP';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);

    final lookupValue = values[0];
    final lookupArray = values[1];

    if (lookupArray is! RangeValue) {
      return const FormulaValue.error(FormulaError.value);
    }

    final lookupList = _to1DList(lookupArray);
    final resultList = args.length > 2
        ? (values[2] is RangeValue
            ? _to1DList(values[2] as RangeValue)
            : [values[2]])
        : lookupList;

    // Approximate match (assumes sorted ascending): find largest <= lookupValue
    final lookupNum = lookupValue.toNumber();
    if (lookupNum != null) {
      int? bestIndex;
      num? bestValue;
      for (var i = 0; i < lookupList.length; i++) {
        final cellNum = lookupList[i].toNumber();
        if (cellNum == null) continue;
        if (cellNum <= lookupNum) {
          if (bestValue == null || cellNum > bestValue) {
            bestValue = cellNum;
            bestIndex = i;
          }
        }
      }
      if (bestIndex == null) return const FormulaValue.error(FormulaError.na);
      if (bestIndex < resultList.length) return resultList[bestIndex];
      return const FormulaValue.error(FormulaError.na);
    }

    // Text match: find last exact (case-insensitive) match
    final lookupText = lookupValue.toText().toLowerCase();
    int? bestIndex;
    for (var i = 0; i < lookupList.length; i++) {
      if (lookupList[i].toText().toLowerCase() == lookupText) {
        bestIndex = i;
      }
    }
    if (bestIndex == null) return const FormulaValue.error(FormulaError.na);
    if (bestIndex < resultList.length) return resultList[bestIndex];
    return const FormulaValue.error(FormulaError.na);
  }
}

/// CHOOSE(index_num, value1, [value2], ...) - Returns value at index.
class ChooseFunction extends FormulaFunction {
  @override
  String get name => 'CHOOSE';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => -1;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final indexValue = args[0].evaluate(context);
    if (indexValue.isError) return indexValue;
    final index = indexValue.toNumber()?.toInt();

    if (index == null || index < 1 || index > args.length - 1) {
      return const FormulaValue.error(FormulaError.value);
    }

    return args[index].evaluate(context);
  }
}

/// XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])
class XmatchFunction extends FormulaFunction {
  @override
  String get name => 'XMATCH';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);

    final lookupValue = values[0];
    final arrayValue = values[1];
    final matchMode = args.length > 2 ? values[2].toNumber()?.toInt() ?? 0 : 0;
    final searchMode =
        args.length > 3 ? values[3].toNumber()?.toInt() ?? 1 : 1;

    if (arrayValue is! RangeValue) {
      return const FormulaValue.error(FormulaError.value);
    }

    final flat = _to1DList(arrayValue);
    final index = _xSearch(lookupValue, flat, matchMode, searchMode);

    if (index == null) return const FormulaValue.error(FormulaError.na);
    return FormulaValue.number(index + 1); // 1-indexed
  }
}

/// XLOOKUP(lookup, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
class XlookupFunction extends FormulaFunction {
  @override
  String get name => 'XLOOKUP';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 6;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final lookupValue = args[0].evaluate(context);
    if (lookupValue.isError) return lookupValue;
    final lookupArray = args[1].evaluate(context);
    if (lookupArray.isError) return lookupArray;
    final returnArray = args[2].evaluate(context);
    if (returnArray.isError) return returnArray;

    final matchMode =
        args.length > 4 ? args[4].evaluate(context).toNumber()?.toInt() ?? 0 : 0;
    final searchMode =
        args.length > 5 ? args[5].evaluate(context).toNumber()?.toInt() ?? 1 : 1;

    if (lookupArray is! RangeValue) {
      return const FormulaValue.error(FormulaError.value);
    }

    final lookupList = _to1DList(lookupArray);
    final returnList = returnArray is RangeValue
        ? _to1DList(returnArray)
        : [returnArray];

    final index = _xSearch(lookupValue, lookupList, matchMode, searchMode);

    if (index == null) {
      // Return if_not_found or #N/A
      if (args.length > 3) return args[3].evaluate(context);
      return const FormulaValue.error(FormulaError.na);
    }

    if (index < returnList.length) return returnList[index];
    return const FormulaValue.error(FormulaError.na);
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

/// Flatten a RangeValue to a 1D list.
List<FormulaValue> _to1DList(RangeValue range) {
  final values = range.values;
  if (values.length == 1) return values[0];
  return [for (final row in values) row[0]];
}

/// Shared search logic for XMATCH and XLOOKUP.
///
/// match_mode: 0=exact, -1=exact or next smaller, 1=exact or next larger, 2=wildcard
/// search_mode: 1=first-to-last, -1=last-to-first
int? _xSearch(
    FormulaValue lookupValue, List<FormulaValue> array, int matchMode, int searchMode) {
  final forward = searchMode >= 0;

  if (matchMode == 0) {
    // Exact match
    if (forward) {
      for (var i = 0; i < array.length; i++) {
        if (_valuesEqual(array[i], lookupValue)) return i;
      }
    } else {
      for (var i = array.length - 1; i >= 0; i--) {
        if (_valuesEqual(array[i], lookupValue)) return i;
      }
    }
    return null;
  }

  if (matchMode == 2) {
    // Wildcard match
    final pattern = _wildcardToRegex(lookupValue.toText());
    if (forward) {
      for (var i = 0; i < array.length; i++) {
        if (pattern.hasMatch(array[i].toText())) return i;
      }
    } else {
      for (var i = array.length - 1; i >= 0; i--) {
        if (pattern.hasMatch(array[i].toText())) return i;
      }
    }
    return null;
  }

  // matchMode -1 (next smaller) or 1 (next larger)
  final lookupNum = lookupValue.toNumber();
  if (lookupNum == null) return null;

  // Try exact first
  for (var i = 0; i < array.length; i++) {
    if (_valuesEqual(array[i], lookupValue)) return i;
  }

  if (matchMode == -1) {
    // Next smaller
    int? bestIndex;
    num? bestValue;
    for (var i = 0; i < array.length; i++) {
      final cellNum = array[i].toNumber();
      if (cellNum == null) continue;
      if (cellNum < lookupNum) {
        if (bestValue == null || cellNum > bestValue) {
          bestValue = cellNum;
          bestIndex = i;
        }
      }
    }
    return bestIndex;
  }

  if (matchMode == 1) {
    // Next larger
    int? bestIndex;
    num? bestValue;
    for (var i = 0; i < array.length; i++) {
      final cellNum = array[i].toNumber();
      if (cellNum == null) continue;
      if (cellNum > lookupNum) {
        if (bestValue == null || cellNum < bestValue) {
          bestValue = cellNum;
          bestIndex = i;
        }
      }
    }
    return bestIndex;
  }

  return null;
}

/// Convert a wildcard pattern (? and *) to a RegExp.
RegExp _wildcardToRegex(String pattern) {
  final buffer = StringBuffer('^');
  for (var i = 0; i < pattern.length; i++) {
    final ch = pattern[i];
    if (ch == '~' && i + 1 < pattern.length) {
      buffer.write(RegExp.escape(pattern[i + 1]));
      i++;
    } else if (ch == '?') {
      buffer.write('.');
    } else if (ch == '*') {
      buffer.write('.*');
    } else {
      buffer.write(RegExp.escape(ch));
    }
  }
  buffer.write(r'$');
  return RegExp(buffer.toString(), caseSensitive: false);
}

// -- Wave 6: Lookup & Reference functions ------------------------------------

/// ROW([reference]) - Returns the row number of a reference.
class RowFunction extends FormulaFunction {
  @override
  String get name => 'ROW';
  @override
  int get minArgs => 0;
  @override
  int get maxArgs => 1;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    if (args.isEmpty) {
      return FormulaValue.number(context.currentCell.row + 1);
    }
    final arg = args[0];
    if (arg is CellRefNode) {
      final cell = arg.reference.from.a1;
      if (cell == null) return const FormulaValue.error(FormulaError.ref);
      return FormulaValue.number(cell.row + 1);
    }
    if (arg is RangeRefNode) {
      final from = arg.reference.from.a1;
      if (from == null) return const FormulaValue.error(FormulaError.ref);
      return FormulaValue.number(from.row + 1);
    }
    return const FormulaValue.error(FormulaError.value);
  }
}

/// COLUMN([reference]) - Returns the column number of a reference.
class ColumnFunction extends FormulaFunction {
  @override
  String get name => 'COLUMN';
  @override
  int get minArgs => 0;
  @override
  int get maxArgs => 1;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    if (args.isEmpty) {
      return FormulaValue.number(context.currentCell.column + 1);
    }
    final arg = args[0];
    if (arg is CellRefNode) {
      final cell = arg.reference.from.a1;
      if (cell == null) return const FormulaValue.error(FormulaError.ref);
      return FormulaValue.number(cell.column + 1);
    }
    if (arg is RangeRefNode) {
      final from = arg.reference.from.a1;
      if (from == null) return const FormulaValue.error(FormulaError.ref);
      return FormulaValue.number(from.column + 1);
    }
    return const FormulaValue.error(FormulaError.value);
  }
}

/// ROWS(reference) - Returns the number of rows in a reference.
class RowsFunction extends FormulaFunction {
  @override
  String get name => 'ROWS';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arg = args[0];
    if (arg is CellRefNode) return const FormulaValue.number(1);
    if (arg is RangeRefNode) {
      final from = arg.reference.from.a1;
      final to = arg.reference.to.a1;
      if (from == null || to == null) {
        return const FormulaValue.error(FormulaError.ref);
      }
      return FormulaValue.number(to.row - from.row + 1);
    }
    final value = arg.evaluate(context);
    if (value is RangeValue) return FormulaValue.number(value.rowCount);
    return const FormulaValue.error(FormulaError.value);
  }
}

/// COLUMNS(reference) - Returns the number of columns in a reference.
class ColumnsFunction extends FormulaFunction {
  @override
  String get name => 'COLUMNS';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arg = args[0];
    if (arg is CellRefNode) return const FormulaValue.number(1);
    if (arg is RangeRefNode) {
      final from = arg.reference.from.a1;
      final to = arg.reference.to.a1;
      if (from == null || to == null) {
        return const FormulaValue.error(FormulaError.ref);
      }
      return FormulaValue.number(to.column - from.column + 1);
    }
    final value = arg.evaluate(context);
    if (value is RangeValue) return FormulaValue.number(value.columnCount);
    return const FormulaValue.error(FormulaError.value);
  }
}

/// ADDRESS(row_num, col_num, [abs_num], [a1], [sheet_text])
///
/// abs_num: 1=$A$1, 2=A$1, 3=$A1, 4=A1 (default 1).
class AddressFunction extends FormulaFunction {
  @override
  String get name => 'ADDRESS';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 5;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);

    final rowNum = values[0].toNumber()?.toInt();
    final colNum = values[1].toNumber()?.toInt();
    final absNum = args.length > 2 ? values[2].toNumber()?.toInt() ?? 1 : 1;
    final a1Style = args.length > 3 ? values[3].toBool() : true;
    final sheet = args.length > 4 ? values[4].toText() : null;

    if (rowNum == null || rowNum < 1 || colNum == null || colNum < 1) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (absNum < 1 || absNum > 4) {
      return const FormulaValue.error(FormulaError.value);
    }

    String result;
    if (a1Style) {
      final colAbsolute = absNum == 1 || absNum == 3;
      final rowAbsolute = absNum == 1 || absNum == 2;
      final cell = A1.fromVector(colNum - 1, rowNum - 1);
      final colStr = colAbsolute ? '\$${cell.letters}' : cell.letters;
      final rowStr = rowAbsolute ? '\$${cell.digits}' : '${cell.digits}';
      result = '$colStr$rowStr';
    } else {
      final rowAbsolute = absNum == 1 || absNum == 2;
      final colAbsolute = absNum == 1 || absNum == 3;
      final rowStr = rowAbsolute ? 'R$rowNum' : 'R[$rowNum]';
      final colStr = colAbsolute ? 'C$colNum' : 'C[$colNum]';
      result = '$rowStr$colStr';
    }

    if (sheet != null && sheet.isNotEmpty) {
      result = '$sheet!$result';
    }

    return FormulaValue.text(result);
  }
}

/// INDIRECT(ref_text, [a1]) - Returns the reference specified by a text string.
class IndirectFunction extends FormulaFunction {
  @override
  String get name => 'INDIRECT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    if (values[0].isError) return values[0];

    final refText = values[0].toText();
    final a1Style = args.length > 1 ? values[1].toBool() : true;
    if (!a1Style) return const FormulaValue.error(FormulaError.ref);

    final ref = A1Reference.tryParse(refText);
    if (ref == null) return const FormulaValue.error(FormulaError.ref);

    final from = ref.from.a1;
    final to = ref.to.a1;

    if (from == null) return const FormulaValue.error(FormulaError.ref);

    // Single cell if from == to
    if (to == null || (from.row == to.row && from.column == to.column)) {
      return context.getCellValue(from);
    }

    return context.getRangeValues(ref.range);
  }
}

/// OFFSET(reference, rows, cols, [height], [width])
///
/// Returns a reference offset from a starting reference.
class OffsetFunction extends FormulaFunction {
  @override
  String get name => 'OFFSET';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 5;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final ref = args[0];
    final rowsVal = args[1].evaluate(context);
    final colsVal = args[2].evaluate(context);

    if (rowsVal.isError) return rowsVal;
    if (colsVal.isError) return colsVal;

    final rowOffset = rowsVal.toNumber()?.toInt();
    final colOffset = colsVal.toNumber()?.toInt();
    if (rowOffset == null || colOffset == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    int baseRow, baseCol;
    int defaultHeight = 1, defaultWidth = 1;

    if (ref is CellRefNode) {
      final cell = ref.reference.from.a1;
      if (cell == null) return const FormulaValue.error(FormulaError.ref);
      baseRow = cell.row;
      baseCol = cell.column;
    } else if (ref is RangeRefNode) {
      final from = ref.reference.from.a1;
      final to = ref.reference.to.a1;
      if (from == null || to == null) {
        return const FormulaValue.error(FormulaError.ref);
      }
      baseRow = from.row;
      baseCol = from.column;
      defaultHeight = to.row - from.row + 1;
      defaultWidth = to.column - from.column + 1;
    } else {
      return const FormulaValue.error(FormulaError.value);
    }

    final heightVal = args.length > 3 ? args[3].evaluate(context) : null;
    final widthVal = args.length > 4 ? args[4].evaluate(context) : null;

    final height = heightVal?.toNumber()?.toInt() ?? defaultHeight;
    final width = widthVal?.toNumber()?.toInt() ?? defaultWidth;

    if (height <= 0 || width <= 0) {
      return const FormulaValue.error(FormulaError.ref);
    }

    final newRow = baseRow + rowOffset;
    final newCol = baseCol + colOffset;

    if (newRow < 0 || newCol < 0) {
      return const FormulaValue.error(FormulaError.ref);
    }

    if (height == 1 && width == 1) {
      return context.getCellValue(A1.fromVector(newCol, newRow));
    }

    final fromCell = A1.fromVector(newCol, newRow);
    final toCell = A1.fromVector(newCol + width - 1, newRow + height - 1);
    final range = A1Range.fromA1s(fromCell, toCell);
    return context.getRangeValues(range);
  }
}

/// TRANSPOSE(array) - Swaps rows and columns of an array.
class TransposeFunction extends FormulaFunction {
  @override
  String get name => 'TRANSPOSE';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    if (value.isError) return value;

    if (value is! RangeValue) {
      return RangeValue([
        [value]
      ]);
    }

    final rows = value.values;
    if (rows.isEmpty) return value;

    final numRows = rows.length;
    final numCols = rows[0].length;

    final transposed = List.generate(
      numCols,
      (col) => List.generate(numRows, (row) => rows[row][col]),
    );

    return RangeValue(transposed);
  }
}

/// HYPERLINK(link_location, [friendly_name])
///
/// Returns the friendly name or the URL as text.
class HyperlinkFunction extends FormulaFunction {
  @override
  String get name => 'HYPERLINK';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    if (values[0].isError) return values[0];

    if (args.length > 1) {
      if (values[1].isError) return values[1];
      return FormulaValue.text(values[1].toText());
    }
    return FormulaValue.text(values[0].toText());
  }
}

/// AREAS(reference) - Returns the number of areas in a reference.
///
/// Multi-area references are not supported; always returns 1 for a valid ref.
class AreasFunction extends FormulaFunction {
  @override
  String get name => 'AREAS';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arg = args[0];
    if (arg is CellRefNode || arg is RangeRefNode) {
      return const FormulaValue.number(1);
    }
    return const FormulaValue.error(FormulaError.value);
  }
}
