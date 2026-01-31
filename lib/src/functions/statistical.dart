import '../ast/nodes.dart';
import '../evaluation/context.dart';
import '../evaluation/errors.dart';
import '../evaluation/value.dart';
import 'function.dart';
import 'registry.dart';

/// Register all statistical functions.
void registerStatisticalFunctions(FunctionRegistry registry) {
  registry.registerAll([
    CountFunction(),
    CountAFunction(),
    CountBlankFunction(),
    CountIfFunction(),
    SumIfFunction(),
    AverageIfFunction(),
  ]);
}

/// COUNT(value1, [value2], ...) - Counts numeric values.
class CountFunction extends FormulaFunction {
  @override
  String get name => 'COUNT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    var count = 0;
    for (final arg in args) {
      final value = arg.evaluate(context);
      switch (value) {
        case NumberValue():
          count++;
        case RangeValue(values: final matrix):
          for (final row in matrix) {
            for (final cell in row) {
              if (cell is NumberValue) count++;
            }
          }
        default:
          break;
      }
    }
    return FormulaValue.number(count);
  }
}

/// COUNTA(value1, [value2], ...) - Counts non-empty values.
class CountAFunction extends FormulaFunction {
  @override
  String get name => 'COUNTA';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    var count = 0;
    for (final arg in args) {
      final value = arg.evaluate(context);
      switch (value) {
        case EmptyValue():
          break;
        case RangeValue(values: final matrix):
          for (final row in matrix) {
            for (final cell in row) {
              if (cell is! EmptyValue) count++;
            }
          }
        default:
          count++;
      }
    }
    return FormulaValue.number(count);
  }
}

/// COUNTBLANK(range) - Counts empty cells in a range.
class CountBlankFunction extends FormulaFunction {
  @override
  String get name => 'COUNTBLANK';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    if (value is! RangeValue) {
      return FormulaValue.number(value is EmptyValue ? 1 : 0);
    }
    var count = 0;
    for (final cell in value.flat) {
      if (cell is EmptyValue) count++;
    }
    return FormulaValue.number(count);
  }
}

/// COUNTIF(range, criteria) - Counts cells matching criteria.
class CountIfFunction extends FormulaFunction {
  @override
  String get name => 'COUNTIF';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final rangeValue = args[0].evaluate(context);
    final criteria = args[1].evaluate(context);

    final cells = _flattenValues(rangeValue);
    var count = 0;
    for (final cell in cells) {
      if (_matchesCriteria(cell, criteria)) count++;
    }
    return FormulaValue.number(count);
  }
}

/// SUMIF(range, criteria, [sum_range]) - Sums cells matching criteria.
class SumIfFunction extends FormulaFunction {
  @override
  String get name => 'SUMIF';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final rangeValue = args[0].evaluate(context);
    final criteria = args[1].evaluate(context);
    final sumRange = args.length > 2 ? args[2].evaluate(context) : rangeValue;

    final criteriaList = _flattenValues(rangeValue);
    final sumList = _flattenValues(sumRange);

    var sum = 0.0;
    for (var i = 0; i < criteriaList.length && i < sumList.length; i++) {
      if (_matchesCriteria(criteriaList[i], criteria)) {
        final n = sumList[i].toNumber();
        if (n != null) sum += n;
      }
    }
    return FormulaValue.number(sum);
  }
}

/// AVERAGEIF(range, criteria, [average_range]) - Averages cells matching criteria.
class AverageIfFunction extends FormulaFunction {
  @override
  String get name => 'AVERAGEIF';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final rangeValue = args[0].evaluate(context);
    final criteria = args[1].evaluate(context);
    final avgRange = args.length > 2 ? args[2].evaluate(context) : rangeValue;

    final criteriaList = _flattenValues(rangeValue);
    final avgList = _flattenValues(avgRange);

    var sum = 0.0;
    var count = 0;
    for (var i = 0; i < criteriaList.length && i < avgList.length; i++) {
      if (_matchesCriteria(criteriaList[i], criteria)) {
        final n = avgList[i].toNumber();
        if (n != null) {
          sum += n;
          count++;
        }
      }
    }
    if (count == 0) return const FormulaValue.error(FormulaError.divZero);
    return FormulaValue.number(sum / count);
  }
}

// -- Shared helpers -----------------------------------------------------------

/// Flatten a FormulaValue into a list of individual cell values.
List<FormulaValue> _flattenValues(FormulaValue value) {
  if (value is RangeValue) return value.flat.toList();
  return [value];
}

/// Check if a cell value matches the given criteria.
///
/// Criteria can be:
/// - A number: exact numeric match
/// - A boolean: exact boolean match
/// - A string with operator prefix: ">5", "<10", ">=3", "<=7", "<>0", "=5"
/// - A bare string: case-insensitive text match
bool _matchesCriteria(FormulaValue cellValue, FormulaValue criteria) {
  // Parse criteria
  final criteriaText = criteria.toText();

  // Try to parse operator prefix
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
  if (criteriaNum != null && cellNum != null) {
    return cellNum == criteriaNum;
  }

  // Text exact match (case-insensitive)
  return cellValue.toText().toLowerCase() == criteriaText.toLowerCase();
}

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
