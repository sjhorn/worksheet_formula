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
    SumIfsFunction(),
    CountIfsFunction(),
    AverageIfsFunction(),
    MedianFunction(),
    ModeSnglFunction(),
    ModeAliasFunction(),
    LargeFunction(),
    SmallFunction(),
    RankEqFunction(),
    RankAliasFunction(),
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

/// SUMIFS(sum_range, criteria_range1, criteria1, ...) - Sum with multiple criteria.
class SumIfsFunction extends FormulaFunction {
  @override
  String get name => 'SUMIFS';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    if ((args.length - 1).isOdd) {
      return const FormulaValue.error(FormulaError.value);
    }

    final sumRange = _flattenValues(args[0].evaluate(context));
    final criteriaPairs = <(List<FormulaValue>, FormulaValue)>[];

    for (var i = 1; i < args.length; i += 2) {
      final range = _flattenValues(args[i].evaluate(context));
      final criteria = args[i + 1].evaluate(context);
      criteriaPairs.add((range, criteria));
    }

    var sum = 0.0;
    for (var i = 0; i < sumRange.length; i++) {
      var matchesAll = true;
      for (final (range, criteria) in criteriaPairs) {
        if (i >= range.length || !_matchesCriteria(range[i], criteria)) {
          matchesAll = false;
          break;
        }
      }
      if (matchesAll) {
        final n = sumRange[i].toNumber();
        if (n != null) sum += n;
      }
    }
    return FormulaValue.number(sum);
  }
}

/// COUNTIFS(criteria_range1, criteria1, ...) - Count with multiple criteria.
class CountIfsFunction extends FormulaFunction {
  @override
  String get name => 'COUNTIFS';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    if (args.length.isOdd) {
      return const FormulaValue.error(FormulaError.value);
    }

    final criteriaPairs = <(List<FormulaValue>, FormulaValue)>[];
    for (var i = 0; i < args.length; i += 2) {
      final range = _flattenValues(args[i].evaluate(context));
      final criteria = args[i + 1].evaluate(context);
      criteriaPairs.add((range, criteria));
    }

    final length = criteriaPairs.first.$1.length;
    var count = 0;
    for (var i = 0; i < length; i++) {
      var matchesAll = true;
      for (final (range, criteria) in criteriaPairs) {
        if (i >= range.length || !_matchesCriteria(range[i], criteria)) {
          matchesAll = false;
          break;
        }
      }
      if (matchesAll) count++;
    }
    return FormulaValue.number(count);
  }
}

/// AVERAGEIFS(average_range, criteria_range1, criteria1, ...) - Average with multiple criteria.
class AverageIfsFunction extends FormulaFunction {
  @override
  String get name => 'AVERAGEIFS';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    if ((args.length - 1).isOdd) {
      return const FormulaValue.error(FormulaError.value);
    }

    final avgRange = _flattenValues(args[0].evaluate(context));
    final criteriaPairs = <(List<FormulaValue>, FormulaValue)>[];

    for (var i = 1; i < args.length; i += 2) {
      final range = _flattenValues(args[i].evaluate(context));
      final criteria = args[i + 1].evaluate(context);
      criteriaPairs.add((range, criteria));
    }

    var sum = 0.0;
    var count = 0;
    for (var i = 0; i < avgRange.length; i++) {
      var matchesAll = true;
      for (final (range, criteria) in criteriaPairs) {
        if (i >= range.length || !_matchesCriteria(range[i], criteria)) {
          matchesAll = false;
          break;
        }
      }
      if (matchesAll) {
        final n = avgRange[i].toNumber();
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

/// MEDIAN(number1, [number2], ...) - Returns the middle value.
class MedianFunction extends FormulaFunction {
  @override
  String get name => 'MEDIAN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = collectNumbers(args, context).toList();
    if (numbers.isEmpty) {
      return const FormulaValue.error(FormulaError.num);
    }
    numbers.sort();
    final mid = numbers.length ~/ 2;
    if (numbers.length.isOdd) {
      return FormulaValue.number(numbers[mid]);
    }
    return FormulaValue.number((numbers[mid - 1] + numbers[mid]) / 2);
  }
}

/// MODE.SNGL(number1, [number2], ...) - Returns the most frequent value.
class ModeSnglFunction extends FormulaFunction {
  @override
  String get name => 'MODE.SNGL';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = collectNumbers(args, context).toList();
    if (numbers.isEmpty) {
      return const FormulaValue.error(FormulaError.na);
    }

    final freq = <num, int>{};
    for (final n in numbers) {
      freq[n] = (freq[n] ?? 0) + 1;
    }

    num? modeValue;
    var maxCount = 1;
    for (final entry in freq.entries) {
      if (entry.value > maxCount) {
        maxCount = entry.value;
        modeValue = entry.key;
      }
    }

    if (modeValue == null) {
      return const FormulaValue.error(FormulaError.na);
    }
    return FormulaValue.number(modeValue);
  }
}

/// MODE - Alias for MODE.SNGL.
class ModeAliasFunction extends ModeSnglFunction {
  @override
  String get name => 'MODE';
}

/// LARGE(array, k) - Returns the k-th largest value.
class LargeFunction extends FormulaFunction {
  @override
  String get name => 'LARGE';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arrayValue = args[0].evaluate(context);
    if (arrayValue.isError) return arrayValue;
    final kValue = args[1].evaluate(context);
    final k = kValue.toNumber()?.toInt();

    if (k == null) return const FormulaValue.error(FormulaError.value);

    final numbers = <num>[];
    if (arrayValue is RangeValue) {
      numbers.addAll(arrayValue.numbers);
    } else {
      final n = arrayValue.toNumber();
      if (n != null) numbers.add(n);
    }

    if (k < 1 || k > numbers.length) {
      return const FormulaValue.error(FormulaError.num);
    }

    numbers.sort((a, b) => b.compareTo(a)); // Descending
    return FormulaValue.number(numbers[k - 1]);
  }
}

/// SMALL(array, k) - Returns the k-th smallest value.
class SmallFunction extends FormulaFunction {
  @override
  String get name => 'SMALL';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arrayValue = args[0].evaluate(context);
    if (arrayValue.isError) return arrayValue;
    final kValue = args[1].evaluate(context);
    final k = kValue.toNumber()?.toInt();

    if (k == null) return const FormulaValue.error(FormulaError.value);

    final numbers = <num>[];
    if (arrayValue is RangeValue) {
      numbers.addAll(arrayValue.numbers);
    } else {
      final n = arrayValue.toNumber();
      if (n != null) numbers.add(n);
    }

    if (k < 1 || k > numbers.length) {
      return const FormulaValue.error(FormulaError.num);
    }

    numbers.sort();
    return FormulaValue.number(numbers[k - 1]);
  }
}

/// RANK.EQ(number, ref, [order]) - Rank of a number in a list.
class RankEqFunction extends FormulaFunction {
  @override
  String get name => 'RANK.EQ';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final number = values[0].toNumber();
    final refValue = values[1];
    final order = args.length > 2 ? values[2].toNumber()?.toInt() ?? 0 : 0;

    if (number == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    final numbers = <num>[];
    if (refValue is RangeValue) {
      numbers.addAll(refValue.numbers);
    } else {
      final n = refValue.toNumber();
      if (n != null) numbers.add(n);
    }

    if (!numbers.contains(number)) {
      return const FormulaValue.error(FormulaError.na);
    }

    if (order == 0) {
      // Descending rank (largest = 1)
      final rank = numbers.where((n) => n > number).length + 1;
      return FormulaValue.number(rank);
    } else {
      // Ascending rank (smallest = 1)
      final rank = numbers.where((n) => n < number).length + 1;
      return FormulaValue.number(rank);
    }
  }
}

/// RANK - Alias for RANK.EQ.
class RankAliasFunction extends RankEqFunction {
  @override
  String get name => 'RANK';
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
