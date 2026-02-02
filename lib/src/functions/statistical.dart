import 'dart:math' as math;

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
    StdevSFunction(),
    StdevPFunction(),
    VarSFunction(),
    VarPFunction(),
    PercentileIncFunction(),
    PercentileExcFunction(),
    PercentRankIncFunction(),
    PercentRankExcFunction(),
    RankAvgFunction(),
    FrequencyFunction(),
    AveDevFunction(),
    AverageAFunction(),
    MaxAFunction(),
    MinAFunction(),
    TrimMeanFunction(),
    GeoMeanFunction(),
    HarMeanFunction(),
    MaxIfsFunction(),
    MinIfsFunction(),
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

/// Collect all values from args, treating TRUE=1, FALSE=0, text=0, skip empties in ranges.
List<num> _collectAllValues(
    List<FormulaNode> args, EvaluationContext context) {
  final result = <num>[];
  for (final arg in args) {
    final value = arg.evaluate(context);
    switch (value) {
      case NumberValue(value: final n):
        result.add(n);
      case BooleanValue(value: final b):
        result.add(b ? 1 : 0);
      case TextValue():
        result.add(0);
      case RangeValue(values: final matrix):
        for (final row in matrix) {
          for (final cell in row) {
            switch (cell) {
              case NumberValue(value: final n):
                result.add(n);
              case BooleanValue(value: final b):
                result.add(b ? 1 : 0);
              case TextValue():
                result.add(0);
              default:
                break;
            }
          }
        }
      default:
        break;
    }
  }
  return result;
}

/// STDEV.S(number1, [number2], ...) - Sample standard deviation.
class StdevSFunction extends FormulaFunction {
  @override
  String get name => 'STDEV.S';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = collectNumbers(args, context).toList();
    final v = _variance(numbers, sample: true);
    if (v == null) return const FormulaValue.error(FormulaError.divZero);
    return FormulaValue.number(math.sqrt(v));
  }
}

/// STDEV.P(number1, [number2], ...) - Population standard deviation.
class StdevPFunction extends FormulaFunction {
  @override
  String get name => 'STDEV.P';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = collectNumbers(args, context).toList();
    final v = _variance(numbers, sample: false);
    if (v == null) return const FormulaValue.error(FormulaError.divZero);
    return FormulaValue.number(math.sqrt(v));
  }
}

/// VAR.S(number1, [number2], ...) - Sample variance.
class VarSFunction extends FormulaFunction {
  @override
  String get name => 'VAR.S';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = collectNumbers(args, context).toList();
    final v = _variance(numbers, sample: true);
    if (v == null) return const FormulaValue.error(FormulaError.divZero);
    return FormulaValue.number(v);
  }
}

/// VAR.P(number1, [number2], ...) - Population variance.
class VarPFunction extends FormulaFunction {
  @override
  String get name => 'VAR.P';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = collectNumbers(args, context).toList();
    final v = _variance(numbers, sample: false);
    if (v == null) return const FormulaValue.error(FormulaError.divZero);
    return FormulaValue.number(v);
  }
}

/// PERCENTILE.INC(array, k) - Returns the k-th percentile (inclusive).
class PercentileIncFunction extends FormulaFunction {
  @override
  String get name => 'PERCENTILE.INC';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arrayValue = args[0].evaluate(context);
    if (arrayValue.isError) return arrayValue;
    final kValue = args[1].evaluate(context);
    final k = kValue.toNumber()?.toDouble();
    if (k == null) return const FormulaValue.error(FormulaError.value);
    if (k < 0 || k > 1) return const FormulaValue.error(FormulaError.num);

    final numbers = <num>[];
    if (arrayValue is RangeValue) {
      numbers.addAll(arrayValue.numbers);
    } else {
      final n = arrayValue.toNumber();
      if (n != null) numbers.add(n);
    }
    if (numbers.isEmpty) return const FormulaValue.error(FormulaError.num);

    numbers.sort();
    final index = k * (numbers.length - 1);
    final lower = index.floor();
    final upper = index.ceil();
    if (lower == upper) return FormulaValue.number(numbers[lower]);
    final frac = index - lower;
    return FormulaValue.number(
        numbers[lower] + frac * (numbers[upper] - numbers[lower]));
  }
}

/// PERCENTILE.EXC(array, k) - Returns the k-th percentile (exclusive).
class PercentileExcFunction extends FormulaFunction {
  @override
  String get name => 'PERCENTILE.EXC';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arrayValue = args[0].evaluate(context);
    if (arrayValue.isError) return arrayValue;
    final kValue = args[1].evaluate(context);
    final k = kValue.toNumber()?.toDouble();
    if (k == null) return const FormulaValue.error(FormulaError.value);
    if (k <= 0 || k >= 1) return const FormulaValue.error(FormulaError.num);

    final numbers = <num>[];
    if (arrayValue is RangeValue) {
      numbers.addAll(arrayValue.numbers);
    } else {
      final n = arrayValue.toNumber();
      if (n != null) numbers.add(n);
    }
    if (numbers.isEmpty) return const FormulaValue.error(FormulaError.num);

    numbers.sort();
    final n = numbers.length;
    final index = k * (n + 1) - 1;
    if (index < 0 || index > n - 1) {
      return const FormulaValue.error(FormulaError.num);
    }
    final lower = index.floor();
    final upper = index.ceil();
    if (lower == upper) return FormulaValue.number(numbers[lower]);
    final frac = index - lower;
    return FormulaValue.number(
        numbers[lower] + frac * (numbers[upper] - numbers[lower]));
  }
}

/// PERCENTRANK.INC(array, x, [significance]) - Returns percent rank (inclusive).
class PercentRankIncFunction extends FormulaFunction {
  @override
  String get name => 'PERCENTRANK.INC';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final x = values[1].toNumber()?.toDouble();
    if (x == null) return const FormulaValue.error(FormulaError.value);
    final sig = args.length > 2 ? values[2].toNumber()?.toInt() ?? 3 : 3;

    final numbers = <num>[];
    final arrayValue = values[0];
    if (arrayValue is RangeValue) {
      numbers.addAll(arrayValue.numbers);
    } else {
      final n = arrayValue.toNumber();
      if (n != null) numbers.add(n);
    }
    if (numbers.isEmpty) return const FormulaValue.error(FormulaError.num);

    numbers.sort();
    if (x < numbers.first || x > numbers.last) {
      return const FormulaValue.error(FormulaError.na);
    }

    // Find rank by interpolation
    final n = numbers.length;
    for (var i = 0; i < n; i++) {
      if (numbers[i] == x) {
        final rank = i / (n - 1);
        final multiplier = math.pow(10, sig);
        return FormulaValue.number(
            (rank * multiplier).truncateToDouble() / multiplier);
      }
      if (i < n - 1 && numbers[i] < x && x < numbers[i + 1]) {
        final frac = (x - numbers[i]) / (numbers[i + 1] - numbers[i]);
        final rank = (i + frac) / (n - 1);
        final multiplier = math.pow(10, sig);
        return FormulaValue.number(
            (rank * multiplier).truncateToDouble() / multiplier);
      }
    }
    return const FormulaValue.error(FormulaError.na);
  }
}

/// PERCENTRANK.EXC(array, x, [significance]) - Returns percent rank (exclusive).
class PercentRankExcFunction extends FormulaFunction {
  @override
  String get name => 'PERCENTRANK.EXC';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final x = values[1].toNumber()?.toDouble();
    if (x == null) return const FormulaValue.error(FormulaError.value);
    final sig = args.length > 2 ? values[2].toNumber()?.toInt() ?? 3 : 3;

    final numbers = <num>[];
    final arrayValue = values[0];
    if (arrayValue is RangeValue) {
      numbers.addAll(arrayValue.numbers);
    } else {
      final n = arrayValue.toNumber();
      if (n != null) numbers.add(n);
    }
    if (numbers.isEmpty) return const FormulaValue.error(FormulaError.num);

    numbers.sort();
    if (x < numbers.first || x > numbers.last) {
      return const FormulaValue.error(FormulaError.na);
    }

    final n = numbers.length;
    for (var i = 0; i < n; i++) {
      if (numbers[i] == x) {
        final rank = (i + 1) / (n + 1);
        final multiplier = math.pow(10, sig);
        return FormulaValue.number(
            (rank * multiplier).truncateToDouble() / multiplier);
      }
      if (i < n - 1 && numbers[i] < x && x < numbers[i + 1]) {
        final frac = (x - numbers[i]) / (numbers[i + 1] - numbers[i]);
        final rank = (i + 1 + frac) / (n + 1);
        final multiplier = math.pow(10, sig);
        return FormulaValue.number(
            (rank * multiplier).truncateToDouble() / multiplier);
      }
    }
    return const FormulaValue.error(FormulaError.na);
  }
}

/// RANK.AVG(number, ref, [order]) - Rank with averaged ties.
class RankAvgFunction extends FormulaFunction {
  @override
  String get name => 'RANK.AVG';
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

    final tieCount = numbers.where((n) => n == number).length;
    int baseRank;
    if (order == 0) {
      baseRank = numbers.where((n) => n > number).length + 1;
    } else {
      baseRank = numbers.where((n) => n < number).length + 1;
    }
    // Average of ranks for ties: baseRank, baseRank+1, ..., baseRank+tieCount-1
    final avg = baseRank + (tieCount - 1) / 2;
    return FormulaValue.number(avg);
  }
}

/// FREQUENCY(data_array, bins_array) - Returns frequency distribution.
class FrequencyFunction extends FormulaFunction {
  @override
  String get name => 'FREQUENCY';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final dataValue = args[0].evaluate(context);
    final binsValue = args[1].evaluate(context);

    final data = <num>[];
    if (dataValue is RangeValue) {
      data.addAll(dataValue.numbers);
    } else {
      final n = dataValue.toNumber();
      if (n != null) data.add(n);
    }

    final bins = <num>[];
    if (binsValue is RangeValue) {
      bins.addAll(binsValue.numbers);
    } else {
      final n = binsValue.toNumber();
      if (n != null) bins.add(n);
    }

    bins.sort();
    // Result has bins.length + 1 entries
    final counts = List.filled(bins.length + 1, 0);
    for (final d in data) {
      var placed = false;
      for (var i = 0; i < bins.length; i++) {
        if (d <= bins[i]) {
          counts[i]++;
          placed = true;
          break;
        }
      }
      if (!placed) counts[bins.length]++;
    }

    return FormulaValue.range([
      for (final c in counts) [FormulaValue.number(c)]
    ]);
  }
}

/// AVEDEV(number1, [number2], ...) - Average absolute deviation.
class AveDevFunction extends FormulaFunction {
  @override
  String get name => 'AVEDEV';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = collectNumbers(args, context).toList();
    if (numbers.isEmpty) return const FormulaValue.error(FormulaError.num);
    final mean = numbers.fold(0.0, (a, b) => a + b) / numbers.length;
    var sum = 0.0;
    for (final n in numbers) {
      sum += (n - mean).abs();
    }
    return FormulaValue.number(sum / numbers.length);
  }
}

/// AVERAGEA(value1, [value2], ...) - Average including text and logical values.
class AverageAFunction extends FormulaFunction {
  @override
  String get name => 'AVERAGEA';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = _collectAllValues(args, context);
    if (numbers.isEmpty) {
      return const FormulaValue.error(FormulaError.divZero);
    }
    final sum = numbers.fold(0.0, (a, b) => a + b);
    return FormulaValue.number(sum / numbers.length);
  }
}

/// MAXA(value1, [value2], ...) - Max including text and logical values.
class MaxAFunction extends FormulaFunction {
  @override
  String get name => 'MAXA';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = _collectAllValues(args, context);
    if (numbers.isEmpty) return const FormulaValue.number(0);
    return FormulaValue.number(numbers.reduce((a, b) => a > b ? a : b));
  }
}

/// MINA(value1, [value2], ...) - Min including text and logical values.
class MinAFunction extends FormulaFunction {
  @override
  String get name => 'MINA';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = _collectAllValues(args, context);
    if (numbers.isEmpty) return const FormulaValue.number(0);
    return FormulaValue.number(numbers.reduce((a, b) => a < b ? a : b));
  }
}

/// TRIMMEAN(array, percent) - Mean excluding outliers.
class TrimMeanFunction extends FormulaFunction {
  @override
  String get name => 'TRIMMEAN';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arrayValue = args[0].evaluate(context);
    final pctValue = args[1].evaluate(context);
    final percent = pctValue.toNumber()?.toDouble();
    if (percent == null) return const FormulaValue.error(FormulaError.value);
    if (percent < 0 || percent >= 1) {
      return const FormulaValue.error(FormulaError.num);
    }

    final numbers = <num>[];
    if (arrayValue is RangeValue) {
      numbers.addAll(arrayValue.numbers);
    } else {
      final n = arrayValue.toNumber();
      if (n != null) numbers.add(n);
    }
    if (numbers.isEmpty) return const FormulaValue.error(FormulaError.num);

    numbers.sort();
    final trimCount = (numbers.length * percent / 2).floor();
    final trimmed = numbers.sublist(trimCount, numbers.length - trimCount);
    if (trimmed.isEmpty) return const FormulaValue.error(FormulaError.num);
    final sum = trimmed.fold(0.0, (a, b) => a + b);
    return FormulaValue.number(sum / trimmed.length);
  }
}

/// GEOMEAN(number1, [number2], ...) - Geometric mean.
class GeoMeanFunction extends FormulaFunction {
  @override
  String get name => 'GEOMEAN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = collectNumbers(args, context).toList();
    if (numbers.isEmpty) return const FormulaValue.error(FormulaError.num);
    var logSum = 0.0;
    for (final n in numbers) {
      if (n <= 0) return const FormulaValue.error(FormulaError.num);
      logSum += math.log(n);
    }
    return FormulaValue.number(math.exp(logSum / numbers.length));
  }
}

/// HARMEAN(number1, [number2], ...) - Harmonic mean.
class HarMeanFunction extends FormulaFunction {
  @override
  String get name => 'HARMEAN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = collectNumbers(args, context).toList();
    if (numbers.isEmpty) return const FormulaValue.error(FormulaError.num);
    var reciprocalSum = 0.0;
    for (final n in numbers) {
      if (n <= 0) return const FormulaValue.error(FormulaError.num);
      reciprocalSum += 1 / n;
    }
    return FormulaValue.number(numbers.length / reciprocalSum);
  }
}

/// MAXIFS(max_range, criteria_range1, criteria1, ...) - Max with criteria.
class MaxIfsFunction extends FormulaFunction {
  @override
  String get name => 'MAXIFS';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    if ((args.length - 1).isOdd) {
      return const FormulaValue.error(FormulaError.value);
    }

    final maxRange = _flattenValues(args[0].evaluate(context));
    final criteriaPairs = <(List<FormulaValue>, FormulaValue)>[];

    for (var i = 1; i < args.length; i += 2) {
      final range = _flattenValues(args[i].evaluate(context));
      final criteria = args[i + 1].evaluate(context);
      criteriaPairs.add((range, criteria));
    }

    num? maxVal;
    for (var i = 0; i < maxRange.length; i++) {
      var matchesAll = true;
      for (final (range, criteria) in criteriaPairs) {
        if (i >= range.length || !_matchesCriteria(range[i], criteria)) {
          matchesAll = false;
          break;
        }
      }
      if (matchesAll) {
        final n = maxRange[i].toNumber();
        if (n != null && (maxVal == null || n > maxVal)) {
          maxVal = n;
        }
      }
    }
    return FormulaValue.number(maxVal ?? 0);
  }
}

/// MINIFS(min_range, criteria_range1, criteria1, ...) - Min with criteria.
class MinIfsFunction extends FormulaFunction {
  @override
  String get name => 'MINIFS';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    if ((args.length - 1).isOdd) {
      return const FormulaValue.error(FormulaError.value);
    }

    final minRange = _flattenValues(args[0].evaluate(context));
    final criteriaPairs = <(List<FormulaValue>, FormulaValue)>[];

    for (var i = 1; i < args.length; i += 2) {
      final range = _flattenValues(args[i].evaluate(context));
      final criteria = args[i + 1].evaluate(context);
      criteriaPairs.add((range, criteria));
    }

    num? minVal;
    for (var i = 0; i < minRange.length; i++) {
      var matchesAll = true;
      for (final (range, criteria) in criteriaPairs) {
        if (i >= range.length || !_matchesCriteria(range[i], criteria)) {
          matchesAll = false;
          break;
        }
      }
      if (matchesAll) {
        final n = minRange[i].toNumber();
        if (n != null && (minVal == null || n < minVal)) {
          minVal = n;
        }
      }
    }
    return FormulaValue.number(minVal ?? 0);
  }
}
