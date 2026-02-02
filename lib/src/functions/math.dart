import 'dart:math' as math;

import '../ast/nodes.dart';
import '../evaluation/context.dart';
import '../evaluation/errors.dart';
import '../evaluation/value.dart';
import 'function.dart';
import 'registry.dart';

/// Register all math functions.
void registerMathFunctions(FunctionRegistry registry) {
  registry.registerAll([
    SumFunction(),
    AverageFunction(),
    MinFunction(),
    MaxFunction(),
    AbsFunction(),
    RoundFunction(),
    IntFunction(),
    ModFunction(),
    SqrtFunction(),
    PowerFunction(),
    SumProductFunction(),
    RoundUpFunction(),
    RoundDownFunction(),
    CeilingFunction(),
    FloorFunction(),
    SignFunction(),
    ProductFunction(),
    RandFunction(),
    RandBetweenFunction(),
  ]);
}

/// SUM(number1, [number2], ...) - Adds all numbers.
class SumFunction extends FormulaFunction {
  @override
  String get name => 'SUM';
  @override
  int get minArgs => 0;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    var sum = 0.0;
    for (final n in collectNumbers(args, context)) {
      sum += n;
    }
    return FormulaValue.number(sum);
  }
}

/// AVERAGE(number1, [number2], ...) - Returns the average.
class AverageFunction extends FormulaFunction {
  @override
  String get name => 'AVERAGE';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = collectNumbers(args, context).toList();
    if (numbers.isEmpty) {
      return const FormulaValue.error(FormulaError.divZero);
    }
    final sum = numbers.fold(0.0, (a, b) => a + b);
    return FormulaValue.number(sum / numbers.length);
  }
}

/// MIN(number1, [number2], ...) - Returns the minimum value.
class MinFunction extends FormulaFunction {
  @override
  String get name => 'MIN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = collectNumbers(args, context).toList();
    if (numbers.isEmpty) return const FormulaValue.number(0);
    return FormulaValue.number(numbers.reduce((a, b) => a < b ? a : b));
  }
}

/// MAX(number1, [number2], ...) - Returns the maximum value.
class MaxFunction extends FormulaFunction {
  @override
  String get name => 'MAX';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = collectNumbers(args, context).toList();
    if (numbers.isEmpty) return const FormulaValue.number(0);
    return FormulaValue.number(numbers.reduce((a, b) => a > b ? a : b));
  }
}

/// ABS(number) - Returns the absolute value.
class AbsFunction extends FormulaFunction {
  @override
  String get name => 'ABS';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(n.abs());
  }
}

/// ROUND(number, num_digits) - Rounds a number.
class RoundFunction extends FormulaFunction {
  @override
  String get name => 'ROUND';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final number = values[0].toNumber();
    final digits = values[1].toNumber()?.toInt();

    if (number == null || digits == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    final multiplier = math.pow(10, digits);
    final rounded = (number * multiplier).round() / multiplier;
    return FormulaValue.number(rounded);
  }
}

/// INT(number) - Rounds down to the nearest integer.
class IntFunction extends FormulaFunction {
  @override
  String get name => 'INT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(n.floor());
  }
}

/// MOD(number, divisor) - Returns the remainder.
class ModFunction extends FormulaFunction {
  @override
  String get name => 'MOD';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final number = values[0].toNumber();
    final divisor = values[1].toNumber();

    if (number == null || divisor == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (divisor == 0) {
      return const FormulaValue.error(FormulaError.divZero);
    }
    return FormulaValue.number(number % divisor);
  }
}

/// SQRT(number) - Returns the square root.
class SqrtFunction extends FormulaFunction {
  @override
  String get name => 'SQRT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    if (n < 0) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(math.sqrt(n));
  }
}

/// POWER(number, power) - Returns number raised to a power.
class PowerFunction extends FormulaFunction {
  @override
  String get name => 'POWER';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final base = values[0].toNumber();
    final exponent = values[1].toNumber();

    if (base == null || exponent == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    return FormulaValue.number(math.pow(base, exponent));
  }
}

/// SUMPRODUCT(array1, [array2], ...) - Sum of element-wise products.
class SumProductFunction extends FormulaFunction {
  @override
  String get name => 'SUMPRODUCT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arrays = <List<List<FormulaValue>>>[];

    for (final arg in args) {
      final value = arg.evaluate(context);
      if (value.isError) return value;
      if (value is RangeValue) {
        arrays.add(value.values);
      } else {
        // Single value treated as 1x1 array
        arrays.add([
          [value]
        ]);
      }
    }

    if (arrays.isEmpty) return const FormulaValue.number(0);

    final rows = arrays[0].length;
    final cols = arrays[0].isEmpty ? 0 : arrays[0][0].length;

    // Validate all arrays have same dimensions
    for (final array in arrays) {
      if (array.length != rows) {
        return const FormulaValue.error(FormulaError.value);
      }
      for (final row in array) {
        if (row.length != cols) {
          return const FormulaValue.error(FormulaError.value);
        }
      }
    }

    var sum = 0.0;
    for (var r = 0; r < rows; r++) {
      for (var c = 0; c < cols; c++) {
        var product = 1.0;
        for (final array in arrays) {
          final n = array[r][c].toNumber();
          product *= n ?? 0;
        }
        sum += product;
      }
    }
    return FormulaValue.number(sum);
  }
}

/// ROUNDUP(number, num_digits) - Rounds away from zero.
class RoundUpFunction extends FormulaFunction {
  @override
  String get name => 'ROUNDUP';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final number = values[0].toNumber()?.toDouble();
    final digits = values[1].toNumber()?.toInt();

    if (number == null || digits == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    final multiplier = math.pow(10, digits);
    final scaled = number * multiplier;
    final rounded =
        number >= 0 ? scaled.ceilToDouble() : scaled.floorToDouble();
    return FormulaValue.number(rounded / multiplier);
  }
}

/// ROUNDDOWN(number, num_digits) - Rounds toward zero.
class RoundDownFunction extends FormulaFunction {
  @override
  String get name => 'ROUNDDOWN';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final number = values[0].toNumber()?.toDouble();
    final digits = values[1].toNumber()?.toInt();

    if (number == null || digits == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    final multiplier = math.pow(10, digits);
    final rounded = (number * multiplier).truncateToDouble();
    return FormulaValue.number(rounded / multiplier);
  }
}

/// CEILING(number, significance) - Rounds up to nearest multiple.
class CeilingFunction extends FormulaFunction {
  @override
  String get name => 'CEILING';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final number = values[0].toNumber()?.toDouble();
    final significance = values[1].toNumber()?.toDouble();

    if (number == null || significance == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (significance == 0) return const FormulaValue.number(0);
    if (number > 0 && significance < 0) {
      return const FormulaValue.error(FormulaError.num);
    }

    return FormulaValue.number((number / significance).ceil() * significance);
  }
}

/// FLOOR(number, significance) - Rounds down to nearest multiple.
class FloorFunction extends FormulaFunction {
  @override
  String get name => 'FLOOR';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final number = values[0].toNumber()?.toDouble();
    final significance = values[1].toNumber()?.toDouble();

    if (number == null || significance == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (significance == 0) return const FormulaValue.error(FormulaError.divZero);
    if (number > 0 && significance < 0) {
      return const FormulaValue.error(FormulaError.num);
    }

    return FormulaValue.number(
        (number / significance).floor() * significance);
  }
}

/// SIGN(number) - Returns -1, 0, or 1.
class SignFunction extends FormulaFunction {
  @override
  String get name => 'SIGN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(n > 0 ? 1 : (n < 0 ? -1 : 0));
  }
}

/// PRODUCT(number1, [number2], ...) - Multiplies all numbers.
class ProductFunction extends FormulaFunction {
  @override
  String get name => 'PRODUCT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    var product = 1.0;
    for (final n in collectNumbers(args, context)) {
      product *= n;
    }
    return FormulaValue.number(product);
  }
}

final math.Random _random = math.Random();

/// RAND() - Returns random number between 0 and 1.
class RandFunction extends FormulaFunction {
  @override
  String get name => 'RAND';
  @override
  int get minArgs => 0;
  @override
  int get maxArgs => 0;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    return FormulaValue.number(_random.nextDouble());
  }
}

/// RANDBETWEEN(bottom, top) - Returns random integer in range.
class RandBetweenFunction extends FormulaFunction {
  @override
  String get name => 'RANDBETWEEN';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final bottom = values[0].toNumber()?.toInt();
    final top = values[1].toNumber()?.toInt();

    if (bottom == null || top == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (bottom > top) {
      return const FormulaValue.error(FormulaError.num);
    }

    return FormulaValue.number(bottom + _random.nextInt(top - bottom + 1));
  }
}
