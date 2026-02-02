import 'dart:math' as math;

import 'package:a1/a1.dart';

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
    PiFunction(),
    LnFunction(),
    LogFunction(),
    Log10Function(),
    ExpFunction(),
    SinFunction(),
    CosFunction(),
    TanFunction(),
    AsinFunction(),
    AcosFunction(),
    AtanFunction(),
    Atan2Function(),
    DegreesFunction(),
    RadiansFunction(),
    EvenFunction(),
    OddFunction(),
    GcdFunction(),
    LcmFunction(),
    TruncFunction(),
    MroundFunction(),
    QuotientFunction(),
    CombinFunction(),
    CombinaFunction(),
    FactFunction(),
    FactDoubleFunction(),
    SumSqFunction(),
    SubtotalFunction(),
    AggregateFunction(),
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

// --- Shared helpers ---

int _gcd(int a, int b) {
  a = a.abs();
  b = b.abs();
  while (b != 0) {
    final t = b;
    b = a % b;
    a = t;
  }
  return a;
}

double _factorial(int n) {
  if (n < 0) return double.nan;
  var result = 1.0;
  for (var i = 2; i <= n; i++) {
    result *= i;
  }
  return result;
}

/// PI() - Returns the mathematical constant pi.
class PiFunction extends FormulaFunction {
  @override
  String get name => 'PI';
  @override
  int get minArgs => 0;
  @override
  int get maxArgs => 0;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    return const FormulaValue.number(math.pi);
  }
}

/// LN(number) - Returns the natural logarithm.
class LnFunction extends FormulaFunction {
  @override
  String get name => 'LN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    if (n <= 0) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(math.log(n));
  }
}

/// LOG(number, [base]) - Returns the logarithm with specified base.
class LogFunction extends FormulaFunction {
  @override
  String get name => 'LOG';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final n = values[0].toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    if (n <= 0) return const FormulaValue.error(FormulaError.num);

    final base = args.length > 1 ? values[1].toNumber() : 10;
    if (base == null) return const FormulaValue.error(FormulaError.value);
    if (base <= 0) return const FormulaValue.error(FormulaError.num);
    if (base == 1) return const FormulaValue.error(FormulaError.divZero);

    return FormulaValue.number(math.log(n) / math.log(base));
  }
}

/// LOG10(number) - Returns the base-10 logarithm.
class Log10Function extends FormulaFunction {
  @override
  String get name => 'LOG10';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    if (n <= 0) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(math.log(n) / math.ln10);
  }
}

/// EXP(number) - Returns e raised to a power.
class ExpFunction extends FormulaFunction {
  @override
  String get name => 'EXP';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(math.exp(n));
  }
}

/// SIN(number) - Returns the sine (radians).
class SinFunction extends FormulaFunction {
  @override
  String get name => 'SIN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(math.sin(n));
  }
}

/// COS(number) - Returns the cosine (radians).
class CosFunction extends FormulaFunction {
  @override
  String get name => 'COS';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(math.cos(n));
  }
}

/// TAN(number) - Returns the tangent (radians).
class TanFunction extends FormulaFunction {
  @override
  String get name => 'TAN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(math.tan(n));
  }
}

/// ASIN(number) - Returns the arcsine.
class AsinFunction extends FormulaFunction {
  @override
  String get name => 'ASIN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    if (n.abs() > 1) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(math.asin(n));
  }
}

/// ACOS(number) - Returns the arccosine.
class AcosFunction extends FormulaFunction {
  @override
  String get name => 'ACOS';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    if (n.abs() > 1) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(math.acos(n));
  }
}

/// ATAN(number) - Returns the arctangent.
class AtanFunction extends FormulaFunction {
  @override
  String get name => 'ATAN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(math.atan(n));
  }
}

/// ATAN2(x_num, y_num) - Returns the arctangent from x and y coordinates.
class Atan2Function extends FormulaFunction {
  @override
  String get name => 'ATAN2';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final x = values[0].toNumber();
    final y = values[1].toNumber();
    if (x == null || y == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (x == 0 && y == 0) {
      return const FormulaValue.error(FormulaError.divZero);
    }
    // Excel ATAN2(x_num, y_num) is math.atan2(y_num, x_num)
    return FormulaValue.number(math.atan2(y, x));
  }
}

/// DEGREES(angle) - Converts radians to degrees.
class DegreesFunction extends FormulaFunction {
  @override
  String get name => 'DEGREES';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(n * 180 / math.pi);
  }
}

/// RADIANS(angle) - Converts degrees to radians.
class RadiansFunction extends FormulaFunction {
  @override
  String get name => 'RADIANS';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(n * math.pi / 180);
  }
}

/// EVEN(number) - Rounds away from zero to the nearest even integer.
class EvenFunction extends FormulaFunction {
  @override
  String get name => 'EVEN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber()?.toDouble();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    if (n == 0) return const FormulaValue.number(0);
    int result;
    if (n > 0) {
      result = n.ceil();
      if (result.isOdd) result++;
    } else {
      result = n.floor();
      if (result.isOdd) result--;
    }
    return FormulaValue.number(result);
  }
}

/// ODD(number) - Rounds away from zero to the nearest odd integer.
class OddFunction extends FormulaFunction {
  @override
  String get name => 'ODD';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber()?.toDouble();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    if (n == 0) return const FormulaValue.number(1);
    int result;
    if (n > 0) {
      result = n.ceil();
      if (result.isEven) result++;
    } else {
      result = n.floor();
      if (result.isEven) result--;
    }
    return FormulaValue.number(result);
  }
}

/// GCD(number1, [number2], ...) - Returns the greatest common divisor.
class GcdFunction extends FormulaFunction {
  @override
  String get name => 'GCD';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = collectNumbers(args, context).toList();
    if (numbers.isEmpty) return const FormulaValue.error(FormulaError.value);
    var result = numbers[0].truncate().abs();
    for (var i = 1; i < numbers.length; i++) {
      result = _gcd(result, numbers[i].truncate().abs());
    }
    return FormulaValue.number(result);
  }
}

/// LCM(number1, [number2], ...) - Returns the least common multiple.
class LcmFunction extends FormulaFunction {
  @override
  String get name => 'LCM';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = collectNumbers(args, context).toList();
    if (numbers.isEmpty) return const FormulaValue.error(FormulaError.value);
    var result = numbers[0].truncate().abs();
    for (var i = 1; i < numbers.length; i++) {
      final b = numbers[i].truncate().abs();
      if (result == 0 || b == 0) return const FormulaValue.number(0);
      result = (result * b) ~/ _gcd(result, b);
    }
    return FormulaValue.number(result);
  }
}

/// TRUNC(number, [num_digits]) - Truncates toward zero.
class TruncFunction extends FormulaFunction {
  @override
  String get name => 'TRUNC';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final number = values[0].toNumber()?.toDouble();
    if (number == null) return const FormulaValue.error(FormulaError.value);
    final digits = args.length > 1 ? values[1].toNumber()?.toInt() ?? 0 : 0;

    final multiplier = math.pow(10, digits);
    return FormulaValue.number(
        (number * multiplier).truncateToDouble() / multiplier);
  }
}

/// MROUND(number, multiple) - Rounds to nearest multiple.
class MroundFunction extends FormulaFunction {
  @override
  String get name => 'MROUND';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final number = values[0].toNumber()?.toDouble();
    final multiple = values[1].toNumber()?.toDouble();
    if (number == null || multiple == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (multiple == 0) return const FormulaValue.number(0);
    if (number > 0 && multiple < 0 || number < 0 && multiple > 0) {
      return const FormulaValue.error(FormulaError.num);
    }
    return FormulaValue.number((number / multiple).round() * multiple);
  }
}

/// QUOTIENT(numerator, denominator) - Returns the integer portion of a division.
class QuotientFunction extends FormulaFunction {
  @override
  String get name => 'QUOTIENT';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final numerator = values[0].toNumber();
    final denominator = values[1].toNumber();
    if (numerator == null || denominator == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (denominator == 0) {
      return const FormulaValue.error(FormulaError.divZero);
    }
    return FormulaValue.number((numerator / denominator).truncate());
  }
}

/// COMBIN(n, k) - Returns the number of combinations.
class CombinFunction extends FormulaFunction {
  @override
  String get name => 'COMBIN';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final n = values[0].toNumber()?.truncate();
    final k = values[1].toNumber()?.truncate();
    if (n == null || k == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (n < 0 || k < 0 || k > n) {
      return const FormulaValue.error(FormulaError.num);
    }
    // Use iterative product to avoid overflow: C(n,k) = n*(n-1)*...*(n-k+1) / k!
    var result = 1.0;
    final kk = k > n - k ? n - k : k; // optimize: use smaller k
    for (var i = 0; i < kk; i++) {
      result = result * (n - i) / (i + 1);
    }
    return FormulaValue.number(result.round());
  }
}

/// COMBINA(n, k) - Returns the number of combinations with repetitions.
class CombinaFunction extends FormulaFunction {
  @override
  String get name => 'COMBINA';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final n = values[0].toNumber()?.truncate();
    final k = values[1].toNumber()?.truncate();
    if (n == null || k == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (n < 0 || k < 0) {
      return const FormulaValue.error(FormulaError.num);
    }
    if (n == 0 && k == 0) return const FormulaValue.number(1);
    // COMBINA(n, k) = COMBIN(n+k-1, k)
    final nn = n + k - 1;
    var result = 1.0;
    final kk = k > nn - k ? nn - k : k;
    for (var i = 0; i < kk; i++) {
      result = result * (nn - i) / (i + 1);
    }
    return FormulaValue.number(result.round());
  }
}

/// FACT(number) - Returns the factorial.
class FactFunction extends FormulaFunction {
  @override
  String get name => 'FACT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber()?.truncate();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    if (n < 0) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(_factorial(n));
  }
}

/// FACTDOUBLE(number) - Returns the double factorial.
class FactDoubleFunction extends FormulaFunction {
  @override
  String get name => 'FACTDOUBLE';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber()?.truncate();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    if (n < -1) return const FormulaValue.error(FormulaError.num);
    if (n <= 0) return const FormulaValue.number(1);
    var result = 1.0;
    for (var i = n; i > 0; i -= 2) {
      result *= i;
    }
    return FormulaValue.number(result);
  }
}

/// SUMSQ(number1, [number2], ...) - Returns the sum of squares.
class SumSqFunction extends FormulaFunction {
  @override
  String get name => 'SUMSQ';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    var sum = 0.0;
    for (final n in collectNumbers(args, context)) {
      sum += n * n;
    }
    return FormulaValue.number(sum);
  }
}

/// SUBTOTAL(function_num, ref1, [ref2], ...) - Applies a function to data.
class SubtotalFunction extends FormulaFunction {
  @override
  String get name => 'SUBTOTAL';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => -1;
  @override
  bool get isLazy => true;

  static const _functionMap = {
    1: 'AVERAGE',
    2: 'COUNT',
    3: 'COUNTA',
    4: 'MAX',
    5: 'MIN',
    6: 'PRODUCT',
    7: 'STDEV.S',
    8: 'STDEV.P',
    9: 'SUM',
    10: 'VAR.S',
    11: 'VAR.P',
    101: 'AVERAGE',
    102: 'COUNT',
    103: 'COUNTA',
    104: 'MAX',
    105: 'MIN',
    106: 'PRODUCT',
    107: 'STDEV.S',
    108: 'STDEV.P',
    109: 'SUM',
    110: 'VAR.S',
    111: 'VAR.P',
  };

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final funcNumValue = args[0].evaluate(context);
    final funcNum = funcNumValue.toNumber()?.toInt();
    if (funcNum == null) return const FormulaValue.error(FormulaError.value);

    final funcName = _functionMap[funcNum];
    if (funcName == null) return const FormulaValue.error(FormulaError.value);

    final func = context.getFunction(funcName);
    if (func == null) return const FormulaValue.error(FormulaError.value);

    return func.call(args.sublist(1), context);
  }
}

/// AGGREGATE(function_num, options, ref1, [ref2], ...) - Applies a function
/// with options to ignore errors.
class AggregateFunction extends FormulaFunction {
  @override
  String get name => 'AGGREGATE';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => -1;
  @override
  bool get isLazy => true;

  static const _functionMap = {
    1: 'AVERAGE',
    2: 'COUNT',
    3: 'COUNTA',
    4: 'MAX',
    5: 'MIN',
    6: 'PRODUCT',
    7: 'STDEV.S',
    8: 'STDEV.P',
    9: 'SUM',
    10: 'VAR.S',
    11: 'VAR.P',
    12: 'MEDIAN',
    13: 'MODE.SNGL',
    14: 'LARGE',
    15: 'SMALL',
    16: 'PERCENTILE.INC',
    17: 'QUARTILE.INC',
    18: 'PERCENTILE.EXC',
    19: 'QUARTILE.EXC',
  };

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final funcNumValue = args[0].evaluate(context);
    final funcNum = funcNumValue.toNumber()?.toInt();
    if (funcNum == null) return const FormulaValue.error(FormulaError.value);

    final optionsValue = args[1].evaluate(context);
    final options = optionsValue.toNumber()?.toInt() ?? 0;

    final funcName = _functionMap[funcNum];
    if (funcName == null) return const FormulaValue.error(FormulaError.value);

    final func = context.getFunction(funcName);
    if (func == null) return const FormulaValue.error(FormulaError.value);

    final dataArgs = args.sublist(2);

    // Options 2, 3, 6, 7 include "ignore error values"
    final ignoreErrors =
        options == 2 || options == 3 || options == 6 || options == 7;

    if (ignoreErrors) {
      return func.call(dataArgs, _ErrorFilteringContext(context));
    }

    return func.call(dataArgs, context);
  }
}

/// A context wrapper that replaces error values with empty values.
class _ErrorFilteringContext implements EvaluationContext {
  final EvaluationContext _inner;
  const _ErrorFilteringContext(this._inner);

  @override
  A1 get currentCell => _inner.currentCell;
  @override
  String? get currentSheet => _inner.currentSheet;
  @override
  bool get isCancelled => _inner.isCancelled;
  @override
  FormulaFunction? getFunction(String name) => _inner.getFunction(name);

  @override
  FormulaValue getCellValue(A1 cell) {
    final value = _inner.getCellValue(cell);
    return value.isError ? const EmptyValue() : value;
  }

  @override
  FormulaValue getRangeValues(A1Range range) {
    final value = _inner.getRangeValues(range);
    if (value is RangeValue) {
      return RangeValue(value.values
          .map((row) =>
              row.map((c) => c.isError ? const EmptyValue() : c).toList())
          .toList());
    }
    return value.isError ? const EmptyValue() : value;
  }
}
