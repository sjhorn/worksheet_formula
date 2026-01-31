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
