import '../ast/nodes.dart';
import '../evaluation/context.dart';
import '../evaluation/errors.dart';
import '../evaluation/value.dart';
import 'function.dart';
import 'registry.dart';

/// Register all logical functions.
void registerLogicalFunctions(FunctionRegistry registry) {
  registry.registerAll([
    IfFunction(),
    AndFunction(),
    OrFunction(),
    NotFunction(),
    IfErrorFunction(),
    IfNaFunction(),
    TrueFunction(),
    FalseFunction(),
    IfsFunction(),
    SwitchFunction(),
    XorFunction(),
  ]);
}

/// IF(logical_test, value_if_true, [value_if_false])
class IfFunction extends FormulaFunction {
  @override
  String get name => 'IF';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final condition = args[0].evaluate(context);
    if (condition.isError) return condition;

    if (condition.isTruthy) {
      return args[1].evaluate(context);
    } else if (args.length > 2) {
      return args[2].evaluate(context);
    } else {
      return const FormulaValue.boolean(false);
    }
  }
}

/// AND(logical1, [logical2], ...) - Returns TRUE if all arguments are TRUE.
class AndFunction extends FormulaFunction {
  @override
  String get name => 'AND';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    for (final arg in args) {
      final value = arg.evaluate(context);
      if (value.isError) return value;
      if (!value.isTruthy) return const FormulaValue.boolean(false);
    }
    return const FormulaValue.boolean(true);
  }
}

/// OR(logical1, [logical2], ...) - Returns TRUE if any argument is TRUE.
class OrFunction extends FormulaFunction {
  @override
  String get name => 'OR';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    for (final arg in args) {
      final value = arg.evaluate(context);
      if (value.isError) return value;
      if (value.isTruthy) return const FormulaValue.boolean(true);
    }
    return const FormulaValue.boolean(false);
  }
}

/// NOT(logical) - Reverses the logic of its argument.
class NotFunction extends FormulaFunction {
  @override
  String get name => 'NOT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    if (value.isError) return value;
    return FormulaValue.boolean(!value.isTruthy);
  }
}

/// IFERROR(value, value_if_error) - Returns value_if_error if value is an error.
class IfErrorFunction extends FormulaFunction {
  @override
  String get name => 'IFERROR';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    if (value.isError) return args[1].evaluate(context);
    return value;
  }
}

/// IFNA(value, value_if_na) - Returns value_if_na if value is #N/A.
class IfNaFunction extends FormulaFunction {
  @override
  String get name => 'IFNA';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    if (value case ErrorValue(error: FormulaError.na)) {
      return args[1].evaluate(context);
    }
    return value;
  }
}

/// TRUE() - Returns the logical value TRUE.
class TrueFunction extends FormulaFunction {
  @override
  String get name => 'TRUE';
  @override
  int get minArgs => 0;
  @override
  int get maxArgs => 0;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) =>
      const FormulaValue.boolean(true);
}

/// FALSE() - Returns the logical value FALSE.
class FalseFunction extends FormulaFunction {
  @override
  String get name => 'FALSE';
  @override
  int get minArgs => 0;
  @override
  int get maxArgs => 0;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) =>
      const FormulaValue.boolean(false);
}

/// IFS(condition1, value1, [condition2, value2], ...) - Chained IF.
class IfsFunction extends FormulaFunction {
  @override
  String get name => 'IFS';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => -1;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    if (args.length.isOdd) {
      return const FormulaValue.error(FormulaError.na);
    }
    for (var i = 0; i < args.length; i += 2) {
      final condition = args[i].evaluate(context);
      if (condition.isError) return condition;
      if (condition.isTruthy) {
        return args[i + 1].evaluate(context);
      }
    }
    return const FormulaValue.error(FormulaError.na);
  }
}

/// SWITCH(expression, val1, result1, ..., [default]) - Match against cases.
class SwitchFunction extends FormulaFunction {
  @override
  String get name => 'SWITCH';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => -1;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final expression = args[0].evaluate(context);
    if (expression.isError) return expression;

    // Remaining args after expression: pairs of (value, result), optional default
    final remaining = args.length - 1;
    final hasDefault = remaining.isOdd;
    final pairCount = hasDefault ? (remaining - 1) ~/ 2 : remaining ~/ 2;

    for (var i = 0; i < pairCount; i++) {
      final caseValue = args[1 + i * 2].evaluate(context);
      if (_switchValuesEqual(expression, caseValue)) {
        return args[2 + i * 2].evaluate(context);
      }
    }

    if (hasDefault) {
      return args.last.evaluate(context);
    }
    return const FormulaValue.error(FormulaError.na);
  }

  bool _switchValuesEqual(FormulaValue a, FormulaValue b) {
    final aNum = a.toNumber();
    final bNum = b.toNumber();
    if (aNum != null && bNum != null) return aNum == bNum;
    return a.toText().toLowerCase() == b.toText().toLowerCase();
  }
}

/// XOR(logical1, [logical2], ...) - Exclusive OR.
class XorFunction extends FormulaFunction {
  @override
  String get name => 'XOR';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    var trueCount = 0;
    for (final arg in args) {
      final value = arg.evaluate(context);
      if (value.isError) return value;
      if (value.isTruthy) trueCount++;
    }
    return FormulaValue.boolean(trueCount.isOdd);
  }
}
