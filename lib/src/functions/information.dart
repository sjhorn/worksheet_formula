import '../ast/nodes.dart';
import '../evaluation/context.dart';
import '../evaluation/errors.dart';
import '../evaluation/value.dart';
import 'function.dart';
import 'registry.dart';

/// Register all information functions.
void registerInformationFunctions(FunctionRegistry registry) {
  registry.registerAll([
    IsBlankFunction(),
    IsErrorFunction(),
    IsNumberFunction(),
    IsTextFunction(),
    IsLogicalFunction(),
    IsNaFunction(),
    TypeFunction(),
  ]);
}

/// ISBLANK(value) - Returns TRUE if value is empty.
class IsBlankFunction extends FormulaFunction {
  @override
  String get name => 'ISBLANK';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    return FormulaValue.boolean(value is EmptyValue);
  }
}

/// ISERROR(value) - Returns TRUE if value is any error.
class IsErrorFunction extends FormulaFunction {
  @override
  String get name => 'ISERROR';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    return FormulaValue.boolean(value is ErrorValue);
  }
}

/// ISNUMBER(value) - Returns TRUE if value is a number.
class IsNumberFunction extends FormulaFunction {
  @override
  String get name => 'ISNUMBER';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    return FormulaValue.boolean(value is NumberValue);
  }
}

/// ISTEXT(value) - Returns TRUE if value is text.
class IsTextFunction extends FormulaFunction {
  @override
  String get name => 'ISTEXT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    return FormulaValue.boolean(value is TextValue);
  }
}

/// ISLOGICAL(value) - Returns TRUE if value is a boolean.
class IsLogicalFunction extends FormulaFunction {
  @override
  String get name => 'ISLOGICAL';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    return FormulaValue.boolean(value is BooleanValue);
  }
}

/// ISNA(value) - Returns TRUE if value is #N/A.
class IsNaFunction extends FormulaFunction {
  @override
  String get name => 'ISNA';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    return FormulaValue.boolean(
      value is ErrorValue && value.error == FormulaError.na,
    );
  }
}

/// TYPE(value) - Returns the type of a value.
///
/// Returns: 1=number, 2=text, 4=boolean, 16=error, 64=array.
class TypeFunction extends FormulaFunction {
  @override
  String get name => 'TYPE';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    return switch (value) {
      NumberValue() => const FormulaValue.number(1),
      TextValue() => const FormulaValue.number(2),
      BooleanValue() => const FormulaValue.number(4),
      ErrorValue() => const FormulaValue.number(16),
      RangeValue() => const FormulaValue.number(64),
      EmptyValue() => const FormulaValue.number(1), // Empty treated as number (0)
    };
  }
}
