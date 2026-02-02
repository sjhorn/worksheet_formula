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
    IsErrFunction(),
    IsNonTextFunction(),
    IsEvenFunction(),
    IsOddFunction(),
    IsRefFunction(),
    NFunction(),
    NaFunction(),
    ErrorTypeFunction(),
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

/// ISERR(value) - Returns TRUE if value is an error other than #N/A.
class IsErrFunction extends FormulaFunction {
  @override
  String get name => 'ISERR';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    return FormulaValue.boolean(
      value is ErrorValue && value.error != FormulaError.na,
    );
  }
}

/// ISNONTEXT(value) - Returns TRUE if value is not text.
class IsNonTextFunction extends FormulaFunction {
  @override
  String get name => 'ISNONTEXT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    return FormulaValue.boolean(value is! TextValue);
  }
}

/// ISEVEN(number) - Returns TRUE if number is even.
class IsEvenFunction extends FormulaFunction {
  @override
  String get name => 'ISEVEN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.boolean(n.truncate() % 2 == 0);
  }
}

/// ISODD(number) - Returns TRUE if number is odd.
class IsOddFunction extends FormulaFunction {
  @override
  String get name => 'ISODD';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.boolean(n.truncate() % 2 != 0);
  }
}

/// ISREF(value) - Returns TRUE if value is a cell or range reference.
class IsRefFunction extends FormulaFunction {
  @override
  String get name => 'ISREF';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;
  @override
  bool get isLazy => true;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    return FormulaValue.boolean(
      args[0] is CellRefNode || args[0] is RangeRefNode,
    );
  }
}

/// N(value) - Converts a value to a number.
class NFunction extends FormulaFunction {
  @override
  String get name => 'N';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    return switch (value) {
      NumberValue() => value,
      BooleanValue() => FormulaValue.number(value.value ? 1 : 0),
      ErrorValue() => value,
      _ => const FormulaValue.number(0),
    };
  }
}

/// NA() - Returns the #N/A error value.
class NaFunction extends FormulaFunction {
  @override
  String get name => 'NA';
  @override
  int get minArgs => 0;
  @override
  int get maxArgs => 0;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    return const FormulaValue.error(FormulaError.na);
  }
}

/// ERROR.TYPE(error_val) - Returns a number corresponding to an error type.
class ErrorTypeFunction extends FormulaFunction {
  @override
  String get name => 'ERROR.TYPE';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    if (value is! ErrorValue) {
      return const FormulaValue.error(FormulaError.na);
    }
    return switch (value.error) {
      FormulaError.null_ => const FormulaValue.number(1),
      FormulaError.divZero => const FormulaValue.number(2),
      FormulaError.value => const FormulaValue.number(3),
      FormulaError.ref => const FormulaValue.number(4),
      FormulaError.name => const FormulaValue.number(5),
      FormulaError.num => const FormulaValue.number(6),
      FormulaError.na => const FormulaValue.number(7),
      _ => const FormulaValue.error(FormulaError.na),
    };
  }
}
