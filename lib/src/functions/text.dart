import '../ast/nodes.dart';
import '../evaluation/context.dart';
import '../evaluation/errors.dart';
import '../evaluation/value.dart';
import 'function.dart';
import 'registry.dart';

/// Register all text functions.
void registerTextFunctions(FunctionRegistry registry) {
  registry.registerAll([
    ConcatFunction(),
    ConcatenateFunction(),
    LeftFunction(),
    RightFunction(),
    MidFunction(),
    LenFunction(),
    LowerFunction(),
    UpperFunction(),
    TrimFunction(),
    TextFunction(),
  ]);
}

/// CONCAT(text1, [text2], ...) - Joins text strings.
class ConcatFunction extends FormulaFunction {
  @override
  String get name => 'CONCAT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final buffer = StringBuffer();
    for (final arg in args) {
      final value = arg.evaluate(context);
      if (value.isError) return value;
      buffer.write(value.toText());
    }
    return FormulaValue.text(buffer.toString());
  }
}

/// CONCATENATE(text1, [text2], ...) - Legacy version of CONCAT.
class ConcatenateFunction extends ConcatFunction {
  @override
  String get name => 'CONCATENATE';
}

/// LEFT(text, [num_chars]) - Returns leftmost characters.
class LeftFunction extends FormulaFunction {
  @override
  String get name => 'LEFT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final text = values[0].toText();
    final numChars =
        args.length > 1 ? values[1].toNumber()?.toInt() ?? 1 : 1;

    if (numChars < 0) return const FormulaValue.error(FormulaError.value);

    final end = numChars > text.length ? text.length : numChars;
    return FormulaValue.text(text.substring(0, end));
  }
}

/// RIGHT(text, [num_chars]) - Returns rightmost characters.
class RightFunction extends FormulaFunction {
  @override
  String get name => 'RIGHT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final text = values[0].toText();
    final numChars =
        args.length > 1 ? values[1].toNumber()?.toInt() ?? 1 : 1;

    if (numChars < 0) return const FormulaValue.error(FormulaError.value);

    final start = text.length - numChars;
    return FormulaValue.text(text.substring(start < 0 ? 0 : start));
  }
}

/// MID(text, start_num, num_chars) - Returns characters from middle.
class MidFunction extends FormulaFunction {
  @override
  String get name => 'MID';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final text = values[0].toText();
    final startNum = values[1].toNumber()?.toInt();
    final numChars = values[2].toNumber()?.toInt();

    if (startNum == null || numChars == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (startNum < 1 || numChars < 0) {
      return const FormulaValue.error(FormulaError.value);
    }

    final start = startNum - 1; // Excel is 1-indexed
    if (start >= text.length) return const FormulaValue.text('');

    final end = start + numChars;
    return FormulaValue.text(
      text.substring(start, end > text.length ? text.length : end),
    );
  }
}

/// LEN(text) - Returns length of text.
class LenFunction extends FormulaFunction {
  @override
  String get name => 'LEN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    return FormulaValue.number(value.toText().length);
  }
}

/// LOWER(text) - Converts to lowercase.
class LowerFunction extends FormulaFunction {
  @override
  String get name => 'LOWER';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    return FormulaValue.text(value.toText().toLowerCase());
  }
}

/// UPPER(text) - Converts to uppercase.
class UpperFunction extends FormulaFunction {
  @override
  String get name => 'UPPER';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    return FormulaValue.text(value.toText().toUpperCase());
  }
}

/// TRIM(text) - Removes leading/trailing spaces and collapses internal spaces.
class TrimFunction extends FormulaFunction {
  @override
  String get name => 'TRIM';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final text = value.toText().trim().replaceAll(RegExp(r'\s+'), ' ');
    return FormulaValue.text(text);
  }
}

/// TEXT(value, format_text) - Formats a number as text.
class TextFunction extends FormulaFunction {
  @override
  String get name => 'TEXT';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final number = values[0].toNumber();

    if (number == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    // Simplified format handling - returns number as string
    return FormulaValue.text(number.toString());
  }
}
