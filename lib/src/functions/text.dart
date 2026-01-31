import 'dart:math' as math;

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

    final format = values[1].toText();
    return FormulaValue.text(_formatNumber(number.toDouble(), format));
  }
}

/// Format a number according to an Excel-style format code.
String _formatNumber(double number, String format) {
  // Check for percentage
  final isPercent = format.contains('%');
  if (isPercent) {
    number = number * 100;
    format = format.replaceAll('%', '');
  }

  // Check for scientific notation
  final sciMatch = RegExp(r'(.*?)E([+-])(.*)$', caseSensitive: false)
      .firstMatch(format);
  if (sciMatch != null) {
    final result = _formatScientific(number, sciMatch);
    return isPercent ? '$result%' : result;
  }

  // Check for thousands separator
  final useThousands = format.contains(',');
  format = format.replaceAll(',', '');

  // Split into integer and decimal format parts
  final parts = format.split('.');
  final intFormat = parts[0];
  final decFormat = parts.length > 1 ? parts[1] : null;

  // Determine decimal places
  final decimalPlaces = decFormat?.length ?? 0;

  // Round number
  String numStr;
  if (decimalPlaces > 0) {
    numStr = number.toStringAsFixed(decimalPlaces);
  } else {
    numStr = number.round().toString();
  }

  // Handle negative
  final isNegative = numStr.startsWith('-');
  if (isNegative) numStr = numStr.substring(1);

  // Split result into integer and decimal parts
  final numParts = numStr.split('.');
  var intPart = numParts[0];
  var decPart = numParts.length > 1 ? numParts[1] : '';

  // Strip trailing zeros from decimal part when format uses '#'
  if (decFormat != null) {
    // Count required decimal digits (0s) from the right
    final minDecDigits = decFormat.replaceAll('#', '').length;
    while (decPart.length > minDecDigits && decPart.endsWith('0')) {
      decPart = decPart.substring(0, decPart.length - 1);
    }
  }

  // Zero-pad integer part (count '0' chars in intFormat)
  final minIntDigits = intFormat.replaceAll('#', '').length;
  while (intPart.length < minIntDigits) {
    intPart = '0$intPart';
  }

  // Strip leading zeros for '#' format
  if (intFormat.isNotEmpty && intFormat[0] == '#') {
    intPart = intPart.replaceFirst(RegExp(r'^0+'), '');
  }

  // Add thousands separators
  if (useThousands && intPart.length > 3) {
    final buffer = StringBuffer();
    for (var i = 0; i < intPart.length; i++) {
      if (i > 0 && (intPart.length - i) % 3 == 0) {
        buffer.write(',');
      }
      buffer.write(intPart[i]);
    }
    intPart = buffer.toString();
  }

  // Assemble result
  var result = isNegative ? '-$intPart' : intPart;
  if (decFormat != null) result += '.$decPart';
  if (isPercent) result += '%';
  return result;
}

/// Format a number in scientific notation.
String _formatScientific(double number, RegExpMatch sciMatch) {
  final mantissaFormat = sciMatch.group(1)!;
  final sign = sciMatch.group(2)!;
  final expFormat = sciMatch.group(3)!;

  if (number == 0) {
    final decPlaces = mantissaFormat.contains('.')
        ? mantissaFormat.split('.')[1].length
        : 0;
    final mantissa = decPlaces > 0 ? '0.${'0' * decPlaces}' : '0';
    return '${mantissa}E${sign}0';
  }

  // Calculate exponent
  final exp = (math.log(number.abs()) / math.ln10).floor();
  final mantissa = number / _pow10(exp);

  // Format mantissa
  final decPlaces = mantissaFormat.contains('.')
      ? mantissaFormat.split('.')[1].length
      : 0;
  final mantissaStr = mantissa.toStringAsFixed(decPlaces);

  // Format exponent
  final absExp = exp.abs();
  var expStr = absExp.toString();
  final minExpDigits = expFormat.replaceAll('#', '').length;
  while (expStr.length < minExpDigits) {
    expStr = '0$expStr';
  }

  final expSign = exp >= 0 ? '+' : '-';
  return '${mantissaStr}E$expSign$expStr';
}

double _pow10(int exp) {
  var result = 1.0;
  final absExp = exp.abs();
  for (var i = 0; i < absExp; i++) {
    result *= 10;
  }
  return exp >= 0 ? result : 1 / result;
}
