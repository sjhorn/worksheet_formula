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
    FindFunction(),
    SearchFunction(),
    SubstituteFunction(),
    ReplaceFunction(),
    ValueFunction(),
    TextJoinFunction(),
    ProperFunction(),
    ExactFunction(),
    ReptFunction(),
    CharFunction(),
    CodeFunction(),
    CleanFunction(),
    DollarFunction(),
    FixedFunction(),
    TFunction(),
    NumberValueFunction(),
    UnicharFunction(),
    UnicodeFunction(),
    TextBeforeFunction(),
    TextAfterFunction(),
    TextSplitFunction(),
    ArrayToTextFunction(),
    ValueToTextFunction(),
    AscFunction(),
    DbcsFunction(),
    BahtTextFunction(),
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

/// FIND(find_text, within_text, [start_num]) - Case-sensitive search.
class FindFunction extends FormulaFunction {
  @override
  String get name => 'FIND';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final findText = values[0].toText();
    final withinText = values[1].toText();
    final startNum = args.length > 2 ? values[2].toNumber()?.toInt() ?? 1 : 1;

    if (startNum < 1) return const FormulaValue.error(FormulaError.value);

    final startIndex = startNum - 1;
    if (startIndex >= withinText.length && findText.isNotEmpty) {
      return const FormulaValue.error(FormulaError.value);
    }

    final pos = withinText.indexOf(findText, startIndex);
    if (pos == -1) return const FormulaValue.error(FormulaError.value);

    return FormulaValue.number(pos + 1); // 1-indexed
  }
}

/// SEARCH(find_text, within_text, [start_num]) - Case-insensitive search with wildcards.
class SearchFunction extends FormulaFunction {
  @override
  String get name => 'SEARCH';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final findText = values[0].toText();
    final withinText = values[1].toText();
    final startNum = args.length > 2 ? values[2].toNumber()?.toInt() ?? 1 : 1;

    if (startNum < 1) return const FormulaValue.error(FormulaError.value);

    final startIndex = startNum - 1;
    if (startIndex >= withinText.length && findText.isNotEmpty) {
      return const FormulaValue.error(FormulaError.value);
    }

    // Convert wildcards to regex: ? -> . and * -> .*
    final pattern = StringBuffer();
    for (var i = 0; i < findText.length; i++) {
      final ch = findText[i];
      if (ch == '~' && i + 1 < findText.length) {
        // Escape sequence: ~? ~* ~~ are literal
        pattern.write(RegExp.escape(findText[i + 1]));
        i++;
      } else if (ch == '?') {
        pattern.write('.');
      } else if (ch == '*') {
        pattern.write('.*');
      } else {
        pattern.write(RegExp.escape(ch));
      }
    }

    final regex = RegExp(pattern.toString(), caseSensitive: false);
    final searchIn = withinText.substring(startIndex);
    final match = regex.firstMatch(searchIn);

    if (match == null) return const FormulaValue.error(FormulaError.value);

    return FormulaValue.number(startIndex + match.start + 1); // 1-indexed
  }
}

/// SUBSTITUTE(text, old_text, new_text, [instance_num]) - Replace text.
class SubstituteFunction extends FormulaFunction {
  @override
  String get name => 'SUBSTITUTE';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final text = values[0].toText();
    final oldText = values[1].toText();
    final newText = values[2].toText();
    final instanceNum =
        args.length > 3 ? values[3].toNumber()?.toInt() : null;

    if (instanceNum != null && instanceNum < 1) {
      return const FormulaValue.error(FormulaError.value);
    }

    if (oldText.isEmpty) return FormulaValue.text(text);

    if (instanceNum == null) {
      // Replace all occurrences
      return FormulaValue.text(text.replaceAll(oldText, newText));
    }

    // Replace only the nth occurrence
    var count = 0;
    var startIndex = 0;
    while (startIndex < text.length) {
      final pos = text.indexOf(oldText, startIndex);
      if (pos == -1) break;
      count++;
      if (count == instanceNum) {
        final result = text.substring(0, pos) +
            newText +
            text.substring(pos + oldText.length);
        return FormulaValue.text(result);
      }
      startIndex = pos + oldText.length;
    }

    return FormulaValue.text(text); // Instance not found
  }
}

/// REPLACE(old_text, start_num, num_chars, new_text) - Replace by position.
class ReplaceFunction extends FormulaFunction {
  @override
  String get name => 'REPLACE';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final text = values[0].toText();
    final startNum = values[1].toNumber()?.toInt();
    final numChars = values[2].toNumber()?.toInt();
    final newText = values[3].toText();

    if (startNum == null || numChars == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (startNum < 1 || numChars < 0) {
      return const FormulaValue.error(FormulaError.value);
    }

    final start = startNum - 1; // 1-indexed to 0-indexed
    final end = start + numChars;
    final before = text.substring(0, start > text.length ? text.length : start);
    final after =
        end >= text.length ? '' : text.substring(end);

    return FormulaValue.text('$before$newText$after');
  }
}

/// VALUE(text) - Convert text to number.
class ValueFunction extends FormulaFunction {
  @override
  String get name => 'VALUE';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    if (value is NumberValue) return value;
    final text = value.toText().trim();
    final n = num.tryParse(text);
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(n);
  }
}

/// TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...) - Join with delimiter.
class TextJoinFunction extends FormulaFunction {
  @override
  String get name => 'TEXTJOIN';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final delimiterValue = args[0].evaluate(context);
    if (delimiterValue.isError) return delimiterValue;
    final delimiter = delimiterValue.toText();

    final ignoreEmptyValue = args[1].evaluate(context);
    if (ignoreEmptyValue.isError) return ignoreEmptyValue;
    final ignoreEmpty = ignoreEmptyValue.isTruthy;

    final parts = <String>[];
    for (var i = 2; i < args.length; i++) {
      final value = args[i].evaluate(context);
      if (value.isError) return value;
      if (value is RangeValue) {
        for (final cell in value.flat) {
          final text = cell.toText();
          if (ignoreEmpty && text.isEmpty) continue;
          parts.add(text);
        }
      } else {
        final text = value.toText();
        if (ignoreEmpty && text.isEmpty) continue;
        parts.add(text);
      }
    }

    return FormulaValue.text(parts.join(delimiter));
  }
}

/// PROPER(text) - Capitalize first letter of each word.
class ProperFunction extends FormulaFunction {
  @override
  String get name => 'PROPER';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final text = value.toText();
    final buffer = StringBuffer();
    var capitalizeNext = true;
    for (final ch in text.runes) {
      final char = String.fromCharCode(ch);
      if (char.trim().isEmpty || !RegExp(r'[a-zA-Z]').hasMatch(char)) {
        buffer.write(char);
        capitalizeNext = true;
      } else {
        buffer.write(capitalizeNext ? char.toUpperCase() : char.toLowerCase());
        capitalizeNext = false;
      }
    }
    return FormulaValue.text(buffer.toString());
  }
}

/// EXACT(text1, text2) - Case-sensitive comparison.
class ExactFunction extends FormulaFunction {
  @override
  String get name => 'EXACT';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    return FormulaValue.boolean(values[0].toText() == values[1].toText());
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

/// REPT(text, number_times) - Repeats text a given number of times.
class ReptFunction extends FormulaFunction {
  @override
  String get name => 'REPT';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final text = values[0].toText();
    final count = values[1].toNumber()?.toInt();
    if (count == null) return const FormulaValue.error(FormulaError.value);
    if (count < 0) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.text(text * count);
  }
}

/// CHAR(number) - Returns the character specified by the code number.
class CharFunction extends FormulaFunction {
  @override
  String get name => 'CHAR';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber()?.toInt();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    if (n < 1 || n > 255) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.text(String.fromCharCode(n));
  }
}

/// CODE(text) - Returns a numeric code for the first character.
class CodeFunction extends FormulaFunction {
  @override
  String get name => 'CODE';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final text = value.toText();
    if (text.isEmpty) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(text.codeUnitAt(0));
  }
}

/// CLEAN(text) - Removes non-printable characters (codes 0-31).
class CleanFunction extends FormulaFunction {
  @override
  String get name => 'CLEAN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final text = value.toText();
    return FormulaValue.text(text.replaceAll(RegExp(r'[\x00-\x1F]'), ''));
  }
}

/// DOLLAR(number, [decimals]) - Formats a number as currency.
class DollarFunction extends FormulaFunction {
  @override
  String get name => 'DOLLAR';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final number = values[0].toNumber()?.toDouble();
    if (number == null) return const FormulaValue.error(FormulaError.value);
    final decimals = args.length > 1 ? values[1].toNumber()?.toInt() ?? 2 : 2;

    String formatted;
    if (decimals >= 0) {
      final multiplier = math.pow(10, decimals);
      final rounded = (number.abs() * multiplier).round() / multiplier;
      formatted = rounded.toStringAsFixed(decimals);
    } else {
      final factor = math.pow(10, -decimals).toInt();
      final rounded = (number.abs() / factor).round() * factor;
      formatted = rounded.toString();
    }

    // Add thousands separators to integer part
    final parts = formatted.split('.');
    var intPart = parts[0];
    if (intPart.length > 3) {
      final buffer = StringBuffer();
      for (var i = 0; i < intPart.length; i++) {
        if (i > 0 && (intPart.length - i) % 3 == 0) buffer.write(',');
        buffer.write(intPart[i]);
      }
      intPart = buffer.toString();
    }

    final formatted2 =
        parts.length > 1 ? '\$$intPart.${parts[1]}' : '\$$intPart';
    if (number < 0) return FormulaValue.text('($formatted2)');
    return FormulaValue.text(formatted2);
  }
}

/// FIXED(number, decimals, [no_commas]) - Formats a number with fixed decimals.
class FixedFunction extends FormulaFunction {
  @override
  String get name => 'FIXED';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final number = values[0].toNumber()?.toDouble();
    final decimals = values[1].toNumber()?.toInt();
    if (number == null || decimals == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    final noCommas = args.length > 2 ? values[2].isTruthy : false;

    String formatted;
    if (decimals >= 0) {
      final multiplier = math.pow(10, decimals);
      final rounded = (number * multiplier).round() / multiplier;
      formatted = rounded.toStringAsFixed(decimals);
    } else {
      final factor = math.pow(10, -decimals).toInt();
      final rounded = (number / factor).round() * factor;
      formatted = rounded.toString();
    }

    if (!noCommas) {
      final isNegative = formatted.startsWith('-');
      var s = isNegative ? formatted.substring(1) : formatted;
      final parts = s.split('.');
      var intPart = parts[0];
      if (intPart.length > 3) {
        final buffer = StringBuffer();
        for (var i = 0; i < intPart.length; i++) {
          if (i > 0 && (intPart.length - i) % 3 == 0) buffer.write(',');
          buffer.write(intPart[i]);
        }
        intPart = buffer.toString();
      }
      s = parts.length > 1 ? '$intPart.${parts[1]}' : intPart;
      formatted = isNegative ? '-$s' : s;
    }

    return FormulaValue.text(formatted);
  }
}

/// T(value) - Returns the text if value is text, otherwise empty string.
class TFunction extends FormulaFunction {
  @override
  String get name => 'T';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    if (value is TextValue) return value;
    if (value.isError) return value;
    return const FormulaValue.text('');
  }
}

/// NUMBERVALUE(text, [decimal_separator], [group_separator]) - Converts text to number.
class NumberValueFunction extends FormulaFunction {
  @override
  String get name => 'NUMBERVALUE';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    var text = values[0].toText().trim();
    final decimalSep = args.length > 1 ? values[1].toText() : '.';
    final groupSep = args.length > 2 ? values[2].toText() : ',';

    // Remove group separators
    if (groupSep.isNotEmpty) text = text.replaceAll(groupSep, '');
    // Replace decimal separator with '.'
    if (decimalSep != '.') text = text.replaceAll(decimalSep, '.');

    // Handle percentage
    if (text.endsWith('%')) {
      text = text.substring(0, text.length - 1);
      final n = num.tryParse(text);
      if (n == null) return const FormulaValue.error(FormulaError.value);
      return FormulaValue.number(n / 100);
    }

    final n = num.tryParse(text);
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(n);
  }
}

/// UNICHAR(number) - Returns the Unicode character for a code point.
class UnicharFunction extends FormulaFunction {
  @override
  String get name => 'UNICHAR';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber()?.toInt();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    if (n < 1) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.text(String.fromCharCode(n));
  }
}

/// UNICODE(text) - Returns the Unicode code point for the first character.
class UnicodeFunction extends FormulaFunction {
  @override
  String get name => 'UNICODE';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final text = value.toText();
    if (text.isEmpty) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(text.runes.first);
  }
}

/// TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
class TextBeforeFunction extends FormulaFunction {
  @override
  String get name => 'TEXTBEFORE';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 6;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final text = values[0].toText();
    final delimiter = values[1].toText();
    final instanceNum =
        args.length > 2 ? values[2].toNumber()?.toInt() ?? 1 : 1;
    final matchMode =
        args.length > 3 ? values[3].toNumber()?.toInt() ?? 0 : 0;
    final ifNotFound = args.length > 5 ? values[5] : null;

    if (instanceNum == 0) {
      return const FormulaValue.error(FormulaError.value);
    }

    if (delimiter.isEmpty) {
      return ifNotFound ?? const FormulaValue.error(FormulaError.na);
    }

    final caseSensitive = matchMode == 0;
    final searchText = caseSensitive ? text : text.toLowerCase();
    final searchDelim = caseSensitive ? delimiter : delimiter.toLowerCase();

    // Find all occurrences
    final positions = <int>[];
    var start = 0;
    while (true) {
      final pos = searchText.indexOf(searchDelim, start);
      if (pos == -1) break;
      positions.add(pos);
      start = pos + searchDelim.length;
    }

    if (positions.isEmpty) {
      return ifNotFound ?? const FormulaValue.error(FormulaError.na);
    }

    int targetIndex;
    if (instanceNum > 0) {
      targetIndex = instanceNum - 1;
    } else {
      targetIndex = positions.length + instanceNum;
    }

    if (targetIndex < 0 || targetIndex >= positions.length) {
      return ifNotFound ?? const FormulaValue.error(FormulaError.na);
    }

    return FormulaValue.text(text.substring(0, positions[targetIndex]));
  }
}

/// TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
class TextAfterFunction extends FormulaFunction {
  @override
  String get name => 'TEXTAFTER';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 6;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final text = values[0].toText();
    final delimiter = values[1].toText();
    final instanceNum =
        args.length > 2 ? values[2].toNumber()?.toInt() ?? 1 : 1;
    final matchMode =
        args.length > 3 ? values[3].toNumber()?.toInt() ?? 0 : 0;
    final ifNotFound = args.length > 5 ? values[5] : null;

    if (instanceNum == 0) {
      return const FormulaValue.error(FormulaError.value);
    }

    if (delimiter.isEmpty) {
      return ifNotFound ?? const FormulaValue.error(FormulaError.na);
    }

    final caseSensitive = matchMode == 0;
    final searchText = caseSensitive ? text : text.toLowerCase();
    final searchDelim = caseSensitive ? delimiter : delimiter.toLowerCase();

    // Find all occurrences
    final positions = <int>[];
    var start = 0;
    while (true) {
      final pos = searchText.indexOf(searchDelim, start);
      if (pos == -1) break;
      positions.add(pos);
      start = pos + searchDelim.length;
    }

    if (positions.isEmpty) {
      return ifNotFound ?? const FormulaValue.error(FormulaError.na);
    }

    int targetIndex;
    if (instanceNum > 0) {
      targetIndex = instanceNum - 1;
    } else {
      targetIndex = positions.length + instanceNum;
    }

    if (targetIndex < 0 || targetIndex >= positions.length) {
      return ifNotFound ?? const FormulaValue.error(FormulaError.na);
    }

    return FormulaValue.text(
        text.substring(positions[targetIndex] + delimiter.length));
  }
}

/// TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])
class TextSplitFunction extends FormulaFunction {
  @override
  String get name => 'TEXTSPLIT';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 6;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final text = values[0].toText();
    final colDelimiter = values[1].toText();
    final rowDelimiter =
        args.length > 2 && values[2] is! EmptyValue ? values[2].toText() : null;
    final ignoreEmpty = args.length > 3 ? values[3].isTruthy : false;
    final matchMode =
        args.length > 4 ? values[4].toNumber()?.toInt() ?? 0 : 0;
    final padWith = args.length > 5
        ? values[5]
        : const FormulaValue.error(FormulaError.na);

    final caseSensitive = matchMode == 0;

    List<String> rows;
    if (rowDelimiter != null && rowDelimiter.isNotEmpty) {
      rows = _splitText(text, rowDelimiter, caseSensitive, ignoreEmpty);
    } else {
      rows = [text];
    }

    final result = <List<FormulaValue>>[];
    var maxCols = 0;
    for (final row in rows) {
      final cols = _splitText(row, colDelimiter, caseSensitive, ignoreEmpty);
      if (cols.length > maxCols) maxCols = cols.length;
      result.add(cols.map<FormulaValue>(FormulaValue.text).toList());
    }

    // Pad rows to same number of columns
    for (final row in result) {
      while (row.length < maxCols) {
        row.add(padWith);
      }
    }

    if (result.length == 1 && result[0].length == 1) {
      return result[0][0];
    }

    return RangeValue(result);
  }

  List<String> _splitText(
      String text, String delimiter, bool caseSensitive, bool ignoreEmpty) {
    if (delimiter.isEmpty) return [text];
    final parts = <String>[];
    final searchText = caseSensitive ? text : text.toLowerCase();
    final searchDelim = caseSensitive ? delimiter : delimiter.toLowerCase();
    var start = 0;
    while (true) {
      final pos = searchText.indexOf(searchDelim, start);
      if (pos == -1) {
        final part = text.substring(start);
        if (!ignoreEmpty || part.isNotEmpty) parts.add(part);
        break;
      }
      final part = text.substring(start, pos);
      if (!ignoreEmpty || part.isNotEmpty) parts.add(part);
      start = pos + delimiter.length;
    }
    return parts;
  }
}

/// ARRAYTOTEXT(array, [format]) - Converts an array to text.
/// format=0 (default): concise — comma-separated, semicolons between rows
/// format=1: strict — braces, quoted strings: {1,"hello";3,4}
class ArrayToTextFunction extends FormulaFunction {
  @override
  String get name => 'ARRAYTOTEXT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final arrayVal = values[0];
    final format = args.length > 1 ? values[1].toNumber()?.toInt() ?? 0 : 0;

    if (arrayVal is RangeValue) {
      final rows = <String>[];
      for (final row in arrayVal.values) {
        final cells = <String>[];
        for (final cell in row) {
          cells.add(_formatCell(cell, format));
        }
        rows.add(cells.join(format == 1 ? ',' : ', '));
      }
      final joined = rows.join(format == 1 ? ';' : '; ');
      return FormulaValue.text(format == 1 ? '{$joined}' : joined);
    }

    return FormulaValue.text(_formatCell(arrayVal, format));
  }

  String _formatCell(FormulaValue cell, int format) {
    if (format == 1 && cell is TextValue) {
      return '"${cell.value}"';
    }
    return cell.toText();
  }
}

/// VALUETOTEXT(value, [format]) - Converts a value to text.
/// format=0 (default): concise — plain value
/// format=1: strict — text gets quotes
class ValueToTextFunction extends FormulaFunction {
  @override
  String get name => 'VALUETOTEXT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final value = values[0];
    final format = args.length > 1 ? values[1].toNumber()?.toInt() ?? 0 : 0;

    if (format == 1 && value is TextValue) {
      return FormulaValue.text('"${value.value}"');
    }
    return FormulaValue.text(value.toText());
  }
}

/// ASC(text) - Converts full-width (CJK) characters to half-width.
class AscFunction extends FormulaFunction {
  @override
  String get name => 'ASC';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final text = value.toText();
    final buffer = StringBuffer();
    for (final codeUnit in text.runes) {
      if (codeUnit >= 0xFF01 && codeUnit <= 0xFF5E) {
        // Fullwidth ASCII variants → ASCII
        buffer.writeCharCode(codeUnit - 0xFF01 + 0x0021);
      } else if (codeUnit == 0x3000) {
        // Ideographic space → regular space
        buffer.writeCharCode(0x0020);
      } else {
        buffer.writeCharCode(codeUnit);
      }
    }
    return FormulaValue.text(buffer.toString());
  }
}

/// DBCS(text) - Converts half-width characters to full-width (CJK).
class DbcsFunction extends FormulaFunction {
  @override
  String get name => 'DBCS';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final text = value.toText();
    final buffer = StringBuffer();
    for (final codeUnit in text.runes) {
      if (codeUnit >= 0x0021 && codeUnit <= 0x007E) {
        // ASCII → fullwidth ASCII variants
        buffer.writeCharCode(codeUnit - 0x0021 + 0xFF01);
      } else if (codeUnit == 0x0020) {
        // Regular space → ideographic space
        buffer.writeCharCode(0x3000);
      } else {
        buffer.writeCharCode(codeUnit);
      }
    }
    return FormulaValue.text(buffer.toString());
  }
}

/// BAHTTEXT(number) - Converts a number to Thai Baht text.
class BahtTextFunction extends FormulaFunction {
  @override
  String get name => 'BAHTTEXT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  static const _digits = [
    '',
    '\u0E2B\u0E19\u0E36\u0E48\u0E07', // หนึ่ง
    '\u0E2A\u0E2D\u0E07', // สอง
    '\u0E2A\u0E32\u0E21', // สาม
    '\u0E2A\u0E35\u0E48', // สี่
    '\u0E2B\u0E49\u0E32', // ห้า
    '\u0E2B\u0E01', // หก
    '\u0E40\u0E08\u0E47\u0E14', // เจ็ด
    '\u0E41\u0E1B\u0E14', // แปด
    '\u0E40\u0E01\u0E49\u0E32', // เก้า
  ];

  static const _positions = [
    '',
    '\u0E2A\u0E34\u0E1A', // สิบ
    '\u0E23\u0E49\u0E2D\u0E22', // ร้อย
    '\u0E1E\u0E31\u0E19', // พัน
    '\u0E2B\u0E21\u0E37\u0E48\u0E19', // หมื่น
    '\u0E41\u0E2A\u0E19', // แสน
  ];

  static const _million = '\u0E25\u0E49\u0E32\u0E19'; // ล้าน
  static const _baht = '\u0E1A\u0E32\u0E17'; // บาท
  static const _satang = '\u0E2A\u0E15\u0E32\u0E07\u0E04\u0E4C'; // สตางค์
  static const _exact = '\u0E16\u0E49\u0E27\u0E19'; // ถ้วน
  static const _negative = '\u0E25\u0E1A'; // ลบ
  static const _zero = '\u0E28\u0E39\u0E19\u0E22\u0E4C'; // ศูนย์
  static const _ed = '\u0E40\u0E2D\u0E47\u0E14'; // เอ็ด
  static const _yee = '\u0E22\u0E35\u0E48'; // ยี่

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber()?.toDouble();
    if (n == null) return const FormulaValue.error(FormulaError.value);

    if (n == 0) return FormulaValue.text('$_zero$_baht$_exact');

    final isNeg = n < 0;
    final absN = n.abs();

    // Split into integer and satang (2 decimal places)
    final rounded = (absN * 100).round();
    final intPart = rounded ~/ 100;
    final satangPart = rounded % 100;

    final buffer = StringBuffer();
    if (isNeg) buffer.write(_negative);

    if (intPart > 0) {
      buffer.write(_convertGroup(intPart));
      buffer.write(_baht);
    }

    if (satangPart > 0) {
      if (intPart == 0) {
        // No baht prefix needed but we still need to say zero baht
      }
      buffer.write(_convertGroup(satangPart));
      buffer.write(_satang);
    } else {
      buffer.write(_exact);
    }

    return FormulaValue.text(buffer.toString());
  }

  /// Convert an integer to Thai text (handles millions recursively).
  String _convertGroup(int n) {
    if (n == 0) return '';
    if (n >= 1000000) {
      final millions = n ~/ 1000000;
      final remainder = n % 1000000;
      return '${_convertGroup(millions)}$_million${_convertGroup(remainder)}';
    }

    final buffer = StringBuffer();
    final digits = <int>[];
    var temp = n;
    while (temp > 0) {
      digits.insert(0, temp % 10);
      temp ~/= 10;
    }

    for (var i = 0; i < digits.length; i++) {
      final pos = digits.length - 1 - i;
      final d = digits[i];
      if (d == 0) continue;
      if (pos == 1 && d == 1) {
        // 1 in tens position is just สิบ (not หนึ่งสิบ)
        buffer.write(_positions[pos]);
      } else if (pos == 1 && d == 2) {
        // 2 in tens position is ยี่สิบ
        buffer.write(_yee);
        buffer.write(_positions[pos]);
      } else if (pos == 0 && d == 1 && digits.length > 1) {
        // 1 in ones position (when not alone) is เอ็ด
        buffer.write(_ed);
      } else {
        buffer.write(_digits[d]);
        if (pos > 0) buffer.write(_positions[pos]);
      }
    }
    return buffer.toString();
  }
}
