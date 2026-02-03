import '../ast/nodes.dart';
import '../evaluation/context.dart';
import '../evaluation/errors.dart';
import '../evaluation/value.dart';
import 'function.dart';
import 'registry.dart';

/// Register all web and regex functions.
void registerWebFunctions(FunctionRegistry registry) {
  registry.registerAll([
    EncodeUrlFunction(),
    RegexMatchFunction(),
    RegexExtractFunction(),
    RegexReplaceFunction(),
  ]);
}

/// ENCODEURL(text) - URI percent-encodes a string.
class EncodeUrlFunction extends FormulaFunction {
  @override
  String get name => 'ENCODEURL';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    if (value.isError) return value;
    return FormulaValue.text(Uri.encodeComponent(value.toText()));
  }
}

/// REGEXMATCH(text, regular_expression) - Returns TRUE if text matches pattern.
class RegexMatchFunction extends FormulaFunction {
  @override
  String get name => 'REGEXMATCH';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final text = values[0].toText();
    final pattern = values[1].toText();

    try {
      final regex = RegExp(pattern);
      return FormulaValue.boolean(regex.hasMatch(text));
    } on FormatException {
      return const FormulaValue.error(FormulaError.value);
    }
  }
}

/// REGEXEXTRACT(text, regular_expression) - Returns first matching substring.
class RegexExtractFunction extends FormulaFunction {
  @override
  String get name => 'REGEXEXTRACT';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final text = values[0].toText();
    final pattern = values[1].toText();

    try {
      final regex = RegExp(pattern);
      final match = regex.firstMatch(text);
      if (match == null) {
        return const FormulaValue.error(FormulaError.na);
      }
      // If there's a capture group, return it; otherwise return full match
      if (match.groupCount > 0) {
        return FormulaValue.text(match.group(1) ?? '');
      }
      return FormulaValue.text(match.group(0) ?? '');
    } on FormatException {
      return const FormulaValue.error(FormulaError.value);
    }
  }
}

/// REGEXREPLACE(text, regular_expression, replacement) - Replaces all matches.
class RegexReplaceFunction extends FormulaFunction {
  @override
  String get name => 'REGEXREPLACE';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final text = values[0].toText();
    final pattern = values[1].toText();
    final replacement = values[2].toText();

    try {
      final regex = RegExp(pattern);
      return FormulaValue.text(text.replaceAll(regex, replacement));
    } on FormatException {
      return const FormulaValue.error(FormulaError.value);
    }
  }
}
