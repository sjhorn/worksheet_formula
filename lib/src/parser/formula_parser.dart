import 'package:a1/a1.dart';
import 'package:petitparser/petitparser.dart';

import '../ast/nodes.dart';
import '../ast/operators.dart';
import '../evaluation/errors.dart';

/// Parser for spreadsheet formulas.
class FormulaParser {
  late final Parser<FormulaNode> _parser;

  FormulaParser() {
    _parser = _buildParser();
  }

  /// Parse a formula string into an AST.
  ///
  /// Throws [FormulaParseException] if the formula is invalid.
  FormulaNode parse(String formula) {
    final result = _parser.parse(formula);

    if (result is Success<FormulaNode>) {
      return result.value;
    } else {
      throw FormulaParseException(
        message: _improveErrorMessage(result.message, result.position, formula),
        position: result.position,
        formula: formula,
      );
    }
  }

  /// Post-process petitparser error messages for better diagnostics.
  String _improveErrorMessage(
    String originalMessage,
    int position,
    String formula,
  ) {
    // "end of input expected" means extra characters after a complete expression
    if (originalMessage.contains('end of input expected')) {
      if (position < formula.length && formula[position] == ')') {
        return 'Unexpected closing parenthesis at position $position';
      }
      return 'Unexpected character at position $position';
    }

    // Check for unmatched parentheses
    final openCount = '('.allMatches(formula).length;
    final closeCount = ')'.allMatches(formula).length;
    if (openCount > closeCount) {
      return 'Unexpected end of formula: missing closing ")"';
    }

    // At or past end of input
    if (position >= formula.length) {
      return 'Unexpected end of formula';
    }

    return originalMessage;
  }

  /// Try to parse a formula, returning null on failure.
  FormulaNode? tryParse(String formula) {
    final result = _parser.parse(formula);
    if (result is Success<FormulaNode>) {
      return result.value;
    }
    return null;
  }

  /// Check if a string is a valid formula.
  bool isValid(String formula) => tryParse(formula) != null;

  Parser<FormulaNode> _buildParser() {
    // Forward reference for recursive expressions
    final expression = undefined<FormulaNode>();

    // --- Primitives ---

    // Number: 42, 3.14, .5
    final number = (digit().plus() &
                (char('.') & digit().star()).optional() |
            char('.') & digit().plus())
        .flatten()
        .trim()
        .map((str) => NumberNode(num.parse(str)) as FormulaNode);

    // String: "hello", "say ""hi"""
    final text = (char('"') &
            (string('""').map((_) => '"') | pattern('^"'))
                .star()
                .map((chars) => chars.join()) &
            char('"'))
        .map((values) => TextNode(values[1] as String) as FormulaNode);

    // Boolean: TRUE, FALSE (case-insensitive)
    final boolean = (string('TRUE', ignoreCase: true).map((_) => true) |
            string('FALSE', ignoreCase: true).map((_) => false))
        .trim()
        .map((value) => BooleanNode(value as bool) as FormulaNode);

    // Error literals
    final errorLiteral = (string('#DIV/0!')
                .map((_) => FormulaError.divZero) |
            string('#VALUE!').map((_) => FormulaError.value) |
            string('#REF!').map((_) => FormulaError.ref) |
            string('#NAME?').map((_) => FormulaError.name) |
            string('#NUM!').map((_) => FormulaError.num) |
            string('#N/A').map((_) => FormulaError.na) |
            string('#NULL!').map((_) => FormulaError.null_))
        .trim()
        .map((error) => ErrorNode(error as FormulaError) as FormulaNode);

    // Cell/range reference via A1Reference
    // Matches: A1, $A$1, $A1, A$1, A1:B2, $A$1:$B$2, Sheet1!A1, etc.
    final a1Cell = (char(r'$').optional() &
            pattern('A-Za-z').plus().flatten() &
            char(r'$').optional() &
            digit().plus().flatten())
        .flatten();

    final sheetPrefix = ((char("'") & pattern("^'").plus().flatten() &
                    char("'")) |
                (letter() & word().star()).flatten())
            .flatten() &
        char('!');

    final rangeReference = (sheetPrefix.optional().flatten() &
            a1Cell &
            char(':') &
            a1Cell)
        .map((values) {
      final sheetPart = (values[0] as String).isNotEmpty ? values[0] : null;
      final fromCell = values[1] as String;
      final toCell = values[3] as String;
      final refStr =
          sheetPart != null ? '$sheetPart$fromCell:$toCell' : '$fromCell:$toCell';
      final ref = A1Reference.parse(refStr);
      return RangeRefNode(ref) as FormulaNode;
    });

    final cellReference =
        (sheetPrefix.optional().flatten() & a1Cell).map((values) {
      final sheetPart = (values[0] as String).isNotEmpty ? values[0] : null;
      final cellStr = values[1] as String;
      final refStr = sheetPart != null ? '$sheetPart$cellStr' : cellStr;
      final ref = A1Reference.parse(refStr);
      return CellRefNode(ref) as FormulaNode;
    });

    // Function call: SUM(...), IF(...)
    final functionName =
        (letter() & (letter() | digit() | char('_') | char('.')).star())
            .flatten()
            .trim();

    final argumentList = expression
        .plusSeparated(char(',').trim() | char(';').trim())
        .map((result) => result.elements.toList());

    final functionCall =
        (functionName & char('(').trim() & argumentList.optional() &
                char(')').trim())
            .map((values) {
      final name = (values[0] as String).toUpperCase();
      final args = values[2] as List<FormulaNode>? ?? [];
      return FunctionCallNode(name, args) as FormulaNode;
    });

    // Parenthesized expression
    final parenthesized =
        (char('(').trim() & expression & char(')').trim())
            .map((values) =>
                ParenthesizedNode(values[1] as FormulaNode) as FormulaNode);

    // Bare identifier: x, count, my_var (for LAMBDA parameters).
    // Must NOT match TRUE/FALSE, cell refs (letter+digit like A1), or
    // function calls (already matched earlier via functionCall).
    final bareIdentifier =
        (letter() & (letter() | char('_')).star()).flatten().trim().where((s) {
      final upper = s.toUpperCase();
      if (upper == 'TRUE' || upper == 'FALSE') return false;
      return true;
    }).map((name) => NameNode(name) as FormulaNode);

    // --- Primary (order matters: try function before cell ref,
    //     range before cell ref, bare identifier last) ---
    final primary = (parenthesized |
            functionCall |
            rangeReference |
            cellReference |
            number |
            text |
            boolean |
            errorLiteral |
            bareIdentifier)
        .cast<FormulaNode>();

    // --- Expression with operator precedence ---
    final builder = ExpressionBuilder<FormulaNode>();
    builder.primitive(primary);

    // Postfix: (args) for immediate invocation, then %
    builder.group()
      ..postfix(
          (char('(').trim() & argumentList.optional() & char(')').trim()),
          (value, op) {
            final args = op[1] as List<FormulaNode>? ?? [];
            return CallExpressionNode(value, args);
          })
      ..postfix(
          char('%').trim(),
          (value, _) =>
              UnaryOpNode(UnaryOperator.percent, value));

    // Prefix: -, +
    builder.group()
      ..prefix(
          char('-').trim(),
          (_, value) =>
              UnaryOpNode(UnaryOperator.negate, value))
      ..prefix(char('+').trim(), (_, value) => value);

    // Power: ^ (right-associative)
    builder.group().right(
        char('^').trim(),
        (l, _, r) =>
            BinaryOpNode(l, BinaryOperator.power, r));

    // Multiply, Divide
    builder.group()
      ..left(
          char('*').trim(),
          (l, _, r) =>
              BinaryOpNode(l, BinaryOperator.multiply, r))
      ..left(
          char('/').trim(),
          (l, _, r) =>
              BinaryOpNode(l, BinaryOperator.divide, r));

    // Add, Subtract
    builder.group()
      ..left(
          char('+').trim(),
          (l, _, r) =>
              BinaryOpNode(l, BinaryOperator.add, r))
      ..left(
          char('-').trim(),
          (l, _, r) =>
              BinaryOpNode(l, BinaryOperator.subtract, r));

    // Concatenation: &
    builder.group().left(
        char('&').trim(),
        (l, _, r) =>
            BinaryOpNode(l, BinaryOperator.concat, r));

    // Comparison: =, <>, <, >, <=, >=
    builder.group()
      ..left(
          string('<=').trim(),
          (l, _, r) =>
              BinaryOpNode(l, BinaryOperator.lessEqual, r))
      ..left(
          string('>=').trim(),
          (l, _, r) =>
              BinaryOpNode(l, BinaryOperator.greaterEqual, r))
      ..left(
          string('<>').trim(),
          (l, _, r) =>
              BinaryOpNode(l, BinaryOperator.notEqual, r))
      ..left(
          char('<').trim(),
          (l, _, r) =>
              BinaryOpNode(l, BinaryOperator.lessThan, r))
      ..left(
          char('>').trim(),
          (l, _, r) =>
              BinaryOpNode(l, BinaryOperator.greaterThan, r))
      ..left(
          char('=').trim(),
          (l, _, r) =>
              BinaryOpNode(l, BinaryOperator.equal, r));

    final expr = builder.build();
    expression.set(expr);

    // Formula: optional '=' prefix then expression
    return (char('=').optional() & expr).map((values) {
      return values[1] as FormulaNode;
    }).end();
  }
}

/// Exception thrown when formula parsing fails.
class FormulaParseException implements Exception {
  final String message;
  final int position;
  final String formula;

  FormulaParseException({
    required this.message,
    required this.position,
    required this.formula,
  });

  @override
  String toString() {
    final pointer = '${' ' * position}^';
    return 'FormulaParseException: $message\n  $formula\n  $pointer';
  }
}
