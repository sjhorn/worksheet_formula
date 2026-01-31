import 'package:a1/a1.dart';

import '../evaluation/context.dart';
import '../evaluation/errors.dart';
import '../evaluation/value.dart';
import 'operators.dart';

/// Base class for all formula AST nodes.
sealed class FormulaNode {
  const FormulaNode();

  /// Evaluate this node given an evaluation context.
  FormulaValue evaluate(EvaluationContext context);

  /// Get all cell/range references in this node (for dependency tracking).
  Iterable<A1> get cellReferences;

  /// Get the formula string representation.
  String toFormulaString();
}

/// Number literal: 42, 3.14, -17
class NumberNode extends FormulaNode {
  final num value;
  const NumberNode(this.value);

  @override
  FormulaValue evaluate(EvaluationContext context) =>
      FormulaValue.number(value);

  @override
  Iterable<A1> get cellReferences => const [];

  @override
  String toFormulaString() => value.toString();

  @override
  bool operator ==(Object other) =>
      other is NumberNode && other.value == value;

  @override
  int get hashCode => value.hashCode;
}

/// String literal: "hello", "world"
class TextNode extends FormulaNode {
  final String value;
  const TextNode(this.value);

  @override
  FormulaValue evaluate(EvaluationContext context) =>
      FormulaValue.text(value);

  @override
  Iterable<A1> get cellReferences => const [];

  @override
  String toFormulaString() => '"${value.replaceAll('"', '""')}"';

  @override
  bool operator ==(Object other) =>
      other is TextNode && other.value == value;

  @override
  int get hashCode => value.hashCode;
}

/// Boolean literal: TRUE, FALSE
class BooleanNode extends FormulaNode {
  final bool value;
  const BooleanNode(this.value);

  @override
  FormulaValue evaluate(EvaluationContext context) =>
      FormulaValue.boolean(value);

  @override
  Iterable<A1> get cellReferences => const [];

  @override
  String toFormulaString() => value ? 'TRUE' : 'FALSE';

  @override
  bool operator ==(Object other) =>
      other is BooleanNode && other.value == value;

  @override
  int get hashCode => value.hashCode;
}

/// Error literal: #REF!, #VALUE!, etc.
class ErrorNode extends FormulaNode {
  final FormulaError error;
  const ErrorNode(this.error);

  @override
  FormulaValue evaluate(EvaluationContext context) =>
      FormulaValue.error(error);

  @override
  Iterable<A1> get cellReferences => const [];

  @override
  String toFormulaString() => error.code;
}

/// Cell reference: A1, $B$2, Sheet1!C3
class CellRefNode extends FormulaNode {
  final A1 cell;
  final String? sheet;

  const CellRefNode(this.cell, {this.sheet});

  @override
  FormulaValue evaluate(EvaluationContext context) =>
      context.getCellValue(cell);

  @override
  Iterable<A1> get cellReferences => [cell];

  @override
  String toFormulaString() =>
      sheet != null ? '$sheet!$cell' : cell.toString();

  @override
  bool operator ==(Object other) =>
      other is CellRefNode && other.cell == cell && other.sheet == sheet;

  @override
  int get hashCode => Object.hash(cell, sheet);
}

/// Range reference: A1:B10, Sheet1!A1:C3
class RangeRefNode extends FormulaNode {
  final A1Range range;
  final String? sheet;

  const RangeRefNode(this.range, {this.sheet});

  @override
  FormulaValue evaluate(EvaluationContext context) =>
      context.getRangeValues(range);

  @override
  Iterable<A1> get cellReferences {
    final from = range.from.a1;
    final to = range.to.a1;
    if (from == null || to == null) return const [];
    final cells = <A1>[];
    for (var row = from.row; row <= to.row; row++) {
      for (var col = from.column; col <= to.column; col++) {
        cells.add(A1.fromVector(col, row));
      }
    }
    return cells;
  }

  @override
  String toFormulaString() =>
      sheet != null ? '$sheet!$range' : range.toString();

  @override
  bool operator ==(Object other) =>
      other is RangeRefNode && other.range == range && other.sheet == sheet;

  @override
  int get hashCode => Object.hash(range, sheet);
}

/// Binary operation: A1 + B1, 2 * 3, "a" & "b"
class BinaryOpNode extends FormulaNode {
  final FormulaNode left;
  final BinaryOperator operator;
  final FormulaNode right;

  const BinaryOpNode(this.left, this.operator, this.right);

  @override
  FormulaValue evaluate(EvaluationContext context) {
    final leftVal = left.evaluate(context);

    // Short-circuit for errors
    if (leftVal.isError && operator != BinaryOperator.equal) {
      return leftVal;
    }

    final rightVal = right.evaluate(context);
    if (rightVal.isError && operator != BinaryOperator.equal) {
      return rightVal;
    }

    return operator.apply(leftVal, rightVal);
  }

  @override
  Iterable<A1> get cellReferences => [
        ...left.cellReferences,
        ...right.cellReferences,
      ];

  @override
  String toFormulaString() =>
      '${left.toFormulaString()}${operator.symbol}${right.toFormulaString()}';
}

/// Unary operation: -A1, +5, 50%
class UnaryOpNode extends FormulaNode {
  final UnaryOperator operator;
  final FormulaNode operand;

  const UnaryOpNode(this.operator, this.operand);

  @override
  FormulaValue evaluate(EvaluationContext context) {
    final val = operand.evaluate(context);
    if (val.isError) return val;
    return operator.apply(val);
  }

  @override
  Iterable<A1> get cellReferences => operand.cellReferences;

  @override
  String toFormulaString() =>
      '${operator.symbol}${operand.toFormulaString()}';
}

/// Function call: SUM(A1:A10), IF(A1>0, "yes", "no")
class FunctionCallNode extends FormulaNode {
  final String name;
  final List<FormulaNode> arguments;

  const FunctionCallNode(this.name, this.arguments);

  @override
  FormulaValue evaluate(EvaluationContext context) {
    final func = context.getFunction(name);
    if (func == null) {
      return const FormulaValue.error(FormulaError.name);
    }

    if (arguments.length < func.minArgs) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (func.maxArgs >= 0 && arguments.length > func.maxArgs) {
      return const FormulaValue.error(FormulaError.value);
    }

    return func.call(arguments, context);
  }

  @override
  Iterable<A1> get cellReferences =>
      arguments.expand((arg) => arg.cellReferences);

  @override
  String toFormulaString() =>
      '$name(${arguments.map((a) => a.toFormulaString()).join(',')})';
}

/// Parenthesized expression (preserves formatting in toFormulaString).
class ParenthesizedNode extends FormulaNode {
  final FormulaNode inner;
  const ParenthesizedNode(this.inner);

  @override
  FormulaValue evaluate(EvaluationContext context) =>
      inner.evaluate(context);

  @override
  Iterable<A1> get cellReferences => inner.cellReferences;

  @override
  String toFormulaString() => '(${inner.toFormulaString()})';
}
