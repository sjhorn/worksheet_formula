import 'dart:math' as math;

import '../evaluation/errors.dart';
import '../evaluation/value.dart';

/// Binary operators supported in formulas.
enum BinaryOperator {
  // Arithmetic
  add('+', 4),
  subtract('-', 4),
  multiply('*', 5),
  divide('/', 5),
  power('^', 6),

  // Comparison
  equal('=', 2),
  notEqual('<>', 2),
  lessThan('<', 2),
  greaterThan('>', 2),
  lessEqual('<=', 2),
  greaterEqual('>=', 2),

  // Text
  concat('&', 3);

  final String symbol;
  final int precedence;

  const BinaryOperator(this.symbol, this.precedence);

  /// Apply this operator to two values.
  FormulaValue apply(FormulaValue left, FormulaValue right) {
    if (left is ErrorValue) return left;
    if (right is ErrorValue) return right;

    return switch (this) {
      BinaryOperator.add =>
        _applyArithmetic(left, right, (a, b) => a + b),
      BinaryOperator.subtract =>
        _applyArithmetic(left, right, (a, b) => a - b),
      BinaryOperator.multiply =>
        _applyArithmetic(left, right, (a, b) => a * b),
      BinaryOperator.divide => _applyDivide(left, right),
      BinaryOperator.power =>
        _applyArithmetic(left, right, (a, b) => math.pow(a, b)),
      BinaryOperator.equal =>
        _applyComparison(left, right, (cmp) => cmp == 0),
      BinaryOperator.notEqual =>
        _applyComparison(left, right, (cmp) => cmp != 0),
      BinaryOperator.lessThan =>
        _applyComparison(left, right, (cmp) => cmp < 0),
      BinaryOperator.greaterThan =>
        _applyComparison(left, right, (cmp) => cmp > 0),
      BinaryOperator.lessEqual =>
        _applyComparison(left, right, (cmp) => cmp <= 0),
      BinaryOperator.greaterEqual =>
        _applyComparison(left, right, (cmp) => cmp >= 0),
      BinaryOperator.concat =>
        FormulaValue.text(left.toText() + right.toText()),
    };
  }

  FormulaValue _applyArithmetic(
    FormulaValue left,
    FormulaValue right,
    num Function(num, num) op,
  ) {
    final l = left.toNumber();
    final r = right.toNumber();
    if (l == null || r == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    return FormulaValue.number(op(l, r));
  }

  FormulaValue _applyDivide(FormulaValue left, FormulaValue right) {
    final l = left.toNumber();
    final r = right.toNumber();
    if (l == null || r == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (r == 0) {
      return const FormulaValue.error(FormulaError.divZero);
    }
    return FormulaValue.number(l / r);
  }

  FormulaValue _applyComparison(
    FormulaValue left,
    FormulaValue right,
    bool Function(int) predicate,
  ) {
    final cmp = _compare(left, right);
    return FormulaValue.boolean(predicate(cmp));
  }

  int _compare(FormulaValue left, FormulaValue right) {
    // Same types compare directly
    if (left is NumberValue && right is NumberValue) {
      return left.value.compareTo(right.value);
    }
    if (left is TextValue && right is TextValue) {
      return left.value.toLowerCase().compareTo(right.value.toLowerCase());
    }
    if (left is BooleanValue && right is BooleanValue) {
      return left.value == right.value ? 0 : (left.value ? 1 : -1);
    }

    // Mixed types: try numeric comparison
    final l = left.toNumber();
    final r = right.toNumber();
    if (l != null && r != null) {
      return l.compareTo(r);
    }

    // Fall back to string comparison
    return left.toText().toLowerCase().compareTo(right.toText().toLowerCase());
  }
}

/// Unary operators supported in formulas.
enum UnaryOperator {
  negate('-'),
  positive('+'),
  percent('%');

  final String symbol;
  const UnaryOperator(this.symbol);

  /// Apply this operator to a value.
  FormulaValue apply(FormulaValue operand) {
    if (operand is ErrorValue) return operand;

    return switch (this) {
      UnaryOperator.negate => _negate(operand),
      UnaryOperator.positive => operand,
      UnaryOperator.percent => _percent(operand),
    };
  }

  FormulaValue _negate(FormulaValue operand) {
    final n = operand.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(-n);
  }

  FormulaValue _percent(FormulaValue operand) {
    final n = operand.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(n / 100);
  }
}
