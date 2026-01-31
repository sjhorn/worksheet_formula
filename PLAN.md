# Worksheet Formula Engine Architecture

A comprehensive guide to implementing an extensible formula engine for spreadsheet applications, designed to work with the `worksheet` widget while remaining completely standalone.

## Overview

This document describes the architecture for `worksheet_formula`, a standalone Dart package that provides Excel/Google Sheets compatible formula parsing and evaluation. The package is designed to:

- **Be standalone** - No dependency on any UI framework or specific data structure
- **Leverage existing packages** - Built on `a1` for cell references and `petitparser` for parsing
- **Be extensible** - Easy to add custom functions
- **Match Excel/Sheets semantics** - Error types, operator precedence, type coercion

## Package Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                         Package Ecosystem                            │
├─────────────────────────────────────────────────────────────────────┤
│                                                                      │
│  ┌──────────────┐    ┌────────────────────┐    ┌──────────────────┐ │
│  │     a1       │◀───│  worksheet_formula │    │    worksheet     │ │
│  │              │    │                    │    │                  │ │
│  │ • A1         │    │ • FormulaParser    │    │ • Cell           │ │
│  │ • A1Range    │    │ • FormulaNode (AST)│    │ • WorksheetData  │ │
│  │ • A1Reference│    │ • FunctionRegistry │    │ • Worksheet      │ │
│  │ • petitparser│    │ • FormulaValue     │    │                  │ │
│  └──────────────┘    │ • EvaluationContext│    └────────┬─────────┘ │
│                      │   (abstract)       │             │           │
│                      └─────────┬──────────┘             │           │
│                                │                        │           │
│                                └────────────┬───────────┘           │
│                                             │                       │
│                                             ▼                       │
│                               ┌─────────────────────────┐           │
│                               │  example / integration  │           │
│                               │                         │           │
│                               │ • WorksheetFormulaEngine│           │
│                               │   (concrete impl)       │           │
│                               └─────────────────────────┘           │
│                                                                      │
└─────────────────────────────────────────────────────────────────────┘
```

### Key Design Decisions

1. **`worksheet_formula` has no dependency on `worksheet`** - It only depends on `a1` and `petitparser`
2. **Abstract `EvaluationContext`** - Consumers implement this to connect their data source
3. **Integration lives in examples/consumer packages** - The wiring between `worksheet` and `worksheet_formula` is done by the consumer

## Package Structure

```
worksheet_formula/
├── lib/
│   ├── worksheet_formula.dart          # Main export
│   ├── src/
│   │   ├── ast/
│   │   │   ├── nodes.dart              # FormulaNode and subclasses
│   │   │   └── operators.dart          # Binary/Unary operators
│   │   ├── parser/
│   │   │   ├── formula_parser.dart     # Main parser
│   │   │   └── grammar.dart            # PetitParser grammar
│   │   ├── evaluation/
│   │   │   ├── context.dart            # Abstract EvaluationContext
│   │   │   ├── value.dart              # FormulaValue types
│   │   │   └── errors.dart             # FormulaError (#REF!, #VALUE!, etc)
│   │   ├── functions/
│   │   │   ├── registry.dart           # FunctionRegistry
│   │   │   ├── function.dart           # FormulaFunction base
│   │   │   ├── math.dart               # SUM, AVERAGE, etc.
│   │   │   ├── text.dart               # CONCAT, LEFT, RIGHT, etc.
│   │   │   ├── logical.dart            # IF, AND, OR, etc.
│   │   │   ├── lookup.dart             # VLOOKUP, INDEX, MATCH, etc.
│   │   │   ├── date.dart               # DATE, NOW, TODAY, etc.
│   │   │   └── statistical.dart        # COUNT, COUNTIF, etc.
│   │   └── dependencies/
│   │       └── graph.dart              # Dependency tracking
├── pubspec.yaml
├── example/
│   └── worksheet_integration/          # Example integration with worksheet
└── test/
```

### pubspec.yaml

```yaml
name: worksheet_formula
description: A standalone formula engine for spreadsheet-like calculations
version: 1.0.0

environment:
  sdk: ^3.0.0

dependencies:
  a1: ^2.0.0
  petitparser: ^7.0.0

dev_dependencies:
  test: ^1.24.0
```

---

## Core Components

### 1. Formula Errors

Excel-compatible error types that can result from formula evaluation.

```dart
// lib/src/evaluation/errors.dart

/// Excel-compatible formula errors
enum FormulaError {
  /// #DIV/0! - Division by zero
  divZero('#DIV/0!'),
  
  /// #VALUE! - Wrong type of argument
  value('#VALUE!'),
  
  /// #REF! - Invalid cell reference
  ref('#REF!'),
  
  /// #NAME? - Unrecognized formula name
  name('#NAME?'),
  
  /// #NUM! - Invalid numeric value
  num('#NUM!'),
  
  /// #N/A - Value not available
  na('#N/A'),
  
  /// #NULL! - Incorrect range operator
  null_('#NULL!'),
  
  /// #CALC! - Calculation error
  calc('#CALC!'),
  
  /// #CIRCULAR! - Circular reference detected
  circular('#CIRCULAR!');
  
  final String code;
  const FormulaError(this.code);
  
  @override
  String toString() => code;
}
```

### 2. Formula Values

Type-safe representation of all values that can result from formula evaluation.

```dart
// lib/src/evaluation/value.dart

import 'errors.dart';

/// Represents any value that can result from formula evaluation
sealed class FormulaValue {
  const FormulaValue();
  
  // Factory constructors
  const factory FormulaValue.number(num value) = NumberValue;
  const factory FormulaValue.text(String value) = TextValue;
  const factory FormulaValue.boolean(bool value) = BooleanValue;
  const factory FormulaValue.error(FormulaError error) = ErrorValue;
  const factory FormulaValue.range(List<List<FormulaValue>> values) = RangeValue;
  const factory FormulaValue.empty() = EmptyValue;
  
  /// Convert to number (for arithmetic operations)
  num? toNumber();
  
  /// Convert to string (for text operations)
  String toText();
  
  /// Convert to boolean (for logical operations)
  bool toBool();
  
  /// Is this value "truthy" for IF conditions?
  bool get isTruthy;
  
  /// Is this an error value?
  bool get isError => this is ErrorValue;
}

class NumberValue extends FormulaValue {
  final num value;
  const NumberValue(this.value);
  
  @override
  num? toNumber() => value;
  
  @override
  String toText() => value.toString();
  
  @override
  bool toBool() => value != 0;
  
  @override
  bool get isTruthy => value != 0;
  
  @override
  bool operator ==(Object other) =>
    other is NumberValue && other.value == value;
  
  @override
  int get hashCode => value.hashCode;
  
  @override
  String toString() => 'NumberValue($value)';
}

class TextValue extends FormulaValue {
  final String value;
  const TextValue(this.value);
  
  @override
  num? toNumber() => num.tryParse(value);
  
  @override
  String toText() => value;
  
  @override
  bool toBool() => value.isNotEmpty;
  
  @override
  bool get isTruthy => value.isNotEmpty;
  
  @override
  bool operator ==(Object other) =>
    other is TextValue && other.value == value;
  
  @override
  int get hashCode => value.hashCode;
  
  @override
  String toString() => 'TextValue("$value")';
}

class BooleanValue extends FormulaValue {
  final bool value;
  const BooleanValue(this.value);
  
  @override
  num? toNumber() => value ? 1 : 0;
  
  @override
  String toText() => value ? 'TRUE' : 'FALSE';
  
  @override
  bool toBool() => value;
  
  @override
  bool get isTruthy => value;
  
  @override
  bool operator ==(Object other) =>
    other is BooleanValue && other.value == value;
  
  @override
  int get hashCode => value.hashCode;
  
  @override
  String toString() => 'BooleanValue($value)';
}

class ErrorValue extends FormulaValue {
  final FormulaError error;
  const ErrorValue(this.error);
  
  @override
  num? toNumber() => null;
  
  @override
  String toText() => error.code;
  
  @override
  bool toBool() => false;
  
  @override
  bool get isTruthy => false;
  
  @override
  bool operator ==(Object other) =>
    other is ErrorValue && other.error == error;
  
  @override
  int get hashCode => error.hashCode;
  
  @override
  String toString() => 'ErrorValue(${error.code})';
}

class RangeValue extends FormulaValue {
  final List<List<FormulaValue>> values;
  const RangeValue(this.values);
  
  int get rowCount => values.length;
  int get columnCount => values.isEmpty ? 0 : values.first.length;
  
  /// Flatten to a single list of values
  Iterable<FormulaValue> get flat => values.expand((row) => row);
  
  /// Get numeric values only
  Iterable<num> get numbers => flat
    .whereType<NumberValue>()
    .map((v) => v.value);
  
  @override
  num? toNumber() => values.length == 1 && values.first.length == 1
    ? values.first.first.toNumber()
    : null;
  
  @override
  String toText() => values
    .map((row) => row.map((v) => v.toText()).join(','))
    .join(';');
  
  @override
  bool toBool() => toNumber() != 0;
  
  @override
  bool get isTruthy => values.isNotEmpty;
  
  @override
  String toString() => 'RangeValue(${rowCount}x$columnCount)';
}

class EmptyValue extends FormulaValue {
  const EmptyValue();
  
  @override
  num? toNumber() => 0; // Empty cells are 0 in numeric context
  
  @override
  String toText() => '';
  
  @override
  bool toBool() => false;
  
  @override
  bool get isTruthy => false;
  
  @override
  String toString() => 'EmptyValue()';
}
```

### 3. Evaluation Context (Abstract)

The key abstraction that allows the formula engine to work with any data source.

```dart
// lib/src/evaluation/context.dart

import 'package:a1/a1.dart';
import 'value.dart';
import '../functions/function.dart';

/// Abstract context for formula evaluation.
/// 
/// Implement this interface to connect the formula engine to your data source.
/// This is the primary integration point for consumers of the package.
abstract class EvaluationContext {
  /// Get the value of a single cell
  FormulaValue getCellValue(A1 cell);
  
  /// Get values for a range of cells as a 2D matrix
  FormulaValue getRangeValues(A1Range range);
  
  /// Get a function by name (case-insensitive)
  FormulaFunction? getFunction(String name);
  
  /// The cell currently being evaluated (for relative references)
  A1 get currentCell;
  
  /// The current sheet name (for cross-sheet references)
  String? get currentSheet;
  
  /// Optional: Check if evaluation should be cancelled (for long-running calcs)
  bool get isCancelled => false;
}
```

### 4. Operators

Binary and unary operators with their evaluation logic.

```dart
// lib/src/ast/operators.dart

import '../evaluation/value.dart';
import '../evaluation/errors.dart';

/// Binary operators supported in formulas
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
  
  /// Apply this operator to two values
  FormulaValue apply(FormulaValue left, FormulaValue right) {
    // Handle errors
    if (left is ErrorValue) return left;
    if (right is ErrorValue) return right;
    
    return switch (this) {
      BinaryOperator.add => _applyArithmetic(left, right, (a, b) => a + b),
      BinaryOperator.subtract => _applyArithmetic(left, right, (a, b) => a - b),
      BinaryOperator.multiply => _applyArithmetic(left, right, (a, b) => a * b),
      BinaryOperator.divide => _applyDivide(left, right),
      BinaryOperator.power => _applyArithmetic(left, right, (a, b) => _pow(a, b)),
      BinaryOperator.equal => _applyComparison(left, right, (cmp) => cmp == 0),
      BinaryOperator.notEqual => _applyComparison(left, right, (cmp) => cmp != 0),
      BinaryOperator.lessThan => _applyComparison(left, right, (cmp) => cmp < 0),
      BinaryOperator.greaterThan => _applyComparison(left, right, (cmp) => cmp > 0),
      BinaryOperator.lessEqual => _applyComparison(left, right, (cmp) => cmp <= 0),
      BinaryOperator.greaterEqual => _applyComparison(left, right, (cmp) => cmp >= 0),
      BinaryOperator.concat => FormulaValue.text(left.toText() + right.toText()),
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
    if (cmp == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    return FormulaValue.boolean(predicate(cmp));
  }
  
  int? _compare(FormulaValue left, FormulaValue right) {
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
  
  num _pow(num base, num exponent) {
    if (base == 0 && exponent < 0) return double.infinity;
    return _powerImpl(base.toDouble(), exponent.toDouble());
  }
  
  double _powerImpl(double base, double exponent) {
    // Using dart:math would be cleaner, but keeping it simple
    if (exponent == 0) return 1;
    if (exponent == 1) return base;
    if (exponent < 0) return 1 / _powerImpl(base, -exponent);
    if (exponent == exponent.truncate()) {
      var result = 1.0;
      for (var i = 0; i < exponent; i++) {
        result *= base;
      }
      return result;
    }
    // For non-integer exponents, we'd need dart:math.pow
    // This is a simplified implementation
    return base; // Placeholder - use dart:math in real implementation
  }
}

/// Unary operators supported in formulas
enum UnaryOperator {
  negate('-'),
  positive('+'),
  percent('%');
  
  final String symbol;
  const UnaryOperator(this.symbol);
  
  /// Apply this operator to a value
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
```

### 5. AST Nodes

The Abstract Syntax Tree representation of parsed formulas.

```dart
// lib/src/ast/nodes.dart

import 'package:a1/a1.dart';
import '../evaluation/context.dart';
import '../evaluation/value.dart';
import '../evaluation/errors.dart';
import 'operators.dart';

/// Base class for all formula AST nodes
sealed class FormulaNode {
  const FormulaNode();
  
  /// Evaluate this node given an evaluation context
  FormulaValue evaluate(EvaluationContext context);
  
  /// Get all cell/range references in this node (for dependency tracking)
  Iterable<A1Reference> get references;
  
  /// Get the formula string representation
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
  Iterable<A1Reference> get references => const [];
  
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
  Iterable<A1Reference> get references => const [];
  
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
  Iterable<A1Reference> get references => const [];
  
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
  Iterable<A1Reference> get references => const [];
  
  @override
  String toFormulaString() => error.code;
}

/// Cell reference: A1, $B$2, Sheet1!C3
class CellRefNode extends FormulaNode {
  final A1Reference reference;
  const CellRefNode(this.reference);
  
  @override
  FormulaValue evaluate(EvaluationContext context) {
    final cell = reference.range?.from ?? reference.a1;
    if (cell == null) return const FormulaValue.error(FormulaError.ref);
    return context.getCellValue(cell);
  }
  
  @override
  Iterable<A1Reference> get references => [reference];
  
  @override
  String toFormulaString() => reference.toString();
  
  @override
  bool operator ==(Object other) =>
    other is CellRefNode && other.reference == reference;
  
  @override
  int get hashCode => reference.hashCode;
}

/// Range reference: A1:B10, Sheet1!A1:C3
class RangeRefNode extends FormulaNode {
  final A1Reference reference;
  const RangeRefNode(this.reference);
  
  @override
  FormulaValue evaluate(EvaluationContext context) {
    final range = reference.range;
    if (range == null) return const FormulaValue.error(FormulaError.ref);
    return context.getRangeValues(range);
  }
  
  @override
  Iterable<A1Reference> get references => [reference];
  
  @override
  String toFormulaString() => reference.toString();
  
  @override
  bool operator ==(Object other) =>
    other is RangeRefNode && other.reference == reference;
  
  @override
  int get hashCode => reference.hashCode;
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
  Iterable<A1Reference> get references => [
    ...left.references,
    ...right.references,
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
  Iterable<A1Reference> get references => operand.references;
  
  @override
  String toFormulaString() => '${operator.symbol}${operand.toFormulaString()}';
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
    
    // Validate argument count
    if (arguments.length < func.minArgs) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (func.maxArgs >= 0 && arguments.length > func.maxArgs) {
      return const FormulaValue.error(FormulaError.value);
    }
    
    return func.call(arguments, context);
  }
  
  @override
  Iterable<A1Reference> get references =>
    arguments.expand((arg) => arg.references);
  
  @override
  String toFormulaString() =>
    '$name(${arguments.map((a) => a.toFormulaString()).join(',')})';
}

/// Parenthesized expression (preserves formatting in toFormulaString)
class ParenthesizedNode extends FormulaNode {
  final FormulaNode inner;
  const ParenthesizedNode(this.inner);
  
  @override
  FormulaValue evaluate(EvaluationContext context) => inner.evaluate(context);
  
  @override
  Iterable<A1Reference> get references => inner.references;
  
  @override
  String toFormulaString() => '(${inner.toFormulaString()})';
}
```

### 6. Formula Functions

The base class and registry for spreadsheet functions.

```dart
// lib/src/functions/function.dart

import '../ast/nodes.dart';
import '../evaluation/context.dart';
import '../evaluation/value.dart';
import '../evaluation/errors.dart';

/// Base class for all formula functions
abstract class FormulaFunction {
  /// Function name (for error messages and registration)
  String get name;
  
  /// Minimum number of required arguments
  int get minArgs;
  
  /// Maximum number of arguments (-1 for unlimited)
  int get maxArgs;
  
  /// Whether this function uses lazy evaluation (like IF)
  /// 
  /// Lazy functions receive unevaluated FormulaNode arguments and
  /// can choose which ones to evaluate. Eager functions receive
  /// pre-evaluated FormulaValue arguments.
  bool get isLazy => false;
  
  /// Evaluate the function with the given arguments
  FormulaValue call(List<FormulaNode> args, EvaluationContext context);
  
  /// Helper to evaluate all arguments eagerly
  List<FormulaValue> evaluateArgs(
    List<FormulaNode> args, 
    EvaluationContext context,
  ) {
    return args.map((arg) => arg.evaluate(context)).toList();
  }
  
  /// Helper to require a numeric value
  FormulaValue requireNumber(FormulaValue value) {
    if (value.isError) return value;
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(n);
  }
  
  /// Helper to collect all numbers from arguments (including ranges)
  Iterable<num> collectNumbers(
    List<FormulaNode> args, 
    EvaluationContext context,
  ) sync* {
    for (final arg in args) {
      final value = arg.evaluate(context);
      switch (value) {
        case NumberValue(value: final n):
          yield n;
        case RangeValue(values: final matrix):
          for (final row in matrix) {
            for (final cell in row) {
              if (cell is NumberValue) {
                yield cell.value;
              }
            }
          }
        default:
          // Skip non-numeric values (Excel behavior)
          break;
      }
    }
  }
}
```

```dart
// lib/src/functions/registry.dart

import 'function.dart';
import 'math.dart';
import 'text.dart';
import 'logical.dart';
import 'lookup.dart';
import 'date.dart';
import 'statistical.dart';

/// Registry of available formula functions
class FunctionRegistry {
  final Map<String, FormulaFunction> _functions = {};
  
  /// Create a new registry, optionally with built-in functions pre-registered
  FunctionRegistry({bool registerBuiltIns = true}) {
    if (registerBuiltIns) {
      registerMathFunctions(this);
      registerTextFunctions(this);
      registerLogicalFunctions(this);
      registerLookupFunctions(this);
      registerDateFunctions(this);
      registerStatisticalFunctions(this);
    }
  }
  
  /// Register a function
  void register(FormulaFunction function) {
    _functions[function.name.toUpperCase()] = function;
  }
  
  /// Register multiple functions
  void registerAll(Iterable<FormulaFunction> functions) {
    for (final func in functions) {
      register(func);
    }
  }
  
  /// Get a function by name (case-insensitive)
  FormulaFunction? get(String name) => _functions[name.toUpperCase()];
  
  /// Check if a function exists
  bool has(String name) => _functions.containsKey(name.toUpperCase());
  
  /// Get all registered function names
  Iterable<String> get names => _functions.keys;
  
  /// Create a copy with additional functions
  FunctionRegistry copyWith(Iterable<FormulaFunction> additional) {
    final copy = FunctionRegistry(registerBuiltIns: false);
    copy._functions.addAll(_functions);
    copy.registerAll(additional);
    return copy;
  }
}
```

### 7. Built-in Functions

Example implementations of common spreadsheet functions.

```dart
// lib/src/functions/math.dart

import '../ast/nodes.dart';
import '../evaluation/context.dart';
import '../evaluation/value.dart';
import '../evaluation/errors.dart';
import 'function.dart';
import 'registry.dart';

/// Register all math functions
void registerMathFunctions(FunctionRegistry registry) {
  registry.registerAll([
    SumFunction(),
    AverageFunction(),
    MinFunction(),
    MaxFunction(),
    AbsFunction(),
    RoundFunction(),
    IntFunction(),
    ModFunction(),
    SqrtFunction(),
    PowerFunction(),
  ]);
}

/// SUM(number1, [number2], ...) - Adds all numbers
class SumFunction extends FormulaFunction {
  @override
  String get name => 'SUM';
  
  @override
  int get minArgs => 1;
  
  @override
  int get maxArgs => -1;
  
  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    var sum = 0.0;
    for (final n in collectNumbers(args, context)) {
      sum += n;
    }
    return FormulaValue.number(sum);
  }
}

/// AVERAGE(number1, [number2], ...) - Returns the average
class AverageFunction extends FormulaFunction {
  @override
  String get name => 'AVERAGE';
  
  @override
  int get minArgs => 1;
  
  @override
  int get maxArgs => -1;
  
  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = collectNumbers(args, context).toList();
    if (numbers.isEmpty) {
      return const FormulaValue.error(FormulaError.divZero);
    }
    final sum = numbers.fold(0.0, (a, b) => a + b);
    return FormulaValue.number(sum / numbers.length);
  }
}

/// MIN(number1, [number2], ...) - Returns the minimum value
class MinFunction extends FormulaFunction {
  @override
  String get name => 'MIN';
  
  @override
  int get minArgs => 1;
  
  @override
  int get maxArgs => -1;
  
  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = collectNumbers(args, context).toList();
    if (numbers.isEmpty) {
      return const FormulaValue.number(0);
    }
    return FormulaValue.number(numbers.reduce((a, b) => a < b ? a : b));
  }
}

/// MAX(number1, [number2], ...) - Returns the maximum value
class MaxFunction extends FormulaFunction {
  @override
  String get name => 'MAX';
  
  @override
  int get minArgs => 1;
  
  @override
  int get maxArgs => -1;
  
  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = collectNumbers(args, context).toList();
    if (numbers.isEmpty) {
      return const FormulaValue.number(0);
    }
    return FormulaValue.number(numbers.reduce((a, b) => a > b ? a : b));
  }
}

/// ABS(number) - Returns the absolute value
class AbsFunction extends FormulaFunction {
  @override
  String get name => 'ABS';
  
  @override
  int get minArgs => 1;
  
  @override
  int get maxArgs => 1;
  
  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(n.abs());
  }
}

/// ROUND(number, num_digits) - Rounds a number
class RoundFunction extends FormulaFunction {
  @override
  String get name => 'ROUND';
  
  @override
  int get minArgs => 2;
  
  @override
  int get maxArgs => 2;
  
  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final number = values[0].toNumber();
    final digits = values[1].toNumber()?.toInt();
    
    if (number == null || digits == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    
    final multiplier = _pow10(digits);
    final rounded = (number * multiplier).round() / multiplier;
    return FormulaValue.number(rounded);
  }
  
  double _pow10(int n) {
    if (n >= 0) {
      var result = 1.0;
      for (var i = 0; i < n; i++) result *= 10;
      return result;
    } else {
      var result = 1.0;
      for (var i = 0; i > n; i--) result /= 10;
      return result;
    }
  }
}

/// INT(number) - Rounds down to the nearest integer
class IntFunction extends FormulaFunction {
  @override
  String get name => 'INT';
  
  @override
  int get minArgs => 1;
  
  @override
  int get maxArgs => 1;
  
  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(n.floor());
  }
}

/// MOD(number, divisor) - Returns the remainder
class ModFunction extends FormulaFunction {
  @override
  String get name => 'MOD';
  
  @override
  int get minArgs => 2;
  
  @override
  int get maxArgs => 2;
  
  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final number = values[0].toNumber();
    final divisor = values[1].toNumber();
    
    if (number == null || divisor == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (divisor == 0) {
      return const FormulaValue.error(FormulaError.divZero);
    }
    
    return FormulaValue.number(number % divisor);
  }
}

/// SQRT(number) - Returns the square root
class SqrtFunction extends FormulaFunction {
  @override
  String get name => 'SQRT';
  
  @override
  int get minArgs => 1;
  
  @override
  int get maxArgs => 1;
  
  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    if (n < 0) return const FormulaValue.error(FormulaError.num);
    
    // Simple Newton-Raphson sqrt (use dart:math in real implementation)
    if (n == 0) return const FormulaValue.number(0);
    var x = n.toDouble();
    for (var i = 0; i < 20; i++) {
      x = (x + n / x) / 2;
    }
    return FormulaValue.number(x);
  }
}

/// POWER(number, power) - Returns number raised to a power
class PowerFunction extends FormulaFunction {
  @override
  String get name => 'POWER';
  
  @override
  int get minArgs => 2;
  
  @override
  int get maxArgs => 2;
  
  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final base = values[0].toNumber();
    final exponent = values[1].toNumber();
    
    if (base == null || exponent == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    
    // Use dart:math.pow in real implementation
    var result = 1.0;
    final exp = exponent.toInt();
    for (var i = 0; i < exp.abs(); i++) {
      result *= base;
    }
    if (exp < 0) result = 1 / result;
    
    return FormulaValue.number(result);
  }
}
```

```dart
// lib/src/functions/logical.dart

import '../ast/nodes.dart';
import '../evaluation/context.dart';
import '../evaluation/value.dart';
import '../evaluation/errors.dart';
import 'function.dart';
import 'registry.dart';

/// Register all logical functions
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
  ]);
}

/// IF(logical_test, value_if_true, [value_if_false])
/// 
/// This is a lazy function - it only evaluates the branch that's needed.
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
    
    // Errors in condition propagate
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

/// AND(logical1, [logical2], ...) - Returns TRUE if all arguments are TRUE
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

/// OR(logical1, [logical2], ...) - Returns TRUE if any argument is TRUE
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

/// NOT(logical) - Reverses the logic of its argument
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

/// IFERROR(value, value_if_error) - Returns value_if_error if value is an error
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
    if (value.isError) {
      return args[1].evaluate(context);
    }
    return value;
  }
}

/// IFNA(value, value_if_na) - Returns value_if_na if value is #N/A
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

/// TRUE() - Returns the logical value TRUE
class TrueFunction extends FormulaFunction {
  @override
  String get name => 'TRUE';
  
  @override
  int get minArgs => 0;
  
  @override
  int get maxArgs => 0;
  
  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    return const FormulaValue.boolean(true);
  }
}

/// FALSE() - Returns the logical value FALSE
class FalseFunction extends FormulaFunction {
  @override
  String get name => 'FALSE';
  
  @override
  int get minArgs => 0;
  
  @override
  int get maxArgs => 0;
  
  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    return const FormulaValue.boolean(false);
  }
}
```

```dart
// lib/src/functions/text.dart

import '../ast/nodes.dart';
import '../evaluation/context.dart';
import '../evaluation/value.dart';
import '../evaluation/errors.dart';
import 'function.dart';
import 'registry.dart';

/// Register all text functions
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

/// CONCAT(text1, [text2], ...) - Joins text strings
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

/// CONCATENATE(text1, [text2], ...) - Legacy version of CONCAT
class ConcatenateFunction extends ConcatFunction {
  @override
  String get name => 'CONCATENATE';
}

/// LEFT(text, [num_chars]) - Returns leftmost characters
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
    final numChars = args.length > 1 
      ? values[1].toNumber()?.toInt() ?? 1
      : 1;
    
    if (numChars < 0) return const FormulaValue.error(FormulaError.value);
    
    final end = numChars > text.length ? text.length : numChars;
    return FormulaValue.text(text.substring(0, end));
  }
}

/// RIGHT(text, [num_chars]) - Returns rightmost characters
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
    final numChars = args.length > 1 
      ? values[1].toNumber()?.toInt() ?? 1
      : 1;
    
    if (numChars < 0) return const FormulaValue.error(FormulaError.value);
    
    final start = text.length - numChars;
    return FormulaValue.text(text.substring(start < 0 ? 0 : start));
  }
}

/// MID(text, start_num, num_chars) - Returns characters from middle
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
      text.substring(start, end > text.length ? text.length : end)
    );
  }
}

/// LEN(text) - Returns length of text
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

/// LOWER(text) - Converts to lowercase
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

/// UPPER(text) - Converts to uppercase
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

/// TRIM(text) - Removes extra spaces
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

/// TEXT(value, format_text) - Formats a number as text
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
    final format = values[1].toText();
    
    if (number == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    
    // Simplified format handling - real implementation would parse format codes
    return FormulaValue.text(number.toString());
  }
}
```

### 8. Formula Parser with PetitParser

The grammar and parser using `ExpressionBuilder` for correct operator precedence.

```dart
// lib/src/parser/grammar.dart

import 'package:petitparser/petitparser.dart';
import 'package:a1/a1.dart';
import '../ast/nodes.dart';
import '../ast/operators.dart';

/// Grammar definition for spreadsheet formulas
class FormulaGrammarDefinition extends GrammarDefinition<FormulaNode> {
  @override
  Parser<FormulaNode> start() => ref0(formula).end();

  /// A formula optionally starts with '='
  Parser<FormulaNode> formula() => 
    (char('=').optional() & ref0(expression))
      .map((values) => values[1] as FormulaNode);

  /// Top-level expression
  Parser<FormulaNode> expression() => ref0(comparison);

  /// Comparison operators (lowest precedence)
  /// =, <>, <, >, <=, >=
  Parser<FormulaNode> comparison() {
    final builder = ExpressionBuilder<FormulaNode>();
    
    builder.primitive(ref0(concatenation));
    
    builder.group()
      ..left(string('<=').trim(), (l, op, r) => 
        BinaryOpNode(l, BinaryOperator.lessEqual, r))
      ..left(string('>=').trim(), (l, op, r) => 
        BinaryOpNode(l, BinaryOperator.greaterEqual, r))
      ..left(string('<>').trim(), (l, op, r) => 
        BinaryOpNode(l, BinaryOperator.notEqual, r))
      ..left(char('=').trim(), (l, op, r) => 
        BinaryOpNode(l, BinaryOperator.equal, r))
      ..left(char('<').trim(), (l, op, r) => 
        BinaryOpNode(l, BinaryOperator.lessThan, r))
      ..left(char('>').trim(), (l, op, r) => 
        BinaryOpNode(l, BinaryOperator.greaterThan, r));
    
    return builder.build();
  }

  /// Text concatenation operator: &
  Parser<FormulaNode> concatenation() {
    final builder = ExpressionBuilder<FormulaNode>();
    builder.primitive(ref0(additive));
    builder.group()
      .left(char('&').trim(), (l, op, r) => 
        BinaryOpNode(l, BinaryOperator.concat, r));
    return builder.build();
  }

  /// Addition and subtraction: +, -
  Parser<FormulaNode> additive() {
    final builder = ExpressionBuilder<FormulaNode>();
    builder.primitive(ref0(multiplicative));
    builder.group()
      ..left(char('+').trim(), (l, op, r) => 
        BinaryOpNode(l, BinaryOperator.add, r))
      ..left(char('-').trim(), (l, op, r) => 
        BinaryOpNode(l, BinaryOperator.subtract, r));
    return builder.build();
  }

  /// Multiplication and division: *, /
  Parser<FormulaNode> multiplicative() {
    final builder = ExpressionBuilder<FormulaNode>();
    builder.primitive(ref0(power));
    builder.group()
      ..left(char('*').trim(), (l, op, r) => 
        BinaryOpNode(l, BinaryOperator.multiply, r))
      ..left(char('/').trim(), (l, op, r) => 
        BinaryOpNode(l, BinaryOperator.divide, r));
    return builder.build();
  }

  /// Exponentiation: ^ (right-associative)
  Parser<FormulaNode> power() {
    final builder = ExpressionBuilder<FormulaNode>();
    builder.primitive(ref0(percentage));
    builder.group()
      .right(char('^').trim(), (l, op, r) => 
        BinaryOpNode(l, BinaryOperator.power, r));
    return builder.build();
  }

  /// Percentage: % (postfix)
  Parser<FormulaNode> percentage() {
    final builder = ExpressionBuilder<FormulaNode>();
    builder.primitive(ref0(unary));
    builder.group()
      .postfix(char('%').trim(), (value, op) => 
        UnaryOpNode(UnaryOperator.percent, value));
    return builder.build();
  }

  /// Unary operators: -, +
  Parser<FormulaNode> unary() {
    final builder = ExpressionBuilder<FormulaNode>();
    builder.primitive(ref0(primary));
    builder.group()
      ..prefix(char('-').trim(), (op, value) => 
        UnaryOpNode(UnaryOperator.negate, value))
      ..prefix(char('+').trim(), (op, value) => value);
    return builder.build();
  }

  /// Primary expressions: literals, references, functions, parentheses
  Parser<FormulaNode> primary() => [
    ref0(parenthesized),
    ref0(functionCall),
    ref0(rangeReference),
    ref0(cellReference),
    ref0(number),
    ref0(text),
    ref0(boolean),
    ref0(errorLiteral),
  ].toChoiceParser();

  /// Parenthesized expression: (expr)
  Parser<FormulaNode> parenthesized() =>
    (char('(').trim() & ref0(expression) & char(')').trim())
      .map((values) => ParenthesizedNode(values[1] as FormulaNode));

  /// Function call: SUM(A1:A10), IF(A1>0, "yes", "no")
  Parser<FormulaNode> functionCall() =>
    (ref0(functionName) & char('(') & ref0(argumentList).optional() & char(')'))
      .map((values) {
        final name = values[0] as String;
        final args = values[2] as List<FormulaNode>? ?? [];
        return FunctionCallNode(name, args);
      });

  /// Function name: letters followed by letters/digits/underscores
  Parser<String> functionName() =>
    (letter() & (letter() | digit() | char('_')).star())
      .flatten()
      .trim();

  /// Argument list: comma or semicolon separated expressions
  Parser<List<FormulaNode>> argumentList() =>
    ref0(expression).plusSeparated(
      char(',').trim() | char(';').trim(),
    ).map((result) => result.elements.toList());

  /// Range reference: A1:B10, Sheet1!A1:C3
  Parser<FormulaNode> rangeReference() =>
    (ref0(sheetPrefix).optional() & ref0(a1Cell) & char(':') & ref0(a1Cell))
      .map((values) {
        final sheet = values[0] as String?;
        final from = values[1] as A1;
        final to = values[3] as A1;
        return RangeRefNode(A1Reference(
          worksheet: sheet,
          range: A1Range.fromA1s(from, to),
        ));
      });

  /// Cell reference: A1, $B$2, Sheet1!C3
  Parser<FormulaNode> cellReference() =>
    (ref0(sheetPrefix).optional() & ref0(a1Cell))
      .map((values) {
        final sheet = values[0] as String?;
        final cell = values[1] as A1;
        return CellRefNode(A1Reference(
          worksheet: sheet,
          range: A1Range.fromA1s(cell, cell),
        ));
      });

  /// A1-style cell notation: A1, $B$2, AA100
  Parser<A1> a1Cell() =>
    (char('\$').optional() & 
     pattern('A-Za-z').plus().flatten() & 
     char('\$').optional() & 
     digit().plus().flatten())
      .map((values) {
        final col = values[1] as String;
        final row = values[3] as String;
        return A1.parse('$col$row');
      });

  /// Sheet prefix: Sheet1!, 'Sheet Name'!
  Parser<String?> sheetPrefix() =>
    (ref0(sheetName) & char('!'))
      .map((values) => values[0] as String);

  /// Sheet name: quoted or unquoted
  Parser<String> sheetName() => [
    // Quoted: 'Sheet Name'
    (char("'") & pattern("^'").star().flatten() & char("'"))
      .map((values) => values[1] as String),
    // Unquoted: Sheet1
    (letter() & word().star()).flatten(),
  ].toChoiceParser();

  /// Number literal: 42, 3.14, .5, 1e10, 1.5E-3
  Parser<FormulaNode> number() =>
    (digit().plus() & (char('.') & digit().star()).optional() |
     char('.') & digit().plus())
      .flatten()
      .seq((char('e') | char('E'))
        .seq(char('+') | char('-')).optional()
        .seq(digit().plus())
        .flatten()
        .optional())
      .flatten()
      .trim()
      .map((str) => NumberNode(num.parse(str)));

  /// String literal: "hello", "say ""hi"""
  Parser<FormulaNode> text() =>
    (char('"') & 
     (string('""').map((_) => '"') | pattern('^"')).star().flatten() & 
     char('"'))
      .map((values) => TextNode(values[1] as String));

  /// Boolean literal: TRUE, FALSE
  Parser<FormulaNode> boolean() => [
    stringIgnoreCase('TRUE').map((_) => const BooleanNode(true)),
    stringIgnoreCase('FALSE').map((_) => const BooleanNode(false)),
  ].toChoiceParser().trim();

  /// Error literal: #REF!, #VALUE!, etc.
  Parser<FormulaNode> errorLiteral() => [
    string('#DIV/0!').map((_) => const ErrorNode(FormulaError.divZero)),
    string('#VALUE!').map((_) => const ErrorNode(FormulaError.value)),
    string('#REF!').map((_) => const ErrorNode(FormulaError.ref)),
    string('#NAME?').map((_) => const ErrorNode(FormulaError.name)),
    string('#NUM!').map((_) => const ErrorNode(FormulaError.num)),
    string('#N/A').map((_) => const ErrorNode(FormulaError.na)),
    string('#NULL!').map((_) => const ErrorNode(FormulaError.null_)),
  ].toChoiceParser().trim();
}
```

```dart
// lib/src/parser/formula_parser.dart

import 'package:petitparser/petitparser.dart';
import '../ast/nodes.dart';
import 'grammar.dart';

/// Parser for spreadsheet formulas
class FormulaParser {
  late final Parser<FormulaNode> _parser;
  
  FormulaParser() {
    final definition = FormulaGrammarDefinition();
    _parser = definition.build();
  }
  
  /// Parse a formula string into an AST
  /// 
  /// Throws [FormulaParseException] if the formula is invalid.
  FormulaNode parse(String formula) {
    final result = _parser.parse(formula);
    
    if (result is Success<FormulaNode>) {
      return result.value;
    } else if (result is Failure) {
      throw FormulaParseException(
        message: result.message,
        position: result.position,
        formula: formula,
      );
    } else {
      throw FormulaParseException(
        message: 'Unknown parse error',
        position: 0,
        formula: formula,
      );
    }
  }
  
  /// Try to parse a formula, returning null on failure
  FormulaNode? tryParse(String formula) {
    try {
      return parse(formula);
    } catch (_) {
      return null;
    }
  }
  
  /// Check if a string is a valid formula
  bool isValid(String formula) => tryParse(formula) != null;
}

/// Exception thrown when formula parsing fails
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
  String toString() => 
    'FormulaParseException: $message at position $position in "$formula"';
}
```

### 9. Dependency Graph

For tracking cell dependencies and efficient recalculation.

```dart
// lib/src/dependencies/graph.dart

import 'package:a1/a1.dart';

/// Tracks formula dependencies for efficient recalculation
/// 
/// This class maintains a bidirectional graph of cell dependencies:
/// - Which cells does a formula depend on?
/// - Which formulas depend on a given cell?
class DependencyGraph {
  /// Map from cell -> cells that depend on it (dependents)
  final Map<A1, Set<A1>> _dependents = {};
  
  /// Map from cell -> cells it depends on (dependencies)
  final Map<A1, Set<A1>> _dependencies = {};
  
  /// Update the dependencies for a cell
  /// 
  /// Call this when a cell's formula changes.
  void updateDependencies(A1 cell, Set<A1> newDependencies) {
    // Remove old dependency links
    final oldDeps = _dependencies[cell];
    if (oldDeps != null) {
      for (final dep in oldDeps) {
        _dependents[dep]?.remove(cell);
      }
    }
    
    // Add new dependency links
    if (newDependencies.isEmpty) {
      _dependencies.remove(cell);
    } else {
      _dependencies[cell] = newDependencies;
      for (final dep in newDependencies) {
        (_dependents[dep] ??= {}).add(cell);
      }
    }
  }
  
  /// Remove a cell from the dependency graph
  /// 
  /// Call this when a cell is cleared or deleted.
  void removeCell(A1 cell) {
    updateDependencies(cell, {});
    _dependents.remove(cell);
  }
  
  /// Get all cells that depend on the given cell
  Set<A1> getDependents(A1 cell) => _dependents[cell] ?? {};
  
  /// Get all cells that the given cell depends on
  Set<A1> getDependencies(A1 cell) => _dependencies[cell] ?? {};
  
  /// Get all cells that need recalculation when a cell changes
  /// 
  /// Returns cells in topological order (dependencies before dependents).
  List<A1> getCellsToRecalculate(A1 changedCell) {
    final result = <A1>[];
    final visited = <A1>{};
    final inProgress = <A1>{};
    
    void visit(A1 cell) {
      if (visited.contains(cell)) return;
      if (inProgress.contains(cell)) {
        // Circular dependency detected - skip but don't throw
        return;
      }
      
      inProgress.add(cell);
      
      for (final dependent in _dependents[cell] ?? <A1>{}) {
        visit(dependent);
      }
      
      inProgress.remove(cell);
      visited.add(cell);
      result.add(cell);
    }
    
    for (final dependent in _dependents[changedCell] ?? <A1>{}) {
      visit(dependent);
    }
    
    return result.reversed.toList();
  }
  
  /// Check if there's a circular reference involving the given cell
  bool hasCircularReference(A1 cell) {
    final visited = <A1>{};
    
    bool visit(A1 current, A1 target) {
      if (current == target && visited.isNotEmpty) return true;
      if (visited.contains(current)) return false;
      visited.add(current);
      
      for (final dep in _dependencies[current] ?? <A1>{}) {
        if (visit(dep, target)) return true;
      }
      return false;
    }
    
    return visit(cell, cell);
  }
  
  /// Clear all dependency information
  void clear() {
    _dependents.clear();
    _dependencies.clear();
  }
}
```

### 10. Formula Engine

The main entry point that ties everything together.

```dart
// lib/src/formula_engine.dart

import 'package:a1/a1.dart';
import 'ast/nodes.dart';
import 'parser/formula_parser.dart';
import 'evaluation/context.dart';
import 'evaluation/value.dart';
import 'functions/registry.dart';
import 'functions/function.dart';
import 'dependencies/graph.dart';

/// Main entry point for the formula engine
/// 
/// Use this class to parse and evaluate spreadsheet formulas.
/// 
/// ```dart
/// final engine = FormulaEngine();
/// 
/// // Register custom functions
/// engine.registerFunction(MyCustomFunction());
/// 
/// // Parse a formula
/// final ast = engine.parse('=SUM(A1:A10)');
/// 
/// // Evaluate with your data source
/// final result = engine.evaluate(ast, myContext);
/// ```
class FormulaEngine {
  final FormulaParser _parser;
  final FunctionRegistry _functions;
  final Map<String, FormulaNode> _parseCache = {};
  
  /// Create a new formula engine
  /// 
  /// Optionally provide a custom [FunctionRegistry] to use different
  /// or additional functions.
  FormulaEngine({FunctionRegistry? functions})
    : _parser = FormulaParser(),
      _functions = functions ?? FunctionRegistry();
  
  /// Parse a formula string into an AST
  /// 
  /// Results are cached for performance. Throws [FormulaParseException]
  /// if the formula is invalid.
  FormulaNode parse(String formula) {
    return _parseCache.putIfAbsent(
      formula,
      () => _parser.parse(formula),
    );
  }
  
  /// Try to parse a formula, returning null on failure
  FormulaNode? tryParse(String formula) {
    try {
      return parse(formula);
    } catch (_) {
      return null;
    }
  }
  
  /// Check if a string is a valid formula
  bool isValidFormula(String formula) => tryParse(formula) != null;
  
  /// Evaluate a parsed formula with the given context
  FormulaValue evaluate(FormulaNode ast, EvaluationContext context) {
    return ast.evaluate(context);
  }
  
  /// Parse and evaluate a formula string
  FormulaValue evaluateString(String formula, EvaluationContext context) {
    final ast = parse(formula);
    return evaluate(ast, context);
  }
  
  /// Get all cell references in a formula (for dependency tracking)
  Set<A1Reference> getReferences(String formula) {
    final ast = parse(formula);
    return ast.references.toSet();
  }
  
  /// Register a custom function
  void registerFunction(FormulaFunction function) {
    _functions.register(function);
  }
  
  /// Get the function registry
  /// 
  /// Use this when creating evaluation contexts.
  FunctionRegistry get functions => _functions;
  
  /// Clear the parse cache
  void clearCache() => _parseCache.clear();
}
```

### 11. Library Export

```dart
// lib/worksheet_formula.dart

/// A standalone formula engine for spreadsheet-like calculations.
/// 
/// This package provides:
/// - Excel/Google Sheets compatible formula parsing
/// - An extensible function registry with 50+ built-in functions
/// - Dependency tracking for efficient recalculation
/// - Type-safe formula values and error handling
/// 
/// ## Basic Usage
/// 
/// ```dart
/// import 'package:worksheet_formula/worksheet_formula.dart';
/// 
/// // Create the engine
/// final engine = FormulaEngine();
/// 
/// // Parse a formula
/// final ast = engine.parse('=SUM(A1:A10) * 2');
/// 
/// // Implement EvaluationContext to connect your data source
/// final context = MyEvaluationContext(myData);
/// 
/// // Evaluate
/// final result = engine.evaluate(ast, context);
/// ```
/// 
/// ## Custom Functions
/// 
/// ```dart
/// class DiscountFunction extends FormulaFunction {
///   @override String get name => 'DISCOUNT';
///   @override int get minArgs => 2;
///   @override int get maxArgs => 2;
///   
///   @override
///   FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
///     final values = evaluateArgs(args, context);
///     final price = values[0].toNumber() ?? 0;
///     final rate = values[1].toNumber() ?? 0;
///     return FormulaValue.number(price * (1 - rate));
///   }
/// }
/// 
/// engine.registerFunction(DiscountFunction());
/// ```
library worksheet_formula;

// Core engine
export 'src/formula_engine.dart';

// AST
export 'src/ast/nodes.dart';
export 'src/ast/operators.dart';

// Evaluation
export 'src/evaluation/context.dart';
export 'src/evaluation/value.dart';
export 'src/evaluation/errors.dart';

// Functions
export 'src/functions/function.dart';
export 'src/functions/registry.dart';

// Dependencies
export 'src/dependencies/graph.dart';

// Parser (for advanced users)
export 'src/parser/formula_parser.dart';
```

---

## Integration with Worksheet

This section shows how to integrate `worksheet_formula` with the `worksheet` package. This code would live in the `example/` directory or in a separate integration package.

```dart
// example/lib/worksheet_formula_integration.dart

import 'package:a1/a1.dart';
import 'package:worksheet/worksheet.dart';
import 'package:worksheet_formula/worksheet_formula.dart';

/// Concrete implementation of EvaluationContext for WorksheetData
class WorksheetEvaluationContext implements EvaluationContext {
  final WorksheetData data;
  final FunctionRegistry functions;
  final FormulaEngine engine;
  
  @override
  final A1 currentCell;
  
  @override
  final String? currentSheet;
  
  /// Track cells being evaluated to detect circular references
  final Set<A1> _evaluating;
  
  WorksheetEvaluationContext({
    required this.data,
    required this.functions,
    required this.engine,
    required this.currentCell,
    this.currentSheet,
    Set<A1>? evaluating,
  }) : _evaluating = evaluating ?? {};
  
  @override
  FormulaValue getCellValue(A1 cell) {
    // Check for circular reference
    if (_evaluating.contains(cell)) {
      return const FormulaValue.error(FormulaError.circular);
    }
    
    final coord = CellCoordinate(cell.row, cell.column);
    final cellData = data.getCell(coord);
    
    if (cellData == null) {
      return const FormulaValue.empty();
    }
    
    // If it's a formula, evaluate it recursively
    if (cellData.isFormula) {
      final newEvaluating = {..._evaluating, cell};
      final nestedContext = WorksheetEvaluationContext(
        data: data,
        functions: functions,
        engine: engine,
        currentCell: cell,
        currentSheet: currentSheet,
        evaluating: newEvaluating,
      );
      
      try {
        final ast = engine.parse(cellData.formula!);
        return ast.evaluate(nestedContext);
      } catch (e) {
        return const FormulaValue.error(FormulaError.value);
      }
    }
    
    // Return the cell's literal value
    return _cellToFormulaValue(cellData);
  }
  
  @override
  FormulaValue getRangeValues(A1Range range) {
    final values = <List<FormulaValue>>[];
    
    for (var row = range.from.row; row <= range.to.row; row++) {
      final rowValues = <FormulaValue>[];
      for (var col = range.from.column; col <= range.to.column; col++) {
        final cell = A1.fromCoordinates(row, col);
        rowValues.add(getCellValue(cell));
      }
      values.add(rowValues);
    }
    
    return FormulaValue.range(values);
  }
  
  @override
  FormulaFunction? getFunction(String name) => functions.get(name);
  
  FormulaValue _cellToFormulaValue(Cell cell) {
    final value = cell.value;
    return switch (value) {
      num n => FormulaValue.number(n),
      String s => FormulaValue.text(s),
      bool b => FormulaValue.boolean(b),
      DateTime d => FormulaValue.number(
        d.millisecondsSinceEpoch / 86400000 + 25569 // Excel date serial
      ),
      null => const FormulaValue.empty(),
      _ => FormulaValue.text(value.toString()),
    };
  }
}

/// Extension to add formula evaluation to WorksheetData
extension WorksheetFormulaExtension on WorksheetData {
  /// Evaluate a formula in the context of this worksheet
  FormulaValue evaluateFormula(
    String formula,
    CellCoordinate cell, {
    FormulaEngine? engine,
  }) {
    final eng = engine ?? FormulaEngine();
    final context = WorksheetEvaluationContext(
      data: this,
      functions: eng.functions,
      engine: eng,
      currentCell: A1.fromCoordinates(cell.row, cell.column),
    );
    return eng.evaluateString(formula, context);
  }
}

/// A wrapper that adds automatic formula evaluation to WorksheetData
class FormulaAwareWorksheetData implements WorksheetData {
  final WorksheetData _inner;
  final FormulaEngine _engine;
  final DependencyGraph _dependencies = DependencyGraph();
  final Map<CellCoordinate, FormulaValue> _cachedResults = {};
  
  FormulaAwareWorksheetData(this._inner, {FormulaEngine? engine})
    : _engine = engine ?? FormulaEngine();
  
  /// Get the formula engine
  FormulaEngine get engine => _engine;
  
  /// Get the dependency graph
  DependencyGraph get dependencies => _dependencies;
  
  @override
  Cell? getCell(CellCoordinate coord) {
    final cell = _inner.getCell(coord);
    if (cell == null || !cell.isFormula) return cell;
    
    // Return cell with computed value
    final result = _getCachedOrCompute(coord, cell.formula!);
    return cell.copyWithComputedValue(_formulaValueToObject(result));
  }
  
  FormulaValue _getCachedOrCompute(CellCoordinate coord, String formula) {
    // Check cache first
    final cached = _cachedResults[coord];
    if (cached != null) return cached;
    
    // Update dependencies
    final refs = _engine.getReferences(formula);
    final depCells = refs
      .expand((r) => _referenceToCells(r))
      .map((c) => A1.fromCoordinates(c.row, c.column))
      .toSet();
    _dependencies.updateDependencies(
      A1.fromCoordinates(coord.row, coord.column),
      depCells,
    );
    
    // Evaluate
    final context = WorksheetEvaluationContext(
      data: _inner,
      functions: _engine.functions,
      engine: _engine,
      currentCell: A1.fromCoordinates(coord.row, coord.column),
    );
    
    final result = _engine.evaluateString(formula, context);
    _cachedResults[coord] = result;
    return result;
  }
  
  /// Invalidate cache when a cell changes
  void invalidateCell(CellCoordinate coord) {
    _cachedResults.remove(coord);
    
    // Also invalidate all cells that depend on this one
    final cellA1 = A1.fromCoordinates(coord.row, coord.column);
    for (final dependent in _dependencies.getCellsToRecalculate(cellA1)) {
      _cachedResults.remove(CellCoordinate(dependent.row, dependent.column));
    }
  }
  
  /// Clear all cached results
  void clearCache() => _cachedResults.clear();
  
  // Delegate all other WorksheetData methods to _inner
  @override
  int get rowCount => _inner.rowCount;
  
  @override
  int get columnCount => _inner.columnCount;
  
  // ... other delegated methods ...
  
  Iterable<CellCoordinate> _referenceToCells(A1Reference ref) sync* {
    final range = ref.range;
    if (range == null) return;
    
    for (var row = range.from.row; row <= range.to.row; row++) {
      for (var col = range.from.column; col <= range.to.column; col++) {
        yield CellCoordinate(row, col);
      }
    }
  }
  
  Object? _formulaValueToObject(FormulaValue value) {
    return switch (value) {
      NumberValue(value: final n) => n,
      TextValue(value: final s) => s,
      BooleanValue(value: final b) => b,
      ErrorValue(error: final e) => e.code,
      EmptyValue() => null,
      RangeValue() => value.toText(),
    };
  }
}
```

---

## Complete Example Application

```dart
// example/lib/main.dart

import 'package:flutter/material.dart';
import 'package:worksheet/worksheet.dart';
import 'package:worksheet_formula/worksheet_formula.dart';
import 'worksheet_formula_integration.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatefulWidget {
  const MyApp({super.key});
  
  @override
  State<MyApp> createState() => _MyAppState();
}

class _MyAppState extends State<MyApp> {
  late final FormulaEngine _engine;
  late final SparseWorksheetData _rawData;
  late final FormulaAwareWorksheetData _data;
  
  @override
  void initState() {
    super.initState();
    
    // Create engine with a custom function
    _engine = FormulaEngine();
    _engine.registerFunction(DiscountFunction());
    
    // Create worksheet data with formulas
    _rawData = SparseWorksheetData(
      rowCount: 1000,
      columnCount: 26,
      cells: {
        // Headers
        (0, 0): Cell.text('Item'),
        (0, 1): Cell.text('Price'),
        (0, 2): Cell.text('Qty'),
        (0, 3): Cell.text('Total'),
        
        // Data rows
        (1, 0): Cell.text('Apples'),
        (1, 1): Cell.number(1.50),
        (1, 2): Cell.number(10),
        (1, 3): '=B2*C2'.formula,
        
        (2, 0): Cell.text('Oranges'),
        (2, 1): Cell.number(2.00),
        (2, 2): Cell.number(5),
        (2, 3): '=B3*C3'.formula,
        
        (3, 0): Cell.text('Bananas'),
        (3, 1): Cell.number(0.75),
        (3, 2): Cell.number(12),
        (3, 3): '=B4*C4'.formula,
        
        // Summary
        (5, 2): Cell.text('Subtotal:'),
        (5, 3): '=SUM(D2:D4)'.formula,
        
        (6, 2): Cell.text('Tax (10%):'),
        (6, 3): '=D6*0.1'.formula,
        
        (7, 2): Cell.text('Total:'),
        (7, 3): '=D6+D7'.formula,
        
        // Using custom function
        (9, 2): Cell.text('With 15% discount:'),
        (9, 3): '=DISCOUNT(D8, 0.15)'.formula,
      },
    );
    
    // Wrap with formula support
    _data = FormulaAwareWorksheetData(_rawData, engine: _engine);
  }
  
  @override
  void dispose() {
    _rawData.dispose();
    super.dispose();
  }
  
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Worksheet Formula Demo',
      home: Scaffold(
        appBar: AppBar(title: const Text('Worksheet Formula Demo')),
        body: WorksheetTheme(
          data: const WorksheetThemeData(),
          child: Worksheet(
            data: _data,
            rowCount: 1000,
            columnCount: 26,
            onCellTap: (cell) {
              final cellData = _data.getCell(cell);
              if (cellData != null) {
                print('Cell ${cell.toNotation()}: ${cellData.value}');
                if (cellData.isFormula) {
                  print('  Formula: ${cellData.formula}');
                }
              }
            },
          ),
        ),
      ),
    );
  }
}

/// Example custom function: DISCOUNT(price, rate)
class DiscountFunction extends FormulaFunction {
  @override
  String get name => 'DISCOUNT';
  
  @override
  int get minArgs => 2;
  
  @override
  int get maxArgs => 2;
  
  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    
    // Check for errors
    for (final v in values) {
      if (v.isError) return v;
    }
    
    final price = values[0].toNumber();
    final rate = values[1].toNumber();
    
    if (price == null || rate == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    
    if (rate < 0 || rate > 1) {
      return const FormulaValue.error(FormulaError.num);
    }
    
    return FormulaValue.number(price * (1 - rate));
  }
}
```

---

## Summary

### Key Benefits

1. **Standalone Package**: `worksheet_formula` has no dependency on `worksheet` or any UI framework
2. **Leverages Existing Packages**: Uses `a1` for cell references and `petitparser` for parsing
3. **Extensible**: Easy to add custom functions via the registry
4. **Clean Architecture**: Separation between parsing, AST, and evaluation
5. **Excel Compatible**: Matches Excel/Sheets operator precedence, error types, and type coercion
6. **Efficient**: Parse caching and dependency tracking for minimal recalculation
7. **Type-Safe**: Sealed classes for AST nodes and formula values

### Integration Points

- **`EvaluationContext`**: The abstract interface consumers implement to connect their data
- **`FunctionRegistry`**: Where custom functions are registered
- **`DependencyGraph`**: Optional dependency tracking for efficient updates

### Suggested Next Steps

1. Implement additional built-in functions (VLOOKUP, INDEX, MATCH, DATE functions, etc.)
2. Add support for array formulas
3. Implement named ranges
4. Add locale-aware number/date parsing
5. Create comprehensive test suite
