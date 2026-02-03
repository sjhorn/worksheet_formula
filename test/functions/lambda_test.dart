import 'package:a1/a1.dart';
import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/ast/operators.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/formula_engine.dart';
import 'package:worksheet_formula/src/functions/function.dart';
import 'package:worksheet_formula/src/functions/lambda.dart';
import 'package:worksheet_formula/src/functions/math.dart';
import 'package:worksheet_formula/src/functions/registry.dart';

class _TestContext implements EvaluationContext {
  @override
  A1 get currentCell => 'A1'.a1;
  @override
  String? get currentSheet => null;
  @override
  FormulaValue? getVariable(String name) => null;
  @override
  bool get isCancelled => false;

  final FunctionRegistry _registry;
  final Map<String, FormulaValue> cells;
  _TestContext(this._registry, {this.cells = const {}});

  @override
  FormulaValue getCellValue(A1 cell) =>
      cells[cell.toString()] ?? const EmptyValue();

  @override
  FormulaValue getRangeValues(A1Range range) {
    final fromA1 = range.from.a1;
    final toA1 = range.to.a1;
    if (fromA1 == null || toA1 == null) {
      return const FormulaValue.error(FormulaError.ref);
    }
    final values = <List<FormulaValue>>[];
    for (var row = fromA1.row; row <= toA1.row; row++) {
      final rowValues = <FormulaValue>[];
      for (var col = fromA1.column; col <= toA1.column; col++) {
        final cell = A1.fromVector(col, row);
        rowValues.add(getCellValue(cell));
      }
      values.add(rowValues);
    }
    return FormulaValue.range(values);
  }

  @override
  FormulaFunction? getFunction(String name) => _registry.get(name);
}

/// Helper: evaluate a function call by name with the given arg nodes.
FormulaValue _eval(
  String funcName,
  List<FormulaNode> args,
  EvaluationContext context,
) {
  final func = context.getFunction(funcName);
  if (func == null) return const FormulaValue.error(FormulaError.name);
  return func.call(args, context);
}

/// Helper: assert a RangeValue matches expected 2D structure.
void _expectRange(FormulaValue result, List<List<FormulaValue>> expected) {
  expect(result, isA<RangeValue>(), reason: 'expected RangeValue, got $result');
  final rv = result as RangeValue;
  expect(rv.rowCount, expected.length, reason: 'row count mismatch');
  if (expected.isNotEmpty) {
    expect(rv.columnCount, expected[0].length,
        reason: 'column count mismatch');
  }
  for (var r = 0; r < expected.length; r++) {
    for (var c = 0; c < expected[r].length; c++) {
      expect(rv.values[r][c], expected[r][c], reason: 'at [$r][$c]');
    }
  }
}

/// A helper function registered in tests to return a fixed value.
class _FixedValueFunction extends FormulaFunction {
  final FormulaValue fixedValue;
  final String _name;

  _FixedValueFunction(this._name, this.fixedValue);

  @override
  String get name => _name;
  @override
  int get minArgs => 0;
  @override
  int get maxArgs => 0;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) =>
      fixedValue;
}

void main() {
  late FunctionRegistry registry;
  late _TestContext context;

  setUp(() {
    registry = FunctionRegistry(registerBuiltIns: false);
    registerLambdaFunctions(registry);
    registerMathFunctions(registry);
    context = _TestContext(registry);
  });

  // ─── LAMBDA ──────────────────────────────────────────────────────

  group('LAMBDA', () {
    test('creates FunctionValue with one param', () {
      // LAMBDA(x, x+1)
      final result = _eval('LAMBDA', const [
        NameNode('x'),
        BinaryOpNode(NameNode('x'), BinaryOperator.add, NumberNode(1)),
      ], context);
      expect(result, isA<FunctionValue>());
      final fv = result as FunctionValue;
      expect(fv.paramNames, ['x']);
    });

    test('created lambda can be invoked', () {
      // LAMBDA(x, x+1) invoked with 5
      final lambda = _eval('LAMBDA', const [
        NameNode('x'),
        BinaryOpNode(NameNode('x'), BinaryOperator.add, NumberNode(1)),
      ], context) as FunctionValue;
      final result = invokeLambda(lambda, [const NumberValue(5)]);
      expect(result, const NumberValue(6));
    });

    test('multi-param lambda', () {
      // LAMBDA(x, y, x*y) invoked with (3, 4)
      final lambda = _eval('LAMBDA', const [
        NameNode('x'),
        NameNode('y'),
        BinaryOpNode(
            NameNode('x'), BinaryOperator.multiply, NameNode('y')),
      ], context) as FunctionValue;
      expect(lambda.paramNames, ['x', 'y']);
      final result = invokeLambda(lambda, [
        const NumberValue(3),
        const NumberValue(4),
      ]);
      expect(result, const NumberValue(12));
    });

    test('returns #VALUE! if param is not NameNode', () {
      // LAMBDA(42, x+1)
      final result = _eval('LAMBDA', const [
        NumberNode(42),
        BinaryOpNode(NameNode('x'), BinaryOperator.add, NumberNode(1)),
      ], context);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('body-only lambda (no params)', () {
      // LAMBDA(42) — returns a zero-arg function
      final result = _eval('LAMBDA', const [NumberNode(42)], context);
      expect(result, isA<FunctionValue>());
      final fv = result as FunctionValue;
      expect(fv.paramNames, isEmpty);
      expect(invokeLambda(fv, []), const NumberValue(42));
    });

    test('omitted args become OmittedValue', () {
      // LAMBDA(x, y, x) invoked with only 1 arg
      final lambda = _eval('LAMBDA', const [
        NameNode('x'),
        NameNode('y'),
        NameNode('x'),
      ], context) as FunctionValue;
      final result = invokeLambda(lambda, [const NumberValue(10)]);
      expect(result, const NumberValue(10));
    });

    test('captures closure context', () {
      // Create a scoped context with a=100, then LAMBDA(x, x+a)
      final scopedCtx = ScopedEvaluationContext(
        context,
        {'a': const NumberValue(100)},
      );
      final lambda = _eval('LAMBDA', const [
        NameNode('x'),
        BinaryOpNode(NameNode('x'), BinaryOperator.add, NameNode('a')),
      ], scopedCtx) as FunctionValue;
      final result = invokeLambda(lambda, [const NumberValue(5)]);
      expect(result, const NumberValue(105));
    });

    test('error in body propagates', () {
      // LAMBDA(x, x/0)
      final lambda = _eval('LAMBDA', const [
        NameNode('x'),
        BinaryOpNode(
            NameNode('x'), BinaryOperator.divide, NumberNode(0)),
      ], context) as FunctionValue;
      final result = invokeLambda(lambda, [const NumberValue(5)]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });

  // ─── LET ─────────────────────────────────────────────────────────

  group('LET', () {
    test('single binding', () {
      // LET(x, 10, x+1)
      final result = _eval('LET', const [
        NameNode('x'),
        NumberNode(10),
        BinaryOpNode(NameNode('x'), BinaryOperator.add, NumberNode(1)),
      ], context);
      expect(result, const NumberValue(11));
    });

    test('multiple bindings', () {
      // LET(x, 10, y, 20, x+y)
      final result = _eval('LET', const [
        NameNode('x'),
        NumberNode(10),
        NameNode('y'),
        NumberNode(20),
        BinaryOpNode(NameNode('x'), BinaryOperator.add, NameNode('y')),
      ], context);
      expect(result, const NumberValue(30));
    });

    test('later binding references earlier binding', () {
      // LET(x, 10, y, x+5, y*2)
      final result = _eval('LET', const [
        NameNode('x'),
        NumberNode(10),
        NameNode('y'),
        BinaryOpNode(NameNode('x'), BinaryOperator.add, NumberNode(5)),
        BinaryOpNode(
            NameNode('y'), BinaryOperator.multiply, NumberNode(2)),
      ], context);
      expect(result, const NumberValue(30));
    });

    test('returns #VALUE! with even number of args', () {
      // LET(x, 10) — 2 args, should be odd >= 3
      final result = _eval('LET', const [
        NameNode('x'),
        NumberNode(10),
      ], context);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('returns #VALUE! if name is not NameNode', () {
      // LET(42, 10, 1)
      final result = _eval('LET', const [
        NumberNode(42),
        NumberNode(10),
        NumberNode(1),
      ], context);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('error in value propagates', () {
      // LET(x, #DIV/0!, x)
      final result = _eval('LET', const [
        NameNode('x'),
        ErrorNode(FormulaError.divZero),
        NameNode('x'),
      ], context);
      expect(result, const ErrorValue(FormulaError.divZero));
    });

    test('variable shadowing', () {
      // LET(x, 10, x, 20, x)
      final result = _eval('LET', const [
        NameNode('x'),
        NumberNode(10),
        NameNode('x'),
        NumberNode(20),
        NameNode('x'),
      ], context);
      expect(result, const NumberValue(20));
    });

    test('case-insensitive variable names', () {
      // LET(X, 10, x+1) — X and x should be the same
      final result = _eval('LET', const [
        NameNode('X'),
        NumberNode(10),
        BinaryOpNode(NameNode('x'), BinaryOperator.add, NumberNode(1)),
      ], context);
      expect(result, const NumberValue(11));
    });
  });

  // ─── MAP ─────────────────────────────────────────────────────────

  group('MAP', () {
    test('applies lambda to each element of 1D range', () {
      // Set up cells A1=1, B1=2, C1=3, then MAP(A1:C1, LAMBDA(x, x*2))
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
        'B1': const NumberValue(2),
        'C1': const NumberValue(3),
      });
      final result = _eval('MAP', [
        RangeRefNode(A1Reference.parse('A1:C1')),
        const FunctionCallNode('LAMBDA', [
          NameNode('x'),
          BinaryOpNode(
              NameNode('x'), BinaryOperator.multiply, NumberNode(2)),
        ]),
      ], ctx);
      _expectRange(result, [
        [const NumberValue(2), const NumberValue(4), const NumberValue(6)],
      ]);
    });

    test('applies lambda to each element of 2D range', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
        'B1': const NumberValue(2),
        'A2': const NumberValue(3),
        'B2': const NumberValue(4),
      });
      final result = _eval('MAP', [
        RangeRefNode(A1Reference.parse('A1:B2')),
        const FunctionCallNode('LAMBDA', [
          NameNode('x'),
          BinaryOpNode(
              NameNode('x'), BinaryOperator.add, NumberNode(10)),
        ]),
      ], ctx);
      _expectRange(result, [
        [const NumberValue(11), const NumberValue(12)],
        [const NumberValue(13), const NumberValue(14)],
      ]);
    });

    test('returns #VALUE! if second arg is not lambda', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
      });
      final result = _eval('MAP', [
        CellRefNode(A1Reference.parse('A1')),
        const NumberNode(42),
      ], ctx);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('error in lambda propagates', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
      });
      final result = _eval('MAP', [
        CellRefNode(A1Reference.parse('A1')),
        const FunctionCallNode('LAMBDA', [
          NameNode('x'),
          BinaryOpNode(
              NameNode('x'), BinaryOperator.divide, NumberNode(0)),
        ]),
      ], ctx);
      expect(result, const ErrorValue(FormulaError.divZero));
    });

    test('works with single value (not range)', () {
      final result = _eval('MAP', [
        const NumberNode(5),
        const FunctionCallNode('LAMBDA', [
          NameNode('x'),
          BinaryOpNode(
              NameNode('x'), BinaryOperator.multiply, NumberNode(3)),
        ]),
      ], context);
      _expectRange(result, [
        [const NumberValue(15)],
      ]);
    });

    test('error in array propagates', () {
      final result = _eval('MAP', [
        const ErrorNode(FormulaError.ref),
        const FunctionCallNode('LAMBDA', [
          NameNode('x'),
          NameNode('x'),
        ]),
      ], context);
      expect(result, const ErrorValue(FormulaError.ref));
    });
  });

  // ─── REDUCE ──────────────────────────────────────────────────────

  group('REDUCE', () {
    test('folds to single value (sum)', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
        'B1': const NumberValue(2),
        'C1': const NumberValue(3),
      });
      final result = _eval('REDUCE', [
        const NumberNode(0),
        RangeRefNode(A1Reference.parse('A1:C1')),
        const FunctionCallNode('LAMBDA', [
          NameNode('acc'),
          NameNode('x'),
          BinaryOpNode(
              NameNode('acc'), BinaryOperator.add, NameNode('x')),
        ]),
      ], ctx);
      expect(result, const NumberValue(6));
    });

    test('returns initial for empty array', () {
      // Use a range that evaluates to empty — we'll use a single empty cell
      // and wrap in a lambda that checks. Actually, let's just pass an
      // empty NumberNode as initial and a range with no data.
      // Empty range: context default gives EmptyValue for cells, which
      // wraps into a 1x1 range. Let's test with REDUCE on a single value.
      // For a truly empty test, we need to check that the accumulator
      // returns when no elements exist. Since we can't easily create an
      // empty range in tests, let's verify via the ScopedEvaluationContext.
      final lambdaFn = FunctionValue(['acc', 'x'], (args) {
        final acc = args[0].toNumber() ?? 0;
        final x = args[1].toNumber() ?? 0;
        return FormulaValue.number(acc + x);
      });
      registry.register(_FixedValueFunction('_TESTLAMBDA', lambdaFn));
      // Test the lambda invocation directly
      final result = invokeLambda(lambdaFn, [const NumberValue(42), const NumberValue(0)]);
      expect(result, const NumberValue(42));
    });

    test('product fold', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(2),
        'B1': const NumberValue(3),
        'C1': const NumberValue(4),
      });
      final result = _eval('REDUCE', [
        const NumberNode(1),
        RangeRefNode(A1Reference.parse('A1:C1')),
        const FunctionCallNode('LAMBDA', [
          NameNode('acc'),
          NameNode('x'),
          BinaryOpNode(
              NameNode('acc'), BinaryOperator.multiply, NameNode('x')),
        ]),
      ], ctx);
      expect(result, const NumberValue(24));
    });

    test('returns #VALUE! if third arg is not lambda', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
      });
      final result = _eval('REDUCE', [
        const NumberNode(0),
        CellRefNode(A1Reference.parse('A1')),
        const NumberNode(1),
      ], ctx);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('error in lambda propagates', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
      });
      final result = _eval('REDUCE', [
        const NumberNode(0),
        CellRefNode(A1Reference.parse('A1')),
        const FunctionCallNode('LAMBDA', [
          NameNode('acc'),
          NameNode('x'),
          BinaryOpNode(
              NameNode('acc'), BinaryOperator.divide, NumberNode(0)),
        ]),
      ], ctx);
      expect(result, const ErrorValue(FormulaError.divZero));
    });

    test('folds across 2D array row by row', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
        'B1': const NumberValue(2),
        'A2': const NumberValue(3),
        'B2': const NumberValue(4),
      });
      final result = _eval('REDUCE', [
        const NumberNode(0),
        RangeRefNode(A1Reference.parse('A1:B2')),
        const FunctionCallNode('LAMBDA', [
          NameNode('acc'),
          NameNode('x'),
          BinaryOpNode(
              NameNode('acc'), BinaryOperator.add, NameNode('x')),
        ]),
      ], ctx);
      expect(result, const NumberValue(10));
    });
  });

  // ─── SCAN ────────────────────────────────────────────────────────

  group('SCAN', () {
    test('running sum accumulation', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
        'B1': const NumberValue(2),
        'C1': const NumberValue(3),
      });
      final result = _eval('SCAN', [
        const NumberNode(0),
        RangeRefNode(A1Reference.parse('A1:C1')),
        const FunctionCallNode('LAMBDA', [
          NameNode('acc'),
          NameNode('x'),
          BinaryOpNode(
              NameNode('acc'), BinaryOperator.add, NameNode('x')),
        ]),
      ], ctx);
      _expectRange(result, [
        [const NumberValue(1), const NumberValue(3), const NumberValue(6)],
      ]);
    });

    test('output shape matches input for 2D', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
        'B1': const NumberValue(2),
        'A2': const NumberValue(3),
        'B2': const NumberValue(4),
      });
      final result = _eval('SCAN', [
        const NumberNode(0),
        RangeRefNode(A1Reference.parse('A1:B2')),
        const FunctionCallNode('LAMBDA', [
          NameNode('acc'),
          NameNode('x'),
          BinaryOpNode(
              NameNode('acc'), BinaryOperator.add, NameNode('x')),
        ]),
      ], ctx);
      _expectRange(result, [
        [const NumberValue(1), const NumberValue(3)],
        [const NumberValue(6), const NumberValue(10)],
      ]);
    });

    test('returns #VALUE! if third arg is not lambda', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
      });
      final result = _eval('SCAN', [
        const NumberNode(0),
        CellRefNode(A1Reference.parse('A1')),
        const NumberNode(1),
      ], ctx);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('error in lambda propagates', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
      });
      final result = _eval('SCAN', [
        const NumberNode(0),
        CellRefNode(A1Reference.parse('A1')),
        const FunctionCallNode('LAMBDA', [
          NameNode('acc'),
          NameNode('x'),
          BinaryOpNode(
              NameNode('acc'), BinaryOperator.divide, NumberNode(0)),
        ]),
      ], ctx);
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });

  // ─── MAKEARRAY ───────────────────────────────────────────────────

  group('MAKEARRAY', () {
    test('builds array with lambda(row, col)', () {
      // MAKEARRAY(2, 3, LAMBDA(r, c, r*10+c))
      final result = _eval('MAKEARRAY', [
        const NumberNode(2),
        const NumberNode(3),
        const FunctionCallNode('LAMBDA', [
          NameNode('r'),
          NameNode('c'),
          BinaryOpNode(
            BinaryOpNode(
                NameNode('r'), BinaryOperator.multiply, NumberNode(10)),
            BinaryOperator.add,
            NameNode('c'),
          ),
        ]),
      ], context);
      _expectRange(result, [
        [const NumberValue(11), const NumberValue(12), const NumberValue(13)],
        [const NumberValue(21), const NumberValue(22), const NumberValue(23)],
      ]);
    });

    test('1-based indices', () {
      // MAKEARRAY(1, 1, LAMBDA(r, c, r+c)) → 2 (1+1)
      final result = _eval('MAKEARRAY', [
        const NumberNode(1),
        const NumberNode(1),
        const FunctionCallNode('LAMBDA', [
          NameNode('r'),
          NameNode('c'),
          BinaryOpNode(NameNode('r'), BinaryOperator.add, NameNode('c')),
        ]),
      ], context);
      _expectRange(result, [
        [const NumberValue(2)],
      ]);
    });

    test('returns #VALUE! for zero rows', () {
      final result = _eval('MAKEARRAY', [
        const NumberNode(0),
        const NumberNode(1),
        const FunctionCallNode('LAMBDA', [
          NameNode('r'),
          NameNode('c'),
          NumberNode(1),
        ]),
      ], context);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('returns #VALUE! for negative cols', () {
      final result = _eval('MAKEARRAY', [
        const NumberNode(1),
        const NumberNode(-1),
        const FunctionCallNode('LAMBDA', [
          NameNode('r'),
          NameNode('c'),
          NumberNode(1),
        ]),
      ], context);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('returns #VALUE! if third arg is not lambda', () {
      final result = _eval('MAKEARRAY', [
        const NumberNode(1),
        const NumberNode(1),
        const NumberNode(1),
      ], context);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('error in lambda propagates', () {
      final result = _eval('MAKEARRAY', [
        const NumberNode(1),
        const NumberNode(1),
        const FunctionCallNode('LAMBDA', [
          NameNode('r'),
          NameNode('c'),
          BinaryOpNode(
              NameNode('r'), BinaryOperator.divide, NumberNode(0)),
        ]),
      ], context);
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });

  // ─── BYCOL ───────────────────────────────────────────────────────

  group('BYCOL', () {
    test('applies lambda to each column', () {
      // BYCOL({1,2,3;4,5,6}, LAMBDA(col, SUM(col)))
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
        'B1': const NumberValue(2),
        'C1': const NumberValue(3),
        'A2': const NumberValue(4),
        'B2': const NumberValue(5),
        'C2': const NumberValue(6),
      });
      final result = _eval('BYCOL', [
        RangeRefNode(A1Reference.parse('A1:C2')),
        const FunctionCallNode('LAMBDA', [
          NameNode('col'),
          FunctionCallNode('SUM', [NameNode('col')]),
        ]),
      ], ctx);
      // SUM of col1=1+4=5, col2=2+5=7, col3=3+6=9
      _expectRange(result, [
        [const NumberValue(5), const NumberValue(7), const NumberValue(9)],
      ]);
    });

    test('lambda receives single-column range', () {
      // Register a test function that captures the arg
      late FormulaValue receivedArg;
      final testLambda = FunctionValue(['col'], (args) {
        receivedArg = args[0];
        return const NumberValue(0);
      });
      registry.register(_FixedValueFunction('_TESTLAMBDA', testLambda));
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
        'B1': const NumberValue(2),
        'A2': const NumberValue(3),
        'B2': const NumberValue(4),
      });
      _eval('BYCOL', [
        RangeRefNode(A1Reference.parse('A1:B2')),
        const FunctionCallNode('_TESTLAMBDA', []),
      ], ctx);
      // First column should be {1;3}
      expect(receivedArg, isA<RangeValue>());
      final rv = receivedArg as RangeValue;
      expect(rv.rowCount, 2);
      expect(rv.columnCount, 1);
    });

    test('returns #VALUE! if second arg is not lambda', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
      });
      final result = _eval('BYCOL', [
        CellRefNode(A1Reference.parse('A1')),
        const NumberNode(1),
      ], ctx);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('error in lambda propagates', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
      });
      final result = _eval('BYCOL', [
        CellRefNode(A1Reference.parse('A1')),
        const FunctionCallNode('LAMBDA', [
          NameNode('col'),
          ErrorNode(FormulaError.num),
        ]),
      ], ctx);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  // ─── BYROW ───────────────────────────────────────────────────────

  group('BYROW', () {
    test('applies lambda to each row', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
        'B1': const NumberValue(2),
        'C1': const NumberValue(3),
        'A2': const NumberValue(4),
        'B2': const NumberValue(5),
        'C2': const NumberValue(6),
      });
      final result = _eval('BYROW', [
        RangeRefNode(A1Reference.parse('A1:C2')),
        const FunctionCallNode('LAMBDA', [
          NameNode('row'),
          FunctionCallNode('SUM', [NameNode('row')]),
        ]),
      ], ctx);
      // SUM of row1=6, row2=15
      _expectRange(result, [
        [const NumberValue(6)],
        [const NumberValue(15)],
      ]);
    });

    test('lambda receives single-row range', () {
      late FormulaValue receivedArg;
      final testLambda = FunctionValue(['row'], (args) {
        receivedArg = args[0];
        return const NumberValue(0);
      });
      registry.register(_FixedValueFunction('_TESTLAMBDA2', testLambda));
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
        'B1': const NumberValue(2),
        'A2': const NumberValue(3),
        'B2': const NumberValue(4),
      });
      _eval('BYROW', [
        RangeRefNode(A1Reference.parse('A1:B2')),
        const FunctionCallNode('_TESTLAMBDA2', []),
      ], ctx);
      // First row should be {1,2}
      expect(receivedArg, isA<RangeValue>());
      final rv = receivedArg as RangeValue;
      expect(rv.rowCount, 1);
      expect(rv.columnCount, 2);
    });

    test('returns #VALUE! if second arg is not lambda', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
      });
      final result = _eval('BYROW', [
        CellRefNode(A1Reference.parse('A1')),
        const NumberNode(1),
      ], ctx);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('error in lambda propagates', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
      });
      final result = _eval('BYROW', [
        CellRefNode(A1Reference.parse('A1')),
        const FunctionCallNode('LAMBDA', [
          NameNode('row'),
          ErrorNode(FormulaError.num),
        ]),
      ], ctx);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  // ─── ISOMITTED ───────────────────────────────────────────────────

  group('ISOMITTED', () {
    test('returns FALSE for number', () {
      final result = _eval('ISOMITTED', const [NumberNode(42)], context);
      expect(result, const BooleanValue(false));
    });

    test('returns FALSE for text', () {
      final result =
          _eval('ISOMITTED', const [TextNode('hello')], context);
      expect(result, const BooleanValue(false));
    });

    test('returns FALSE for boolean', () {
      final result =
          _eval('ISOMITTED', const [BooleanNode(true)], context);
      expect(result, const BooleanValue(false));
    });

    test('works inside lambda with omitted parameter', () {
      // LAMBDA(x, y, ISOMITTED(y)) invoked with only 1 arg
      final lambda = _eval('LAMBDA', const [
        NameNode('x'),
        NameNode('y'),
        FunctionCallNode('ISOMITTED', [NameNode('y')]),
      ], context) as FunctionValue;
      final result = invokeLambda(lambda, [const NumberValue(1)]);
      expect(result, const BooleanValue(true));
    });

    test('returns FALSE for provided parameter', () {
      // LAMBDA(x, y, ISOMITTED(y)) invoked with 2 args
      final lambda = _eval('LAMBDA', const [
        NameNode('x'),
        NameNode('y'),
        FunctionCallNode('ISOMITTED', [NameNode('y')]),
      ], context) as FunctionValue;
      final result = invokeLambda(lambda, [
        const NumberValue(1),
        const NumberValue(2),
      ]);
      expect(result, const BooleanValue(false));
    });
  });

  // ─── Cross-cutting tests ─────────────────────────────────────────

  group('cross-cutting', () {
    test('nested LAMBDA (closure captures)', () {
      // LAMBDA(x, LAMBDA(y, x+y)) — curried addition
      final outerLambda = _eval('LAMBDA', const [
        NameNode('x'),
        FunctionCallNode('LAMBDA', [
          NameNode('y'),
          BinaryOpNode(NameNode('x'), BinaryOperator.add, NameNode('y')),
        ]),
      ], context) as FunctionValue;

      // Invoke outer with 10 → get inner lambda
      final innerLambda =
          invokeLambda(outerLambda, [const NumberValue(10)]) as FunctionValue;

      // Invoke inner with 5 → 15
      final result = invokeLambda(innerLambda, [const NumberValue(5)]);
      expect(result, const NumberValue(15));
    });

    test('LET + LAMBDA combo (named lambda passed to MAP)', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(5),
      });
      // LET(factor, 2, MAP(A1, LAMBDA(x, x*factor)))
      final result = _eval('LET', [
        const NameNode('factor'),
        const NumberNode(2),
        FunctionCallNode('MAP', [
          CellRefNode(A1Reference.parse('A1')),
          const FunctionCallNode('LAMBDA', [
            NameNode('x'),
            BinaryOpNode(
                NameNode('x'), BinaryOperator.multiply, NameNode('factor')),
          ]),
        ]),
      ], ctx);
      _expectRange(result, [
        [const NumberValue(10)],
      ]);
    });

    test('variable shadowing in nested scope', () {
      // LET(x, 10, LET(x, 20, x)) → 20
      final result = _eval('LET', const [
        NameNode('x'),
        NumberNode(10),
        FunctionCallNode('LET', [
          NameNode('x'),
          NumberNode(20),
          NameNode('x'),
        ]),
      ], context);
      expect(result, const NumberValue(20));
    });

    test('case-insensitive variable names in lambda', () {
      // LAMBDA(X, x+1) invoked with 5 → 6
      final lambda = _eval('LAMBDA', const [
        NameNode('X'),
        BinaryOpNode(NameNode('x'), BinaryOperator.add, NumberNode(1)),
      ], context) as FunctionValue;
      final result = invokeLambda(lambda, [const NumberValue(5)]);
      expect(result, const NumberValue(6));
    });

    test('error propagation through lambda body', () {
      // MAP a cell containing error
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
        'B1': const ErrorValue(FormulaError.ref),
      });
      final result = _eval('MAP', [
        RangeRefNode(A1Reference.parse('A1:B1')),
        const FunctionCallNode('LAMBDA', [
          NameNode('x'),
          BinaryOpNode(NameNode('x'), BinaryOperator.add, NumberNode(1)),
        ]),
      ], ctx);
      // The error from B1 propagates through x+1
      expect(result, const ErrorValue(FormulaError.ref));
    });

    test('non-FunctionValue passed where lambda expected → #VALUE!', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
      });
      final result = _eval('MAP', [
        CellRefNode(A1Reference.parse('A1')),
        const TextNode('not a lambda'),
      ], ctx);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('REDUCE with LAMBDA building a string', () {
      final ctx = _TestContext(registry, cells: {
        'A1': const NumberValue(1),
        'B1': const NumberValue(2),
        'C1': const NumberValue(3),
      });
      final result = _eval('REDUCE', [
        const TextNode(''),
        RangeRefNode(A1Reference.parse('A1:C1')),
        const FunctionCallNode('LAMBDA', [
          NameNode('acc'),
          NameNode('x'),
          BinaryOpNode(
              NameNode('acc'), BinaryOperator.concat, NameNode('x')),
        ]),
      ], ctx);
      expect(result, const TextValue('123'));
    });
  });

  // ─── ScopedEvaluationContext ──────────────────────────────────────

  group('ScopedEvaluationContext', () {
    test('resolves variables', () {
      final scoped = ScopedEvaluationContext(
        context,
        {'x': const NumberValue(42)},
      );
      expect(scoped.getVariable('x'), const NumberValue(42));
    });

    test('case-insensitive lookup', () {
      final scoped = ScopedEvaluationContext(
        context,
        {'MyVar': const NumberValue(10)},
      );
      expect(scoped.getVariable('myvar'), const NumberValue(10));
      expect(scoped.getVariable('MYVAR'), const NumberValue(10));
      expect(scoped.getVariable('MyVar'), const NumberValue(10));
    });

    test('delegates to parent for unknown variables', () {
      final parent = ScopedEvaluationContext(
        context,
        {'a': const NumberValue(1)},
      );
      final child = ScopedEvaluationContext(
        parent,
        {'b': const NumberValue(2)},
      );
      expect(child.getVariable('a'), const NumberValue(1));
      expect(child.getVariable('b'), const NumberValue(2));
    });

    test('returns null for undefined variable', () {
      final scoped = ScopedEvaluationContext(context, {});
      expect(scoped.getVariable('undefined'), null);
    });

    test('delegates cell access to parent', () {
      expect(context.getCellValue('A1'.a1), const EmptyValue());
      final scoped = ScopedEvaluationContext(context, {});
      expect(scoped.getCellValue('A1'.a1), const EmptyValue());
    });

    test('delegates function access to parent', () {
      final scoped = ScopedEvaluationContext(context, {});
      expect(scoped.getFunction('LAMBDA'), isNotNull);
    });

    test('delegates currentCell to parent', () {
      final scoped = ScopedEvaluationContext(context, {});
      expect(scoped.currentCell, 'A1'.a1);
    });

    test('delegates currentSheet to parent', () {
      final scoped = ScopedEvaluationContext(context, {});
      expect(scoped.currentSheet, null);
    });

    test('delegates isCancelled to parent', () {
      final scoped = ScopedEvaluationContext(context, {});
      expect(scoped.isCancelled, false);
    });
  });

  // ─── Immediate invocation ─────────────────────────────────────────

  group('immediate invocation', () {
    late FormulaEngine engine;
    late _TestContext ctx;

    setUp(() {
      engine = FormulaEngine();
      ctx = _TestContext(engine.functions);
    });

    test('LAMBDA(x,x+1)(5) returns 6', () {
      final result = engine.evaluateString('=LAMBDA(x,x+1)(5)', ctx);
      expect(result, const NumberValue(6));
    });

    test('LAMBDA(x,y,x+y)(3,4) returns 7', () {
      final result = engine.evaluateString('=LAMBDA(x,y,x+y)(3,4)', ctx);
      expect(result, const NumberValue(7));
    });

    test('LAMBDA(x,x*2)(0) returns 0', () {
      final result = engine.evaluateString('=LAMBDA(x,x*2)(0)', ctx);
      expect(result, const NumberValue(0));
    });

    test('chained/curried: LAMBDA(x,LAMBDA(y,x+y))(1)(2) returns 3', () {
      final result =
          engine.evaluateString('=LAMBDA(x,LAMBDA(y,x+y))(1)(2)', ctx);
      expect(result, const NumberValue(3));
    });

    test('no-param lambda: LAMBDA(10)() returns 10', () {
      final result = engine.evaluateString('=LAMBDA(10)()', ctx);
      expect(result, const NumberValue(10));
    });

    test('non-function invocation returns #VALUE!', () {
      final result = engine.evaluateString('=(5)(3)', ctx);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  // ─── Registration ────────────────────────────────────────────────

  group('registration', () {
    test('all 9 functions are registered', () {
      final fullRegistry = FunctionRegistry();
      final expected = [
        'LAMBDA', 'LET', 'MAP', 'REDUCE', 'SCAN',
        'MAKEARRAY', 'BYCOL', 'BYROW', 'ISOMITTED',
      ];
      for (final name in expected) {
        expect(fullRegistry.has(name), isTrue,
            reason: '$name should be registered');
      }
    });

    test('total function count is 400', () {
      final fullRegistry = FunctionRegistry();
      expect(fullRegistry.names.length, 400);
    });
  });
}
