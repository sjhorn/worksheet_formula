import 'package:a1/a1.dart';
import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/functions/array.dart';
import 'package:worksheet_formula/src/functions/function.dart';
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
  Map<String, FormulaValue> rangeMap = {};
  _TestContext(this._registry);

  @override
  FormulaValue getCellValue(A1 cell) => const EmptyValue();

  @override
  FormulaValue getRangeValues(A1Range range) {
    final key = range.toString();
    if (rangeMap.containsKey(key)) return rangeMap[key]!;
    return const FormulaValue.error(FormulaError.ref);
  }

  @override
  FormulaFunction? getFunction(String name) => _registry.get(name);
}

/// Helper to assert a RangeValue matches expected 2D structure.
void expectRange(FormulaValue result, List<List<FormulaValue>> expected) {
  expect(result, isA<RangeValue>(), reason: 'expected RangeValue, got $result');
  final rv = result as RangeValue;
  expect(rv.rowCount, expected.length, reason: 'row count mismatch');
  if (expected.isNotEmpty) {
    expect(rv.columnCount, expected[0].length, reason: 'column count mismatch');
  }
  for (var r = 0; r < expected.length; r++) {
    for (var c = 0; c < expected[r].length; c++) {
      expect(rv.values[r][c], expected[r][c], reason: 'at [$r][$c]');
    }
  }
}

void main() {
  late FunctionRegistry registry;
  late _TestContext context;

  setUp(() {
    registry = FunctionRegistry(registerBuiltIns: false);
    registerArrayFunctions(registry);
    context = _TestContext(registry);
  });

  FormulaValue eval(FormulaFunction fn, List<FormulaNode> args) =>
      fn.call(args, context);

  // ─── Wave A — Generators ─────────────────────────────────

  group('SEQUENCE', () {
    test('single row sequence', () {
      final result = eval(registry.get('SEQUENCE')!, [
        const NumberNode(1),
        const NumberNode(5),
      ]);
      expectRange(result, [
        [
          const NumberValue(1),
          const NumberValue(2),
          const NumberValue(3),
          const NumberValue(4),
          const NumberValue(5),
        ],
      ]);
    });

    test('single column sequence', () {
      final result = eval(registry.get('SEQUENCE')!, [
        const NumberNode(3),
      ]);
      expectRange(result, [
        [const NumberValue(1)],
        [const NumberValue(2)],
        [const NumberValue(3)],
      ]);
    });

    test('2x3 grid with custom start and step', () {
      final result = eval(registry.get('SEQUENCE')!, [
        const NumberNode(2),
        const NumberNode(3),
        const NumberNode(10),
        const NumberNode(5),
      ]);
      expectRange(result, [
        [const NumberValue(10), const NumberValue(15), const NumberValue(20)],
        [const NumberValue(25), const NumberValue(30), const NumberValue(35)],
      ]);
    });

    test('negative step', () {
      final result = eval(registry.get('SEQUENCE')!, [
        const NumberNode(1),
        const NumberNode(3),
        const NumberNode(10),
        const NumberNode(-3),
      ]);
      expectRange(result, [
        [const NumberValue(10), const NumberValue(7), const NumberValue(4)],
      ]);
    });

    test('rows < 1 returns #VALUE!', () {
      final result = eval(registry.get('SEQUENCE')!, [const NumberNode(0)]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('cols < 1 returns #VALUE!', () {
      final result = eval(registry.get('SEQUENCE')!, [
        const NumberNode(1),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('RANDARRAY', () {
    test('default 1x1', () {
      final result = eval(registry.get('RANDARRAY')!, []);
      expect(result, isA<RangeValue>());
      final rv = result as RangeValue;
      expect(rv.rowCount, 1);
      expect(rv.columnCount, 1);
      final val = (rv.values[0][0] as NumberValue).value;
      expect(val, greaterThanOrEqualTo(0));
      expect(val, lessThan(1));
    });

    test('3x2 with bounds', () {
      final result = eval(registry.get('RANDARRAY')!, [
        const NumberNode(3),
        const NumberNode(2),
        const NumberNode(10),
        const NumberNode(20),
      ]);
      expect(result, isA<RangeValue>());
      final rv = result as RangeValue;
      expect(rv.rowCount, 3);
      expect(rv.columnCount, 2);
      for (final row in rv.values) {
        for (final cell in row) {
          final val = (cell as NumberValue).value;
          expect(val, greaterThanOrEqualTo(10));
          expect(val, lessThanOrEqualTo(20));
        }
      }
    });

    test('whole numbers', () {
      final result = eval(registry.get('RANDARRAY')!, [
        const NumberNode(2),
        const NumberNode(2),
        const NumberNode(1),
        const NumberNode(100),
        const BooleanNode(true),
      ]);
      expect(result, isA<RangeValue>());
      final rv = result as RangeValue;
      for (final row in rv.values) {
        for (final cell in row) {
          final val = (cell as NumberValue).value;
          expect(val, equals(val.toInt()));
          expect(val, greaterThanOrEqualTo(1));
          expect(val, lessThanOrEqualTo(100));
        }
      }
    });

    test('min > max returns #VALUE!', () {
      final result = eval(registry.get('RANDARRAY')!, [
        const NumberNode(1),
        const NumberNode(1),
        const NumberNode(10),
        const NumberNode(5),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('rows < 1 returns #VALUE!', () {
      final result = eval(registry.get('RANDARRAY')!, [const NumberNode(0)]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  // ─── Wave B — Flatten/Reshape ─────────────────────────────

  group('TOCOL', () {
    test('flattens 2x2 row-by-row', () {
      context.rangeMap['A1:B2'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
      ]);
      final result = eval(registry.get('TOCOL')!, [
        RangeRefNode(A1Reference.parse('A1:B2')),
      ]);
      expectRange(result, [
        [const NumberValue(1)],
        [const NumberValue(2)],
        [const NumberValue(3)],
        [const NumberValue(4)],
      ]);
    });

    test('flattens by column when scan_by_col=TRUE', () {
      context.rangeMap['A1:B2'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
      ]);
      final result = eval(registry.get('TOCOL')!, [
        RangeRefNode(A1Reference.parse('A1:B2')),
        const NumberNode(0),
        const BooleanNode(true),
      ]);
      expectRange(result, [
        [const NumberValue(1)],
        [const NumberValue(3)],
        [const NumberValue(2)],
        [const NumberValue(4)],
      ]);
    });

    test('ignore blanks (ignore=1)', () {
      context.rangeMap['A1:A3'] = const RangeValue([
        [NumberValue(1)],
        [EmptyValue()],
        [NumberValue(3)],
      ]);
      final result = eval(registry.get('TOCOL')!, [
        RangeRefNode(A1Reference.parse('A1:A3')),
        const NumberNode(1),
      ]);
      expectRange(result, [
        [const NumberValue(1)],
        [const NumberValue(3)],
      ]);
    });

    test('ignore errors (ignore=2)', () {
      context.rangeMap['A1:A3'] = const RangeValue([
        [NumberValue(1)],
        [ErrorValue(FormulaError.na)],
        [NumberValue(3)],
      ]);
      final result = eval(registry.get('TOCOL')!, [
        RangeRefNode(A1Reference.parse('A1:A3')),
        const NumberNode(2),
      ]);
      expectRange(result, [
        [const NumberValue(1)],
        [const NumberValue(3)],
      ]);
    });

    test('ignore blanks and errors (ignore=3)', () {
      context.rangeMap['A1:A4'] = const RangeValue([
        [NumberValue(1)],
        [EmptyValue()],
        [ErrorValue(FormulaError.na)],
        [NumberValue(4)],
      ]);
      final result = eval(registry.get('TOCOL')!, [
        RangeRefNode(A1Reference.parse('A1:A4')),
        const NumberNode(3),
      ]);
      expectRange(result, [
        [const NumberValue(1)],
        [const NumberValue(4)],
      ]);
    });

    test('all filtered returns #CALC!', () {
      context.rangeMap['A1:A2'] = const RangeValue([
        [EmptyValue()],
        [EmptyValue()],
      ]);
      final result = eval(registry.get('TOCOL')!, [
        RangeRefNode(A1Reference.parse('A1:A2')),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.calc));
    });

    test('scalar input becomes single column', () {
      final result = eval(registry.get('TOCOL')!, [const NumberNode(42)]);
      expectRange(result, [
        [const NumberValue(42)],
      ]);
    });
  });

  group('TOROW', () {
    test('flattens 2x2 row-by-row', () {
      context.rangeMap['A1:B2'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
      ]);
      final result = eval(registry.get('TOROW')!, [
        RangeRefNode(A1Reference.parse('A1:B2')),
      ]);
      expectRange(result, [
        [
          const NumberValue(1),
          const NumberValue(2),
          const NumberValue(3),
          const NumberValue(4),
        ],
      ]);
    });

    test('flattens by column when scan_by_col=TRUE', () {
      context.rangeMap['A1:B2'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
      ]);
      final result = eval(registry.get('TOROW')!, [
        RangeRefNode(A1Reference.parse('A1:B2')),
        const NumberNode(0),
        const BooleanNode(true),
      ]);
      expectRange(result, [
        [
          const NumberValue(1),
          const NumberValue(3),
          const NumberValue(2),
          const NumberValue(4),
        ],
      ]);
    });

    test('ignore blanks (ignore=1)', () {
      context.rangeMap['A1:C1'] = const RangeValue([
        [NumberValue(1), EmptyValue(), NumberValue(3)],
      ]);
      final result = eval(registry.get('TOROW')!, [
        RangeRefNode(A1Reference.parse('A1:C1')),
        const NumberNode(1),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(3)],
      ]);
    });

    test('all filtered returns #CALC!', () {
      context.rangeMap['A1:A2'] = const RangeValue([
        [ErrorValue(FormulaError.na)],
        [ErrorValue(FormulaError.value)],
      ]);
      final result = eval(registry.get('TOROW')!, [
        RangeRefNode(A1Reference.parse('A1:A2')),
        const NumberNode(2),
      ]);
      expect(result, const ErrorValue(FormulaError.calc));
    });
  });

  group('WRAPROWS', () {
    test('wraps vector into rows', () {
      context.rangeMap['A1:A6'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(6)],
      ]);
      final result = eval(registry.get('WRAPROWS')!, [
        RangeRefNode(A1Reference.parse('A1:A6')),
        const NumberNode(3),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(2), const NumberValue(3)],
        [const NumberValue(4), const NumberValue(5), const NumberValue(6)],
      ]);
    });

    test('pads incomplete last row with #N/A', () {
      context.rangeMap['A1:A5'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(5)],
      ]);
      final result = eval(registry.get('WRAPROWS')!, [
        RangeRefNode(A1Reference.parse('A1:A5')),
        const NumberNode(3),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(2), const NumberValue(3)],
        [
          const NumberValue(4),
          const NumberValue(5),
          const ErrorValue(FormulaError.na),
        ],
      ]);
    });

    test('custom pad value', () {
      context.rangeMap['A1:A5'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(5)],
      ]);
      final result = eval(registry.get('WRAPROWS')!, [
        RangeRefNode(A1Reference.parse('A1:A5')),
        const NumberNode(3),
        const NumberNode(0),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(2), const NumberValue(3)],
        [const NumberValue(4), const NumberValue(5), const NumberValue(0)],
      ]);
    });

    test('wrap_count < 1 returns #VALUE!', () {
      final result = eval(registry.get('WRAPROWS')!, [
        const NumberNode(1),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('scalar input', () {
      final result = eval(registry.get('WRAPROWS')!, [
        const NumberNode(42),
        const NumberNode(1),
      ]);
      expectRange(result, [
        [const NumberValue(42)],
      ]);
    });
  });

  group('WRAPCOLS', () {
    test('wraps vector into columns', () {
      context.rangeMap['A1:A6'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(6)],
      ]);
      final result = eval(registry.get('WRAPCOLS')!, [
        RangeRefNode(A1Reference.parse('A1:A6')),
        const NumberNode(3),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(4)],
        [const NumberValue(2), const NumberValue(5)],
        [const NumberValue(3), const NumberValue(6)],
      ]);
    });

    test('pads incomplete last column with #N/A', () {
      context.rangeMap['A1:A5'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(5)],
      ]);
      final result = eval(registry.get('WRAPCOLS')!, [
        RangeRefNode(A1Reference.parse('A1:A5')),
        const NumberNode(3),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(4)],
        [const NumberValue(2), const NumberValue(5)],
        [const NumberValue(3), const ErrorValue(FormulaError.na)],
      ]);
    });

    test('custom pad value', () {
      context.rangeMap['A1:A5'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(5)],
      ]);
      final result = eval(registry.get('WRAPCOLS')!, [
        RangeRefNode(A1Reference.parse('A1:A5')),
        const NumberNode(3),
        const TextNode('x'),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(4)],
        [const NumberValue(2), const NumberValue(5)],
        [const NumberValue(3), const TextValue('x')],
      ]);
    });

    test('wrap_count < 1 returns #VALUE!', () {
      final result = eval(registry.get('WRAPCOLS')!, [
        const NumberNode(1),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  // ─── Wave C — Slice/Select ────────────────────────────────

  group('CHOOSEROWS', () {
    test('selects specific rows', () {
      context.rangeMap['A1:B3'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
        [NumberValue(5), NumberValue(6)],
      ]);
      final result = eval(registry.get('CHOOSEROWS')!, [
        RangeRefNode(A1Reference.parse('A1:B3')),
        const NumberNode(1),
        const NumberNode(3),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(2)],
        [const NumberValue(5), const NumberValue(6)],
      ]);
    });

    test('negative index from end', () {
      context.rangeMap['A1:B3'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
        [NumberValue(5), NumberValue(6)],
      ]);
      final result = eval(registry.get('CHOOSEROWS')!, [
        RangeRefNode(A1Reference.parse('A1:B3')),
        const NumberNode(-1),
      ]);
      expectRange(result, [
        [const NumberValue(5), const NumberValue(6)],
      ]);
    });

    test('duplicate rows allowed', () {
      context.rangeMap['A1:A2'] = const RangeValue([
        [NumberValue(10)],
        [NumberValue(20)],
      ]);
      final result = eval(registry.get('CHOOSEROWS')!, [
        RangeRefNode(A1Reference.parse('A1:A2')),
        const NumberNode(1),
        const NumberNode(1),
        const NumberNode(2),
      ]);
      expectRange(result, [
        [const NumberValue(10)],
        [const NumberValue(10)],
        [const NumberValue(20)],
      ]);
    });

    test('index 0 returns #VALUE!', () {
      context.rangeMap['A1:A2'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      final result = eval(registry.get('CHOOSEROWS')!, [
        RangeRefNode(A1Reference.parse('A1:A2')),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('out of bounds returns #VALUE!', () {
      context.rangeMap['A1:A2'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      final result = eval(registry.get('CHOOSEROWS')!, [
        RangeRefNode(A1Reference.parse('A1:A2')),
        const NumberNode(5),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('CHOOSECOLS', () {
    test('selects specific columns', () {
      context.rangeMap['A1:C2'] = const RangeValue([
        [NumberValue(1), NumberValue(2), NumberValue(3)],
        [NumberValue(4), NumberValue(5), NumberValue(6)],
      ]);
      final result = eval(registry.get('CHOOSECOLS')!, [
        RangeRefNode(A1Reference.parse('A1:C2')),
        const NumberNode(1),
        const NumberNode(3),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(3)],
        [const NumberValue(4), const NumberValue(6)],
      ]);
    });

    test('negative index from end', () {
      context.rangeMap['A1:C1'] = const RangeValue([
        [NumberValue(10), NumberValue(20), NumberValue(30)],
      ]);
      final result = eval(registry.get('CHOOSECOLS')!, [
        RangeRefNode(A1Reference.parse('A1:C1')),
        const NumberNode(-1),
      ]);
      expectRange(result, [
        [const NumberValue(30)],
      ]);
    });

    test('index 0 returns #VALUE!', () {
      context.rangeMap['A1:B1'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
      ]);
      final result = eval(registry.get('CHOOSECOLS')!, [
        RangeRefNode(A1Reference.parse('A1:B1')),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('out of bounds returns #VALUE!', () {
      context.rangeMap['A1:B1'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
      ]);
      final result = eval(registry.get('CHOOSECOLS')!, [
        RangeRefNode(A1Reference.parse('A1:B1')),
        const NumberNode(5),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('DROP', () {
    test('drop rows from start', () {
      context.rangeMap['A1:B3'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
        [NumberValue(5), NumberValue(6)],
      ]);
      final result = eval(registry.get('DROP')!, [
        RangeRefNode(A1Reference.parse('A1:B3')),
        const NumberNode(1),
      ]);
      expectRange(result, [
        [const NumberValue(3), const NumberValue(4)],
        [const NumberValue(5), const NumberValue(6)],
      ]);
    });

    test('drop rows from end (negative)', () {
      context.rangeMap['A1:B3'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
        [NumberValue(5), NumberValue(6)],
      ]);
      final result = eval(registry.get('DROP')!, [
        RangeRefNode(A1Reference.parse('A1:B3')),
        const NumberNode(-1),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(2)],
        [const NumberValue(3), const NumberValue(4)],
      ]);
    });

    test('drop rows and columns', () {
      context.rangeMap['A1:C3'] = const RangeValue([
        [NumberValue(1), NumberValue(2), NumberValue(3)],
        [NumberValue(4), NumberValue(5), NumberValue(6)],
        [NumberValue(7), NumberValue(8), NumberValue(9)],
      ]);
      final result = eval(registry.get('DROP')!, [
        RangeRefNode(A1Reference.parse('A1:C3')),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      expectRange(result, [
        [const NumberValue(5), const NumberValue(6)],
        [const NumberValue(8), const NumberValue(9)],
      ]);
    });

    test('drop all rows returns #CALC!', () {
      context.rangeMap['A1:A2'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      final result = eval(registry.get('DROP')!, [
        RangeRefNode(A1Reference.parse('A1:A2')),
        const NumberNode(3),
      ]);
      expect(result, const ErrorValue(FormulaError.calc));
    });

    test('drop all columns returns #CALC!', () {
      context.rangeMap['A1:B1'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
      ]);
      final result = eval(registry.get('DROP')!, [
        RangeRefNode(A1Reference.parse('A1:B1')),
        const NumberNode(0),
        const NumberNode(3),
      ]);
      expect(result, const ErrorValue(FormulaError.calc));
    });

    test('drop 0 rows returns original', () {
      context.rangeMap['A1:A2'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      final result = eval(registry.get('DROP')!, [
        RangeRefNode(A1Reference.parse('A1:A2')),
        const NumberNode(0),
      ]);
      expectRange(result, [
        [const NumberValue(1)],
        [const NumberValue(2)],
      ]);
    });
  });

  group('TAKE', () {
    test('take rows from start', () {
      context.rangeMap['A1:B3'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
        [NumberValue(5), NumberValue(6)],
      ]);
      final result = eval(registry.get('TAKE')!, [
        RangeRefNode(A1Reference.parse('A1:B3')),
        const NumberNode(2),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(2)],
        [const NumberValue(3), const NumberValue(4)],
      ]);
    });

    test('take rows from end (negative)', () {
      context.rangeMap['A1:B3'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
        [NumberValue(5), NumberValue(6)],
      ]);
      final result = eval(registry.get('TAKE')!, [
        RangeRefNode(A1Reference.parse('A1:B3')),
        const NumberNode(-1),
      ]);
      expectRange(result, [
        [const NumberValue(5), const NumberValue(6)],
      ]);
    });

    test('take rows and columns', () {
      context.rangeMap['A1:C3'] = const RangeValue([
        [NumberValue(1), NumberValue(2), NumberValue(3)],
        [NumberValue(4), NumberValue(5), NumberValue(6)],
        [NumberValue(7), NumberValue(8), NumberValue(9)],
      ]);
      final result = eval(registry.get('TAKE')!, [
        RangeRefNode(A1Reference.parse('A1:C3')),
        const NumberNode(2),
        const NumberNode(2),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(2)],
        [const NumberValue(4), const NumberValue(5)],
      ]);
    });

    test('take negative rows and columns', () {
      context.rangeMap['A1:C3'] = const RangeValue([
        [NumberValue(1), NumberValue(2), NumberValue(3)],
        [NumberValue(4), NumberValue(5), NumberValue(6)],
        [NumberValue(7), NumberValue(8), NumberValue(9)],
      ]);
      final result = eval(registry.get('TAKE')!, [
        RangeRefNode(A1Reference.parse('A1:C3')),
        const NumberNode(-2),
        const NumberNode(-1),
      ]);
      expectRange(result, [
        [const NumberValue(6)],
        [const NumberValue(9)],
      ]);
    });

    test('rows=0 returns #CALC!', () {
      context.rangeMap['A1:A2'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      final result = eval(registry.get('TAKE')!, [
        RangeRefNode(A1Reference.parse('A1:A2')),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.calc));
    });

    test('cols=0 takes all columns', () {
      context.rangeMap['A1:B2'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
      ]);
      final result = eval(registry.get('TAKE')!, [
        RangeRefNode(A1Reference.parse('A1:B2')),
        const NumberNode(1),
        const NumberNode(0),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(2)],
      ]);
    });
  });

  group('EXPAND', () {
    test('expand rows', () {
      context.rangeMap['A1:B1'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
      ]);
      final result = eval(registry.get('EXPAND')!, [
        RangeRefNode(A1Reference.parse('A1:B1')),
        const NumberNode(3),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(2)],
        [
          const ErrorValue(FormulaError.na),
          const ErrorValue(FormulaError.na),
        ],
        [
          const ErrorValue(FormulaError.na),
          const ErrorValue(FormulaError.na),
        ],
      ]);
    });

    test('expand rows and columns', () {
      context.rangeMap['A1:A1'] = const RangeValue([
        [NumberValue(1)],
      ]);
      final result = eval(registry.get('EXPAND')!, [
        RangeRefNode(A1Reference.parse('A1:A1')),
        const NumberNode(2),
        const NumberNode(2),
      ]);
      expectRange(result, [
        [const NumberValue(1), const ErrorValue(FormulaError.na)],
        [const ErrorValue(FormulaError.na), const ErrorValue(FormulaError.na)],
      ]);
    });

    test('custom pad value', () {
      context.rangeMap['A1:A1'] = const RangeValue([
        [NumberValue(1)],
      ]);
      final result = eval(registry.get('EXPAND')!, [
        RangeRefNode(A1Reference.parse('A1:A1')),
        const NumberNode(2),
        const NumberNode(2),
        const NumberNode(0),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(0)],
        [const NumberValue(0), const NumberValue(0)],
      ]);
    });

    test('cannot shrink rows returns #VALUE!', () {
      context.rangeMap['A1:A3'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]);
      final result = eval(registry.get('EXPAND')!, [
        RangeRefNode(A1Reference.parse('A1:A3')),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('cannot shrink columns returns #VALUE!', () {
      context.rangeMap['A1:C1'] = const RangeValue([
        [NumberValue(1), NumberValue(2), NumberValue(3)],
      ]);
      final result = eval(registry.get('EXPAND')!, [
        RangeRefNode(A1Reference.parse('A1:C1')),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('same size returns same array', () {
      context.rangeMap['A1:B2'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
      ]);
      final result = eval(registry.get('EXPAND')!, [
        RangeRefNode(A1Reference.parse('A1:B2')),
        const NumberNode(2),
        const NumberNode(2),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(2)],
        [const NumberValue(3), const NumberValue(4)],
      ]);
    });
  });

  // ─── Wave D — Concatenation ───────────────────────────────

  group('HSTACK', () {
    test('stacks two arrays horizontally', () {
      context.rangeMap['A1:A2'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      context.rangeMap['B1:B2'] = const RangeValue([
        [NumberValue(3)],
        [NumberValue(4)],
      ]);
      final result = eval(registry.get('HSTACK')!, [
        RangeRefNode(A1Reference.parse('A1:A2')),
        RangeRefNode(A1Reference.parse('B1:B2')),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(3)],
        [const NumberValue(2), const NumberValue(4)],
      ]);
    });

    test('pads shorter arrays with #N/A', () {
      context.rangeMap['A1:A3'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]);
      context.rangeMap['B1:B1'] = const RangeValue([
        [NumberValue(10)],
      ]);
      final result = eval(registry.get('HSTACK')!, [
        RangeRefNode(A1Reference.parse('A1:A3')),
        RangeRefNode(A1Reference.parse('B1:B1')),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(10)],
        [const NumberValue(2), const ErrorValue(FormulaError.na)],
        [const NumberValue(3), const ErrorValue(FormulaError.na)],
      ]);
    });

    test('scalars treated as 1x1', () {
      final result = eval(registry.get('HSTACK')!, [
        const NumberNode(1),
        const NumberNode(2),
        const NumberNode(3),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(2), const NumberValue(3)],
      ]);
    });

    test('single array returns itself', () {
      context.rangeMap['A1:B2'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
      ]);
      final result = eval(registry.get('HSTACK')!, [
        RangeRefNode(A1Reference.parse('A1:B2')),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(2)],
        [const NumberValue(3), const NumberValue(4)],
      ]);
    });
  });

  group('VSTACK', () {
    test('stacks two arrays vertically', () {
      context.rangeMap['A1:B1'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
      ]);
      context.rangeMap['A2:B2'] = const RangeValue([
        [NumberValue(3), NumberValue(4)],
      ]);
      final result = eval(registry.get('VSTACK')!, [
        RangeRefNode(A1Reference.parse('A1:B1')),
        RangeRefNode(A1Reference.parse('A2:B2')),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(2)],
        [const NumberValue(3), const NumberValue(4)],
      ]);
    });

    test('pads narrower arrays with #N/A', () {
      context.rangeMap['A1:B1'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
      ]);
      context.rangeMap['A2:A2'] = const RangeValue([
        [NumberValue(3)],
      ]);
      final result = eval(registry.get('VSTACK')!, [
        RangeRefNode(A1Reference.parse('A1:B1')),
        RangeRefNode(A1Reference.parse('A2:A2')),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(2)],
        [const NumberValue(3), const ErrorValue(FormulaError.na)],
      ]);
    });

    test('scalars treated as 1x1', () {
      final result = eval(registry.get('VSTACK')!, [
        const NumberNode(1),
        const NumberNode(2),
        const NumberNode(3),
      ]);
      expectRange(result, [
        [const NumberValue(1)],
        [const NumberValue(2)],
        [const NumberValue(3)],
      ]);
    });

    test('single array returns itself', () {
      context.rangeMap['A1:B1'] = const RangeValue([
        [NumberValue(10), NumberValue(20)],
      ]);
      final result = eval(registry.get('VSTACK')!, [
        RangeRefNode(A1Reference.parse('A1:B1')),
      ]);
      expectRange(result, [
        [const NumberValue(10), const NumberValue(20)],
      ]);
    });
  });

  // ─── Wave E — Filter/Sort/Unique ──────────────────────────

  group('FILTER', () {
    test('filters rows by boolean column', () {
      context.rangeMap['A1:B3'] = const RangeValue([
        [TextValue('Apple'), NumberValue(100)],
        [TextValue('Banana'), NumberValue(200)],
        [TextValue('Cherry'), NumberValue(50)],
      ]);
      context.rangeMap['C1:C3'] = const RangeValue([
        [BooleanValue(true)],
        [BooleanValue(false)],
        [BooleanValue(true)],
      ]);
      final result = eval(registry.get('FILTER')!, [
        RangeRefNode(A1Reference.parse('A1:B3')),
        RangeRefNode(A1Reference.parse('C1:C3')),
      ]);
      expectRange(result, [
        [const TextValue('Apple'), const NumberValue(100)],
        [const TextValue('Cherry'), const NumberValue(50)],
      ]);
    });

    test('filters rows by numeric include (non-zero = true)', () {
      context.rangeMap['A1:A3'] = const RangeValue([
        [NumberValue(10)],
        [NumberValue(20)],
        [NumberValue(30)],
      ]);
      context.rangeMap['B1:B3'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(0)],
        [NumberValue(1)],
      ]);
      final result = eval(registry.get('FILTER')!, [
        RangeRefNode(A1Reference.parse('A1:A3')),
        RangeRefNode(A1Reference.parse('B1:B3')),
      ]);
      expectRange(result, [
        [const NumberValue(10)],
        [const NumberValue(30)],
      ]);
    });

    test('no matches without if_empty returns #CALC!', () {
      context.rangeMap['A1:A2'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      context.rangeMap['B1:B2'] = const RangeValue([
        [BooleanValue(false)],
        [BooleanValue(false)],
      ]);
      final result = eval(registry.get('FILTER')!, [
        RangeRefNode(A1Reference.parse('A1:A2')),
        RangeRefNode(A1Reference.parse('B1:B2')),
      ]);
      expect(result, const ErrorValue(FormulaError.calc));
    });

    test('no matches with if_empty returns if_empty', () {
      context.rangeMap['A1:A2'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      context.rangeMap['B1:B2'] = const RangeValue([
        [BooleanValue(false)],
        [BooleanValue(false)],
      ]);
      final result = eval(registry.get('FILTER')!, [
        RangeRefNode(A1Reference.parse('A1:A2')),
        RangeRefNode(A1Reference.parse('B1:B2')),
        const TextNode('No results'),
      ]);
      expect(result, const TextValue('No results'));
    });

    test('column filter when include matches columns', () {
      context.rangeMap['A1:C1'] = const RangeValue([
        [NumberValue(10), NumberValue(20), NumberValue(30)],
      ]);
      context.rangeMap['A2:C2'] = const RangeValue([
        [BooleanValue(true), BooleanValue(false), BooleanValue(true)],
      ]);
      final result = eval(registry.get('FILTER')!, [
        RangeRefNode(A1Reference.parse('A1:C1')),
        RangeRefNode(A1Reference.parse('A2:C2')),
      ]);
      expectRange(result, [
        [const NumberValue(10), const NumberValue(30)],
      ]);
    });

    test('include dimension mismatch returns #VALUE!', () {
      context.rangeMap['A1:B2'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
      ]);
      context.rangeMap['C1:C3'] = const RangeValue([
        [BooleanValue(true)],
        [BooleanValue(false)],
        [BooleanValue(true)],
      ]);
      final result = eval(registry.get('FILTER')!, [
        RangeRefNode(A1Reference.parse('A1:B2')),
        RangeRefNode(A1Reference.parse('C1:C3')),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('UNIQUE', () {
    test('removes duplicate rows', () {
      context.rangeMap['A1:A4'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(1)],
        [NumberValue(3)],
      ]);
      final result = eval(registry.get('UNIQUE')!, [
        RangeRefNode(A1Reference.parse('A1:A4')),
      ]);
      expectRange(result, [
        [const NumberValue(1)],
        [const NumberValue(2)],
        [const NumberValue(3)],
      ]);
    });

    test('case-insensitive text comparison', () {
      context.rangeMap['A1:A3'] = const RangeValue([
        [TextValue('Apple')],
        [TextValue('apple')],
        [TextValue('Banana')],
      ]);
      final result = eval(registry.get('UNIQUE')!, [
        RangeRefNode(A1Reference.parse('A1:A3')),
      ]);
      expectRange(result, [
        [const TextValue('Apple')],
        [const TextValue('Banana')],
      ]);
    });

    test('multi-column uniqueness', () {
      context.rangeMap['A1:B4'] = const RangeValue([
        [TextValue('A'), NumberValue(1)],
        [TextValue('B'), NumberValue(2)],
        [TextValue('A'), NumberValue(1)],
        [TextValue('A'), NumberValue(2)],
      ]);
      final result = eval(registry.get('UNIQUE')!, [
        RangeRefNode(A1Reference.parse('A1:B4')),
      ]);
      expectRange(result, [
        [const TextValue('A'), const NumberValue(1)],
        [const TextValue('B'), const NumberValue(2)],
        [const TextValue('A'), const NumberValue(2)],
      ]);
    });

    test('by_col=TRUE unique columns', () {
      context.rangeMap['A1:C1'] = const RangeValue([
        [NumberValue(1), NumberValue(2), NumberValue(1)],
      ]);
      final result = eval(registry.get('UNIQUE')!, [
        RangeRefNode(A1Reference.parse('A1:C1')),
        const BooleanNode(true),
      ]);
      expectRange(result, [
        [const NumberValue(1), const NumberValue(2)],
      ]);
    });

    test('exactly_once=TRUE keeps only non-duplicated', () {
      context.rangeMap['A1:A5'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(1)],
        [NumberValue(3)],
        [NumberValue(2)],
      ]);
      final result = eval(registry.get('UNIQUE')!, [
        RangeRefNode(A1Reference.parse('A1:A5')),
        const BooleanNode(false),
        const BooleanNode(true),
      ]);
      expectRange(result, [
        [const NumberValue(3)],
      ]);
    });

    test('exactly_once with all duplicated returns #CALC!', () {
      context.rangeMap['A1:A4'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      final result = eval(registry.get('UNIQUE')!, [
        RangeRefNode(A1Reference.parse('A1:A4')),
        const BooleanNode(false),
        const BooleanNode(true),
      ]);
      expect(result, const ErrorValue(FormulaError.calc));
    });
  });

  group('SORT', () {
    test('sorts single column ascending', () {
      context.rangeMap['A1:A4'] = const RangeValue([
        [NumberValue(30)],
        [NumberValue(10)],
        [NumberValue(40)],
        [NumberValue(20)],
      ]);
      final result = eval(registry.get('SORT')!, [
        RangeRefNode(A1Reference.parse('A1:A4')),
      ]);
      expectRange(result, [
        [const NumberValue(10)],
        [const NumberValue(20)],
        [const NumberValue(30)],
        [const NumberValue(40)],
      ]);
    });

    test('sorts descending', () {
      context.rangeMap['A1:A3'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(3)],
        [NumberValue(2)],
      ]);
      final result = eval(registry.get('SORT')!, [
        RangeRefNode(A1Reference.parse('A1:A3')),
        const NumberNode(1),
        const NumberNode(-1),
      ]);
      expectRange(result, [
        [const NumberValue(3)],
        [const NumberValue(2)],
        [const NumberValue(1)],
      ]);
    });

    test('sorts by second column', () {
      context.rangeMap['A1:B3'] = const RangeValue([
        [TextValue('C'), NumberValue(3)],
        [TextValue('A'), NumberValue(1)],
        [TextValue('B'), NumberValue(2)],
      ]);
      final result = eval(registry.get('SORT')!, [
        RangeRefNode(A1Reference.parse('A1:B3')),
        const NumberNode(2),
      ]);
      expectRange(result, [
        [const TextValue('A'), const NumberValue(1)],
        [const TextValue('B'), const NumberValue(2)],
        [const TextValue('C'), const NumberValue(3)],
      ]);
    });

    test('sorts by column (by_col=TRUE)', () {
      context.rangeMap['A1:C1'] = const RangeValue([
        [NumberValue(30), NumberValue(10), NumberValue(20)],
      ]);
      final result = eval(registry.get('SORT')!, [
        RangeRefNode(A1Reference.parse('A1:C1')),
        const NumberNode(1),
        const NumberNode(1),
        const BooleanNode(true),
      ]);
      expectRange(result, [
        [const NumberValue(10), const NumberValue(20), const NumberValue(30)],
      ]);
    });

    test('mixed types: numbers < text < booleans', () {
      context.rangeMap['A1:A4'] = const RangeValue([
        [TextValue('B')],
        [BooleanValue(true)],
        [NumberValue(1)],
        [TextValue('A')],
      ]);
      final result = eval(registry.get('SORT')!, [
        RangeRefNode(A1Reference.parse('A1:A4')),
      ]);
      expectRange(result, [
        [const NumberValue(1)],
        [const TextValue('A')],
        [const TextValue('B')],
        [const BooleanValue(true)],
      ]);
    });

    test('invalid sort_index returns #VALUE!', () {
      context.rangeMap['A1:B2'] = const RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
      ]);
      final result = eval(registry.get('SORT')!, [
        RangeRefNode(A1Reference.parse('A1:B2')),
        const NumberNode(5),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('SORTBY', () {
    test('sorts by external array', () {
      context.rangeMap['A1:A3'] = const RangeValue([
        [TextValue('Apple')],
        [TextValue('Banana')],
        [TextValue('Cherry')],
      ]);
      context.rangeMap['B1:B3'] = const RangeValue([
        [NumberValue(3)],
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      final result = eval(registry.get('SORTBY')!, [
        RangeRefNode(A1Reference.parse('A1:A3')),
        RangeRefNode(A1Reference.parse('B1:B3')),
      ]);
      expectRange(result, [
        [const TextValue('Banana')],
        [const TextValue('Cherry')],
        [const TextValue('Apple')],
      ]);
    });

    test('sorts descending', () {
      context.rangeMap['A1:A3'] = const RangeValue([
        [TextValue('A')],
        [TextValue('B')],
        [TextValue('C')],
      ]);
      context.rangeMap['B1:B3'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]);
      final result = eval(registry.get('SORTBY')!, [
        RangeRefNode(A1Reference.parse('A1:A3')),
        RangeRefNode(A1Reference.parse('B1:B3')),
        const NumberNode(-1),
      ]);
      expectRange(result, [
        [const TextValue('C')],
        [const TextValue('B')],
        [const TextValue('A')],
      ]);
    });

    test('multi-key sort', () {
      context.rangeMap['A1:A4'] = const RangeValue([
        [TextValue('X')],
        [TextValue('Y')],
        [TextValue('Z')],
        [TextValue('W')],
      ]);
      context.rangeMap['B1:B4'] = const RangeValue([
        [NumberValue(2)],
        [NumberValue(1)],
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      context.rangeMap['C1:C4'] = const RangeValue([
        [TextValue('b')],
        [TextValue('a')],
        [TextValue('b')],
        [TextValue('a')],
      ]);
      final result = eval(registry.get('SORTBY')!, [
        RangeRefNode(A1Reference.parse('A1:A4')),
        RangeRefNode(A1Reference.parse('B1:B4')),
        const NumberNode(1),
        RangeRefNode(A1Reference.parse('C1:C4')),
        const NumberNode(1),
      ]);
      expectRange(result, [
        [const TextValue('Y')],
        [const TextValue('Z')],
        [const TextValue('W')],
        [const TextValue('X')],
      ]);
    });

    test('mismatched array sizes returns #VALUE!', () {
      context.rangeMap['A1:A3'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]);
      context.rangeMap['B1:B2'] = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      final result = eval(registry.get('SORTBY')!, [
        RangeRefNode(A1Reference.parse('A1:A3')),
        RangeRefNode(A1Reference.parse('B1:B2')),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('missing sort_order defaults to ascending', () {
      context.rangeMap['A1:A3'] = const RangeValue([
        [TextValue('C')],
        [TextValue('A')],
        [TextValue('B')],
      ]);
      context.rangeMap['B1:B3'] = const RangeValue([
        [NumberValue(3)],
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      final result = eval(registry.get('SORTBY')!, [
        RangeRefNode(A1Reference.parse('A1:A3')),
        RangeRefNode(A1Reference.parse('B1:B3')),
      ]);
      expectRange(result, [
        [const TextValue('A')],
        [const TextValue('B')],
        [const TextValue('C')],
      ]);
    });
  });
}
