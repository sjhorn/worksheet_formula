import 'package:a1/a1.dart';
import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/functions/function.dart';
import 'package:worksheet_formula/src/functions/registry.dart';
import 'package:worksheet_formula/src/functions/date.dart';

class _TestContext implements EvaluationContext {
  @override
  A1 get currentCell => 'A1'.a1;
  @override
  String? get currentSheet => null;
  @override
  bool get isCancelled => false;

  final FunctionRegistry _registry;
  _TestContext(this._registry);

  @override
  FormulaValue getCellValue(A1 cell) => const EmptyValue();
  @override
  FormulaValue getRangeValues(A1Range range) =>
      const FormulaValue.error(FormulaError.ref);
  @override
  FormulaFunction? getFunction(String name) => _registry.get(name);
}

void main() {
  late FunctionRegistry registry;
  late _TestContext context;

  setUp(() {
    registry = FunctionRegistry(registerBuiltIns: false);
    registerDateFunctions(registry);
    context = _TestContext(registry);
  });

  FormulaValue eval(FormulaFunction fn, List<FormulaNode> args) =>
      fn.call(args, context);

  group('DATE', () {
    test('creates serial number for known date', () {
      // January 1, 2024 = serial 45292 in Excel
      final result = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      expect(result, const NumberValue(45292));
    });

    test('January 1, 1900 returns 2', () {
      // Excel treats Jan 1, 1900 as serial 1, but due to the Lotus 1-2-3 bug
      // (which treats 1900 as a leap year), the actual difference from epoch
      // (Dec 30, 1899) is 2 days.
      final result = eval(registry.get('DATE')!, [
        const NumberNode(1900),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      expect(result, const NumberValue(2));
    });

    test('month overflow rolls to next year', () {
      // DATE(2024, 13, 1) = DATE(2025, 1, 1)
      final result = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(13),
        const NumberNode(1),
      ]);
      final expected = eval(registry.get('DATE')!, [
        const NumberNode(2025),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      expect(result, expected);
    });

    test('non-numeric args return #VALUE!', () {
      final result = eval(registry.get('DATE')!, [
        const TextNode('abc'),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('TODAY', () {
    test('returns a number', () {
      final result = eval(registry.get('TODAY')!, []);
      expect(result, isA<NumberValue>());
    });

    test('returns reasonable serial number', () {
      // We are past 2023, so serial should be > 45000
      final result = eval(registry.get('TODAY')!, []);
      expect((result as NumberValue).value, greaterThan(45000));
    });

    test('returns integer (no fractional part)', () {
      final result = eval(registry.get('TODAY')!, []);
      final value = (result as NumberValue).value;
      expect(value, equals(value.toInt()));
    });
  });

  group('NOW', () {
    test('returns a number', () {
      final result = eval(registry.get('NOW')!, []);
      expect(result, isA<NumberValue>());
    });

    test('is >= TODAY', () {
      final today = eval(registry.get('TODAY')!, []);
      final now = eval(registry.get('NOW')!, []);
      expect(
        (now as NumberValue).value,
        greaterThanOrEqualTo((today as NumberValue).value),
      );
    });
  });

  group('YEAR', () {
    test('extracts year from date serial', () {
      // DATE(2024, 3, 15) -> YEAR -> 2024
      final dateSerial = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(3),
        const NumberNode(15),
      ]);
      final result = eval(
        registry.get('YEAR')!,
        [NumberNode((dateSerial as NumberValue).value)],
      );
      expect(result, const NumberValue(2024));
    });

    test('non-numeric returns #VALUE!', () {
      final result =
          eval(registry.get('YEAR')!, [const TextNode('abc')]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('MONTH', () {
    test('extracts month from date serial', () {
      final dateSerial = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(3),
        const NumberNode(15),
      ]);
      final result = eval(
        registry.get('MONTH')!,
        [NumberNode((dateSerial as NumberValue).value)],
      );
      expect(result, const NumberValue(3));
    });
  });

  group('DAY', () {
    test('extracts day from date serial', () {
      final dateSerial = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(3),
        const NumberNode(15),
      ]);
      final result = eval(
        registry.get('DAY')!,
        [NumberNode((dateSerial as NumberValue).value)],
      );
      expect(result, const NumberValue(15));
    });
  });

  group('round-trip', () {
    test('DATE(YEAR(d), MONTH(d), DAY(d)) equals d', () {
      final original = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(7),
        const NumberNode(4),
      ]);
      final serial = (original as NumberValue).value;

      final year = eval(registry.get('YEAR')!, [NumberNode(serial)]);
      final month = eval(registry.get('MONTH')!, [NumberNode(serial)]);
      final day = eval(registry.get('DAY')!, [NumberNode(serial)]);

      final roundTrip = eval(registry.get('DATE')!, [
        NumberNode((year as NumberValue).value),
        NumberNode((month as NumberValue).value),
        NumberNode((day as NumberValue).value),
      ]);
      expect(roundTrip, original);
    });
  });
}
