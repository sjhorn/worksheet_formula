import 'package:a1/a1.dart';
import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/functions/function.dart';
import 'package:worksheet_formula/src/functions/math.dart';
import 'package:worksheet_formula/src/functions/registry.dart';

class _TestContext implements EvaluationContext {
  @override
  A1 get currentCell => 'A1'.a1;
  @override
  String? get currentSheet => null;
  @override
  bool get isCancelled => false;

  final FunctionRegistry _registry;
  FormulaValue? rangeOverride;
  _TestContext(this._registry);

  @override
  FormulaValue getCellValue(A1 cell) => const EmptyValue();
  @override
  FormulaValue getRangeValues(A1Range range) =>
      rangeOverride ?? const FormulaValue.error(FormulaError.ref);
  @override
  FormulaFunction? getFunction(String name) => _registry.get(name);
}

void main() {
  late FunctionRegistry registry;
  late _TestContext context;

  setUp(() {
    registry = FunctionRegistry(registerBuiltIns: false);
    registerMathFunctions(registry);
    context = _TestContext(registry);
  });

  FormulaValue eval(FormulaFunction fn, List<FormulaNode> args) =>
      fn.call(args, context);

  group('SUM', () {
    test('sums numbers', () {
      final result = eval(registry.get('SUM')!, [
        const NumberNode(1),
        const NumberNode(2),
        const NumberNode(3),
      ]);
      expect(result, const NumberValue(6));
    });

    test('sums range values', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10), NumberValue(20)],
        [NumberValue(30), TextValue('skip')],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:B2'));
      final result = eval(registry.get('SUM')!, [rangeNode]);
      expect(result, const NumberValue(60));
    });

    test('empty args returns 0', () {
      final result = eval(registry.get('SUM')!, []);
      expect(result, const NumberValue(0));
    });
  });

  group('AVERAGE', () {
    test('averages numbers', () {
      final result = eval(registry.get('AVERAGE')!, [
        const NumberNode(10),
        const NumberNode(20),
      ]);
      expect(result, const NumberValue(15));
    });

    test('no numbers returns #DIV/0!', () {
      final result = eval(registry.get('AVERAGE')!, [const TextNode('a')]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });

  group('MIN', () {
    test('returns minimum', () {
      final result = eval(registry.get('MIN')!, [
        const NumberNode(5),
        const NumberNode(2),
        const NumberNode(8),
      ]);
      expect(result, const NumberValue(2));
    });

    test('no numbers returns 0', () {
      final result = eval(registry.get('MIN')!, [const TextNode('a')]);
      expect(result, const NumberValue(0));
    });
  });

  group('MAX', () {
    test('returns maximum', () {
      final result = eval(registry.get('MAX')!, [
        const NumberNode(5),
        const NumberNode(2),
        const NumberNode(8),
      ]);
      expect(result, const NumberValue(8));
    });

    test('no numbers returns 0', () {
      final result = eval(registry.get('MAX')!, [const TextNode('a')]);
      expect(result, const NumberValue(0));
    });
  });

  group('ABS', () {
    test('positive number unchanged', () {
      final result = eval(registry.get('ABS')!, [const NumberNode(5)]);
      expect(result, const NumberValue(5));
    });

    test('negative number becomes positive', () {
      final result = eval(registry.get('ABS')!, [const NumberNode(-5)]);
      expect(result, const NumberValue(5));
    });

    test('non-numeric returns #VALUE!', () {
      final result = eval(registry.get('ABS')!, [const TextNode('abc')]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('ROUND', () {
    test('rounds to specified digits', () {
      final result = eval(registry.get('ROUND')!, [
        const NumberNode(3.14159),
        const NumberNode(2),
      ]);
      expect(result, const NumberValue(3.14));
    });

    test('rounds to zero digits', () {
      final result = eval(registry.get('ROUND')!, [
        const NumberNode(3.7),
        const NumberNode(0),
      ]);
      expect(result, const NumberValue(4));
    });

    test('negative digits rounds to left of decimal', () {
      final result = eval(registry.get('ROUND')!, [
        const NumberNode(1234),
        const NumberNode(-2),
      ]);
      expect(result, const NumberValue(1200));
    });

    test('non-numeric returns #VALUE!', () {
      final result = eval(registry.get('ROUND')!, [
        const TextNode('abc'),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('INT', () {
    test('floors positive number', () {
      final result = eval(registry.get('INT')!, [const NumberNode(3.7)]);
      expect(result, const NumberValue(3));
    });

    test('floors negative number', () {
      final result = eval(registry.get('INT')!, [const NumberNode(-3.2)]);
      expect(result, const NumberValue(-4));
    });

    test('non-numeric returns #VALUE!', () {
      final result = eval(registry.get('INT')!, [const TextNode('abc')]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('MOD', () {
    test('returns remainder', () {
      final result = eval(registry.get('MOD')!, [
        const NumberNode(10),
        const NumberNode(3),
      ]);
      expect(result, const NumberValue(1));
    });

    test('divide by zero returns #DIV/0!', () {
      final result = eval(registry.get('MOD')!, [
        const NumberNode(10),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });

    test('non-numeric returns #VALUE!', () {
      final result = eval(registry.get('MOD')!, [
        const TextNode('abc'),
        const NumberNode(3),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('SQRT', () {
    test('returns square root', () {
      final result = eval(registry.get('SQRT')!, [const NumberNode(9)]);
      expect((result as NumberValue).value, closeTo(3, 0.0001));
    });

    test('zero returns 0', () {
      final result = eval(registry.get('SQRT')!, [const NumberNode(0)]);
      expect(result, const NumberValue(0));
    });

    test('negative returns #NUM!', () {
      final result = eval(registry.get('SQRT')!, [const NumberNode(-1)]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('non-numeric returns #VALUE!', () {
      final result = eval(registry.get('SQRT')!, [const TextNode('abc')]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('POWER', () {
    test('raises to power', () {
      final result = eval(registry.get('POWER')!, [
        const NumberNode(2),
        const NumberNode(10),
      ]);
      expect(result, const NumberValue(1024));
    });

    test('negative exponent', () {
      final result = eval(registry.get('POWER')!, [
        const NumberNode(2),
        const NumberNode(-1),
      ]);
      expect(result, const NumberValue(0.5));
    });

    test('non-numeric returns #VALUE!', () {
      final result = eval(registry.get('POWER')!, [
        const TextNode('abc'),
        const NumberNode(2),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('SUMPRODUCT', () {
    test('multiplies and sums single array', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
      ]);
      final result = eval(registry.get('SUMPRODUCT')!, [
        RangeRefNode(A1Reference.parse('A1:B2')),
      ]);
      // Just sums: 1+2+3+4 = 10
      expect(result, const NumberValue(10));
    });

    test('non-numeric treated as 0', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(5), TextValue('x')],
      ]);
      final result = eval(registry.get('SUMPRODUCT')!, [
        RangeRefNode(A1Reference.parse('A1:B1')),
      ]);
      // 5 + 0 = 5
      expect(result, const NumberValue(5));
    });
  });

  group('ROUNDUP', () {
    test('rounds up positive number', () {
      final result = eval(registry.get('ROUNDUP')!, [
        const NumberNode(3.14159),
        const NumberNode(2),
      ]);
      expect(result, const NumberValue(3.15));
    });

    test('rounds up negative number (away from zero)', () {
      final result = eval(registry.get('ROUNDUP')!, [
        const NumberNode(-3.14159),
        const NumberNode(2),
      ]);
      expect(result, const NumberValue(-3.15));
    });

    test('non-numeric returns #VALUE!', () {
      final result = eval(registry.get('ROUNDUP')!, [
        const TextNode('abc'),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('ROUNDDOWN', () {
    test('rounds down positive number', () {
      final result = eval(registry.get('ROUNDDOWN')!, [
        const NumberNode(3.999),
        const NumberNode(2),
      ]);
      expect(result, const NumberValue(3.99));
    });

    test('rounds down negative number (toward zero)', () {
      final result = eval(registry.get('ROUNDDOWN')!, [
        const NumberNode(-3.999),
        const NumberNode(2),
      ]);
      expect(result, const NumberValue(-3.99));
    });
  });

  group('CEILING', () {
    test('rounds up to nearest multiple', () {
      final result = eval(registry.get('CEILING')!, [
        const NumberNode(2.3),
        const NumberNode(1),
      ]);
      expect(result, const NumberValue(3));
    });

    test('rounds up to nearest 0.5', () {
      final result = eval(registry.get('CEILING')!, [
        const NumberNode(2.3),
        const NumberNode(0.5),
      ]);
      expect(result, const NumberValue(2.5));
    });

    test('zero significance returns 0', () {
      final result = eval(registry.get('CEILING')!, [
        const NumberNode(5),
        const NumberNode(0),
      ]);
      expect(result, const NumberValue(0));
    });

    test('positive number with negative significance returns #NUM!', () {
      final result = eval(registry.get('CEILING')!, [
        const NumberNode(5),
        const NumberNode(-1),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('FLOOR', () {
    test('rounds down to nearest multiple', () {
      final result = eval(registry.get('FLOOR')!, [
        const NumberNode(2.7),
        const NumberNode(1),
      ]);
      expect(result, const NumberValue(2));
    });

    test('rounds down to nearest 0.5', () {
      final result = eval(registry.get('FLOOR')!, [
        const NumberNode(2.7),
        const NumberNode(0.5),
      ]);
      expect(result, const NumberValue(2.5));
    });

    test('zero significance returns #DIV/0!', () {
      final result = eval(registry.get('FLOOR')!, [
        const NumberNode(5),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });

  group('SIGN', () {
    test('positive returns 1', () {
      final result = eval(registry.get('SIGN')!, [const NumberNode(5)]);
      expect(result, const NumberValue(1));
    });

    test('negative returns -1', () {
      final result = eval(registry.get('SIGN')!, [const NumberNode(-3)]);
      expect(result, const NumberValue(-1));
    });

    test('zero returns 0', () {
      final result = eval(registry.get('SIGN')!, [const NumberNode(0)]);
      expect(result, const NumberValue(0));
    });

    test('non-numeric returns #VALUE!', () {
      final result = eval(registry.get('SIGN')!, [const TextNode('abc')]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('PRODUCT', () {
    test('multiplies numbers', () {
      final result = eval(registry.get('PRODUCT')!, [
        const NumberNode(2),
        const NumberNode(3),
        const NumberNode(4),
      ]);
      expect(result, const NumberValue(24));
    });

    test('multiplies range values', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(2), NumberValue(3)],
        [NumberValue(5), TextValue('skip')],
      ]);
      final result = eval(registry.get('PRODUCT')!, [
        RangeRefNode(A1Reference.parse('A1:B2')),
      ]);
      expect(result, const NumberValue(30));
    });
  });

  group('RAND', () {
    test('returns number between 0 and 1', () {
      final result = eval(registry.get('RAND')!, []);
      expect(result, isA<NumberValue>());
      final value = (result as NumberValue).value;
      expect(value, greaterThanOrEqualTo(0));
      expect(value, lessThan(1));
    });
  });

  group('RANDBETWEEN', () {
    test('returns integer in range', () {
      final result = eval(registry.get('RANDBETWEEN')!, [
        const NumberNode(1),
        const NumberNode(10),
      ]);
      expect(result, isA<NumberValue>());
      final value = (result as NumberValue).value;
      expect(value, greaterThanOrEqualTo(1));
      expect(value, lessThanOrEqualTo(10));
      expect(value, equals(value.toInt()));
    });

    test('bottom > top returns #NUM!', () {
      final result = eval(registry.get('RANDBETWEEN')!, [
        const NumberNode(10),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });
}
