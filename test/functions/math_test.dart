import 'dart:math' as math;

import 'package:a1/a1.dart';
import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/functions/function.dart';
import 'package:worksheet_formula/src/functions/math.dart';
import 'package:worksheet_formula/src/functions/registry.dart';
import 'package:worksheet_formula/src/functions/statistical.dart';

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
    registerStatisticalFunctions(registry);
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

  group('PI', () {
    test('returns pi', () {
      final result = eval(registry.get('PI')!, []);
      expect((result as NumberValue).value, closeTo(3.14159265, 0.00001));
    });
  });

  group('LN', () {
    test('returns natural log', () {
      final result = eval(registry.get('LN')!, [const NumberNode(math.e)]);
      expect((result as NumberValue).value, closeTo(1.0, 0.0001));
    });

    test('n <= 0 returns #NUM!', () {
      final result = eval(registry.get('LN')!, [const NumberNode(0)]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('LOG', () {
    test('default base 10', () {
      final result = eval(registry.get('LOG')!, [const NumberNode(100)]);
      expect((result as NumberValue).value, closeTo(2.0, 0.0001));
    });

    test('specified base', () {
      final result = eval(registry.get('LOG')!, [
        const NumberNode(8),
        const NumberNode(2),
      ]);
      expect((result as NumberValue).value, closeTo(3.0, 0.0001));
    });

    test('base 1 returns #DIV/0!', () {
      final result = eval(registry.get('LOG')!, [
        const NumberNode(10),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });

  group('LOG10', () {
    test('returns base-10 log', () {
      final result = eval(registry.get('LOG10')!, [const NumberNode(1000)]);
      expect((result as NumberValue).value, closeTo(3.0, 0.0001));
    });

    test('n <= 0 returns #NUM!', () {
      final result = eval(registry.get('LOG10')!, [const NumberNode(-1)]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('EXP', () {
    test('e^0 = 1', () {
      final result = eval(registry.get('EXP')!, [const NumberNode(0)]);
      expect(result, const NumberValue(1));
    });

    test('e^1 = e', () {
      final result = eval(registry.get('EXP')!, [const NumberNode(1)]);
      expect((result as NumberValue).value, closeTo(math.e, 0.0001));
    });
  });

  group('SIN', () {
    test('sin(0) = 0', () {
      final result = eval(registry.get('SIN')!, [const NumberNode(0)]);
      expect((result as NumberValue).value, closeTo(0, 0.0001));
    });

    test('sin(pi/2) = 1', () {
      final result =
          eval(registry.get('SIN')!, [NumberNode(math.pi / 2)]);
      expect((result as NumberValue).value, closeTo(1, 0.0001));
    });
  });

  group('COS', () {
    test('cos(0) = 1', () {
      final result = eval(registry.get('COS')!, [const NumberNode(0)]);
      expect((result as NumberValue).value, closeTo(1, 0.0001));
    });
  });

  group('TAN', () {
    test('tan(0) = 0', () {
      final result = eval(registry.get('TAN')!, [const NumberNode(0)]);
      expect((result as NumberValue).value, closeTo(0, 0.0001));
    });
  });

  group('ASIN', () {
    test('asin(1) = pi/2', () {
      final result = eval(registry.get('ASIN')!, [const NumberNode(1)]);
      expect((result as NumberValue).value, closeTo(math.pi / 2, 0.0001));
    });

    test('out of range returns #NUM!', () {
      final result = eval(registry.get('ASIN')!, [const NumberNode(2)]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('ACOS', () {
    test('acos(1) = 0', () {
      final result = eval(registry.get('ACOS')!, [const NumberNode(1)]);
      expect((result as NumberValue).value, closeTo(0, 0.0001));
    });

    test('out of range returns #NUM!', () {
      final result = eval(registry.get('ACOS')!, [const NumberNode(2)]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('ATAN', () {
    test('atan(1) = pi/4', () {
      final result = eval(registry.get('ATAN')!, [const NumberNode(1)]);
      expect((result as NumberValue).value, closeTo(math.pi / 4, 0.0001));
    });
  });

  group('ATAN2', () {
    test('atan2(1, 1) = pi/4', () {
      final result = eval(registry.get('ATAN2')!, [
        const NumberNode(1),
        const NumberNode(1),
      ]);
      expect((result as NumberValue).value, closeTo(math.pi / 4, 0.0001));
    });

    test('both zero returns #DIV/0!', () {
      final result = eval(registry.get('ATAN2')!, [
        const NumberNode(0),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });

  group('DEGREES', () {
    test('converts pi radians to 180 degrees', () {
      final result =
          eval(registry.get('DEGREES')!, [NumberNode(math.pi)]);
      expect((result as NumberValue).value, closeTo(180, 0.0001));
    });
  });

  group('RADIANS', () {
    test('converts 180 degrees to pi radians', () {
      final result =
          eval(registry.get('RADIANS')!, [const NumberNode(180)]);
      expect((result as NumberValue).value, closeTo(math.pi, 0.0001));
    });
  });

  group('EVEN', () {
    test('rounds up positive to even', () {
      final result = eval(registry.get('EVEN')!, [const NumberNode(1.5)]);
      expect(result, const NumberValue(2));
    });

    test('rounds down negative to even (away from zero)', () {
      final result =
          eval(registry.get('EVEN')!, [const NumberNode(-1.5)]);
      expect(result, const NumberValue(-2));
    });

    test('already even unchanged', () {
      final result = eval(registry.get('EVEN')!, [const NumberNode(4)]);
      expect(result, const NumberValue(4));
    });

    test('zero returns 0', () {
      final result = eval(registry.get('EVEN')!, [const NumberNode(0)]);
      expect(result, const NumberValue(0));
    });
  });

  group('ODD', () {
    test('rounds up positive to odd', () {
      final result = eval(registry.get('ODD')!, [const NumberNode(2)]);
      expect(result, const NumberValue(3));
    });

    test('rounds down negative to odd (away from zero)', () {
      final result = eval(registry.get('ODD')!, [const NumberNode(-2)]);
      expect(result, const NumberValue(-3));
    });

    test('already odd unchanged', () {
      final result = eval(registry.get('ODD')!, [const NumberNode(3)]);
      expect(result, const NumberValue(3));
    });

    test('zero returns 1', () {
      final result = eval(registry.get('ODD')!, [const NumberNode(0)]);
      expect(result, const NumberValue(1));
    });
  });

  group('GCD', () {
    test('gcd of two numbers', () {
      final result = eval(registry.get('GCD')!, [
        const NumberNode(12),
        const NumberNode(8),
      ]);
      expect(result, const NumberValue(4));
    });

    test('gcd of multiple numbers', () {
      final result = eval(registry.get('GCD')!, [
        const NumberNode(12),
        const NumberNode(18),
        const NumberNode(24),
      ]);
      expect(result, const NumberValue(6));
    });

    test('gcd with zero', () {
      final result = eval(registry.get('GCD')!, [
        const NumberNode(5),
        const NumberNode(0),
      ]);
      expect(result, const NumberValue(5));
    });
  });

  group('LCM', () {
    test('lcm of two numbers', () {
      final result = eval(registry.get('LCM')!, [
        const NumberNode(4),
        const NumberNode(6),
      ]);
      expect(result, const NumberValue(12));
    });

    test('lcm with zero returns 0', () {
      final result = eval(registry.get('LCM')!, [
        const NumberNode(5),
        const NumberNode(0),
      ]);
      expect(result, const NumberValue(0));
    });
  });

  group('TRUNC', () {
    test('truncates positive number', () {
      final result =
          eval(registry.get('TRUNC')!, [const NumberNode(3.7)]);
      expect(result, const NumberValue(3));
    });

    test('truncates negative number', () {
      final result =
          eval(registry.get('TRUNC')!, [const NumberNode(-3.7)]);
      expect(result, const NumberValue(-3));
    });

    test('with digits', () {
      final result = eval(registry.get('TRUNC')!, [
        const NumberNode(3.14159),
        const NumberNode(2),
      ]);
      expect(result, const NumberValue(3.14));
    });
  });

  group('MROUND', () {
    test('rounds to nearest multiple', () {
      final result = eval(registry.get('MROUND')!, [
        const NumberNode(10),
        const NumberNode(3),
      ]);
      expect(result, const NumberValue(9));
    });

    test('multiple of 0 returns 0', () {
      final result = eval(registry.get('MROUND')!, [
        const NumberNode(5),
        const NumberNode(0),
      ]);
      expect(result, const NumberValue(0));
    });

    test('different signs returns #NUM!', () {
      final result = eval(registry.get('MROUND')!, [
        const NumberNode(5),
        const NumberNode(-3),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('QUOTIENT', () {
    test('returns integer portion', () {
      final result = eval(registry.get('QUOTIENT')!, [
        const NumberNode(7),
        const NumberNode(3),
      ]);
      expect(result, const NumberValue(2));
    });

    test('negative quotient', () {
      final result = eval(registry.get('QUOTIENT')!, [
        const NumberNode(-7),
        const NumberNode(3),
      ]);
      expect(result, const NumberValue(-2));
    });

    test('divisor 0 returns #DIV/0!', () {
      final result = eval(registry.get('QUOTIENT')!, [
        const NumberNode(5),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });

  group('COMBIN', () {
    test('C(5,2) = 10', () {
      final result = eval(registry.get('COMBIN')!, [
        const NumberNode(5),
        const NumberNode(2),
      ]);
      expect(result, const NumberValue(10));
    });

    test('C(10,3) = 120', () {
      final result = eval(registry.get('COMBIN')!, [
        const NumberNode(10),
        const NumberNode(3),
      ]);
      expect(result, const NumberValue(120));
    });

    test('k > n returns #NUM!', () {
      final result = eval(registry.get('COMBIN')!, [
        const NumberNode(3),
        const NumberNode(5),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('COMBINA', () {
    test('COMBINA(4,3) = COMBIN(6,3) = 20', () {
      final result = eval(registry.get('COMBINA')!, [
        const NumberNode(4),
        const NumberNode(3),
      ]);
      expect(result, const NumberValue(20));
    });

    test('COMBINA(0,0) = 1', () {
      final result = eval(registry.get('COMBINA')!, [
        const NumberNode(0),
        const NumberNode(0),
      ]);
      expect(result, const NumberValue(1));
    });
  });

  group('FACT', () {
    test('5! = 120', () {
      final result = eval(registry.get('FACT')!, [const NumberNode(5)]);
      expect(result, const NumberValue(120));
    });

    test('0! = 1', () {
      final result = eval(registry.get('FACT')!, [const NumberNode(0)]);
      expect(result, const NumberValue(1));
    });

    test('negative returns #NUM!', () {
      final result = eval(registry.get('FACT')!, [const NumberNode(-1)]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('FACTDOUBLE', () {
    test('5!! = 15', () {
      final result =
          eval(registry.get('FACTDOUBLE')!, [const NumberNode(5)]);
      expect(result, const NumberValue(15));
    });

    test('6!! = 48', () {
      final result =
          eval(registry.get('FACTDOUBLE')!, [const NumberNode(6)]);
      expect(result, const NumberValue(48));
    });

    test('0 returns 1', () {
      final result =
          eval(registry.get('FACTDOUBLE')!, [const NumberNode(0)]);
      expect(result, const NumberValue(1));
    });

    test('-1 returns 1', () {
      final result =
          eval(registry.get('FACTDOUBLE')!, [const NumberNode(-1)]);
      expect(result, const NumberValue(1));
    });

    test('-2 returns #NUM!', () {
      final result =
          eval(registry.get('FACTDOUBLE')!, [const NumberNode(-2)]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('SUMSQ', () {
    test('sum of squares', () {
      final result = eval(registry.get('SUMSQ')!, [
        const NumberNode(3),
        const NumberNode(4),
      ]);
      expect(result, const NumberValue(25));
    });

    test('single value', () {
      final result = eval(registry.get('SUMSQ')!, [const NumberNode(5)]);
      expect(result, const NumberValue(25));
    });
  });

  group('SUBTOTAL', () {
    test('function 9 delegates to SUM', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]);
      final result = eval(registry.get('SUBTOTAL')!, [
        const NumberNode(9),
        RangeRefNode(A1Reference.parse('A1:A3')),
      ]);
      expect(result, const NumberValue(6));
    });

    test('function 1 delegates to AVERAGE', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10)],
        [NumberValue(20)],
      ]);
      final result = eval(registry.get('SUBTOTAL')!, [
        const NumberNode(1),
        RangeRefNode(A1Reference.parse('A1:A2')),
      ]);
      expect(result, const NumberValue(15));
    });

    test('function 2 delegates to COUNT', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [TextValue('x')],
        [NumberValue(3)],
      ]);
      final result = eval(registry.get('SUBTOTAL')!, [
        const NumberNode(2),
        RangeRefNode(A1Reference.parse('A1:A3')),
      ]);
      expect(result, const NumberValue(2));
    });

    test('function 101-111 works like 1-11', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]);
      final result = eval(registry.get('SUBTOTAL')!, [
        const NumberNode(109),
        RangeRefNode(A1Reference.parse('A1:A3')),
      ]);
      expect(result, const NumberValue(6));
    });

    test('invalid function_num returns #VALUE!', () {
      final result = eval(registry.get('SUBTOTAL')!, [
        const NumberNode(0),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('AGGREGATE', () {
    test('function 9 with option 0 delegates to SUM', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]);
      final result = eval(registry.get('AGGREGATE')!, [
        const NumberNode(9),
        const NumberNode(0),
        RangeRefNode(A1Reference.parse('A1:A3')),
      ]);
      expect(result, const NumberValue(6));
    });

    test('option 2 ignores errors in SUM', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [ErrorValue(FormulaError.divZero)],
        [NumberValue(3)],
      ]);
      final result = eval(registry.get('AGGREGATE')!, [
        const NumberNode(9),
        const NumberNode(2),
        RangeRefNode(A1Reference.parse('A1:A3')),
      ]);
      expect(result, const NumberValue(4));
    });

    test('option 6 ignores errors in AVERAGE', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10)],
        [ErrorValue(FormulaError.na)],
        [NumberValue(20)],
      ]);
      final result = eval(registry.get('AGGREGATE')!, [
        const NumberNode(1),
        const NumberNode(6),
        RangeRefNode(A1Reference.parse('A1:A3')),
      ]);
      expect(result, const NumberValue(15));
    });

    test('function 12 delegates to MEDIAN', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(3)],
        [NumberValue(5)],
      ]);
      final result = eval(registry.get('AGGREGATE')!, [
        const NumberNode(12),
        const NumberNode(0),
        RangeRefNode(A1Reference.parse('A1:A3')),
      ]);
      expect(result, const NumberValue(3));
    });

    test('invalid function_num returns #VALUE!', () {
      final result = eval(registry.get('AGGREGATE')!, [
        const NumberNode(20),
        const NumberNode(0),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('SERIESSUM', () {
    test('basic power series', () {
      // SERIESSUM(2, 0, 1, {1, 1, 1}) = 1*2^0 + 1*2^1 + 1*2^2 = 1+2+4 = 7
      context.rangeOverride = const RangeValue([
        [NumberValue(1), NumberValue(1), NumberValue(1)],
      ]);
      final result = eval(registry.get('SERIESSUM')!, [
        const NumberNode(2),
        const NumberNode(0),
        const NumberNode(1),
        RangeRefNode(A1Reference.parse('A1:C1')),
      ]);
      expect(result, const NumberValue(7));
    });

    test('with non-zero starting power', () {
      // SERIESSUM(3, 1, 2, {1, 1}) = 1*3^1 + 1*3^3 = 3+27 = 30
      context.rangeOverride = const RangeValue([
        [NumberValue(1), NumberValue(1)],
      ]);
      final result = eval(registry.get('SERIESSUM')!, [
        const NumberNode(3),
        const NumberNode(1),
        const NumberNode(2),
        RangeRefNode(A1Reference.parse('A1:B1')),
      ]);
      expect(result, const NumberValue(30));
    });

    test('single coefficient', () {
      // SERIESSUM(5, 2, 1, 3) = 3*5^2 = 75
      final result = eval(registry.get('SERIESSUM')!, [
        const NumberNode(5),
        const NumberNode(2),
        const NumberNode(1),
        const NumberNode(3),
      ]);
      expect(result, const NumberValue(75));
    });

    test('non-numeric x returns #VALUE!', () {
      final result = eval(registry.get('SERIESSUM')!, [
        const TextNode('x'),
        const NumberNode(0),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('SQRTPI', () {
    test('SQRTPI(1) = sqrt(pi)', () {
      final result = eval(registry.get('SQRTPI')!, [const NumberNode(1)]);
      final val = (result as NumberValue).value.toDouble();
      expect(val, closeTo(math.sqrt(math.pi), 1e-10));
    });

    test('SQRTPI(0) = 0', () {
      final result = eval(registry.get('SQRTPI')!, [const NumberNode(0)]);
      expect(result, const NumberValue(0));
    });

    test('SQRTPI(2)', () {
      final result = eval(registry.get('SQRTPI')!, [const NumberNode(2)]);
      final val = (result as NumberValue).value.toDouble();
      expect(val, closeTo(math.sqrt(2 * math.pi), 1e-10));
    });

    test('negative returns #NUM!', () {
      final result = eval(registry.get('SQRTPI')!, [const NumberNode(-1)]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('text returns #VALUE!', () {
      final result = eval(registry.get('SQRTPI')!, [const TextNode('abc')]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('MULTINOMIAL', () {
    test('MULTINOMIAL(2,3) = 5!/(2!*3!) = 10', () {
      final result = eval(registry.get('MULTINOMIAL')!, [
        const NumberNode(2),
        const NumberNode(3),
      ]);
      expect(result, const NumberValue(10));
    });

    test('MULTINOMIAL(2,3,4) = 9!/(2!*3!*4!) = 1260', () {
      final result = eval(registry.get('MULTINOMIAL')!, [
        const NumberNode(2),
        const NumberNode(3),
        const NumberNode(4),
      ]);
      expect(result, const NumberValue(1260));
    });

    test('single arg returns 1', () {
      final result =
          eval(registry.get('MULTINOMIAL')!, [const NumberNode(5)]);
      expect(result, const NumberValue(1));
    });

    test('all zeros returns 1', () {
      final result = eval(registry.get('MULTINOMIAL')!, [
        const NumberNode(0),
        const NumberNode(0),
      ]);
      expect(result, const NumberValue(1));
    });

    test('negative returns #NUM!', () {
      final result = eval(registry.get('MULTINOMIAL')!, [
        const NumberNode(2),
        const NumberNode(-1),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });
}
