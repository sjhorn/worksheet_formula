import 'package:a1/a1.dart';
import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/functions/function.dart';
import 'package:worksheet_formula/src/functions/registry.dart';
import 'package:worksheet_formula/src/functions/statistical_advanced.dart';

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
  final Map<String, FormulaValue> rangeOverrides = {};
  _TestContext(this._registry);

  @override
  FormulaValue getCellValue(A1 cell) => const EmptyValue();
  @override
  FormulaValue getRangeValues(A1Range range) {
    final key = range.toString();
    return rangeOverrides[key] ?? const FormulaValue.error(FormulaError.ref);
  }

  @override
  FormulaFunction? getFunction(String name) => _registry.get(name);

  void setRange(String ref, FormulaValue value) {
    final a1Ref = A1Reference.parse(ref);
    rangeOverrides[a1Ref.range.toString()] = value;
  }
}

RangeRefNode _range(String ref) => RangeRefNode(A1Reference.parse(ref));

void main() {
  late FunctionRegistry registry;
  late _TestContext context;

  setUp(() {
    registry = FunctionRegistry(registerBuiltIns: false);
    registerAdvancedStatisticalFunctions(registry);
    context = _TestContext(registry);
  });

  FormulaValue eval(FormulaFunction fn, List<FormulaNode> args) =>
      fn.call(args, context);

  // ═══════════════════════════════════════════════════════════════════
  // Wave 1 — Scalar Statistical Functions
  // ═══════════════════════════════════════════════════════════════════

  group('FISHER', () {
    test('basic Fisher transform', () {
      final result = eval(registry.get('FISHER')!, [
        const NumberNode(0.5),
      ]);
      expect((result as NumberValue).value, closeTo(0.5493, 0.001));
    });

    test('zero returns zero', () {
      final result = eval(registry.get('FISHER')!, [
        const NumberNode(0),
      ]);
      expect((result as NumberValue).value, closeTo(0, 0.001));
    });

    test('x <= -1 returns #NUM!', () {
      final result = eval(registry.get('FISHER')!, [
        const NumberNode(-1),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('x >= 1 returns #NUM!', () {
      final result = eval(registry.get('FISHER')!, [
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('FISHERINV', () {
    test('inverse of Fisher transform', () {
      final result = eval(registry.get('FISHERINV')!, [
        const NumberNode(0.5493),
      ]);
      expect((result as NumberValue).value, closeTo(0.5, 0.001));
    });

    test('zero returns zero', () {
      final result = eval(registry.get('FISHERINV')!, [
        const NumberNode(0),
      ]);
      expect((result as NumberValue).value, closeTo(0, 0.001));
    });

    test('roundtrip with FISHER', () {
      final fisher = eval(registry.get('FISHER')!, [
        const NumberNode(0.75),
      ]);
      final result = eval(registry.get('FISHERINV')!, [
        NumberNode((fisher as NumberValue).value),
      ]);
      expect((result as NumberValue).value, closeTo(0.75, 0.001));
    });
  });

  group('STANDARDIZE', () {
    test('basic standardization', () {
      final result = eval(registry.get('STANDARDIZE')!, [
        const NumberNode(42),
        const NumberNode(40),
        const NumberNode(1.5),
      ]);
      expect((result as NumberValue).value, closeTo(1.3333, 0.001));
    });

    test('mean value returns 0', () {
      final result = eval(registry.get('STANDARDIZE')!, [
        const NumberNode(40),
        const NumberNode(40),
        const NumberNode(1.5),
      ]);
      expect((result as NumberValue).value, closeTo(0, 0.001));
    });

    test('sigma <= 0 returns #NUM!', () {
      final result = eval(registry.get('STANDARDIZE')!, [
        const NumberNode(42),
        const NumberNode(40),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('PERMUT', () {
    test('basic permutation', () {
      final result = eval(registry.get('PERMUT')!, [
        const NumberNode(8),
        const NumberNode(2),
      ]);
      expect(result, const NumberValue(56));
    });

    test('n choose 0 returns 1', () {
      final result = eval(registry.get('PERMUT')!, [
        const NumberNode(8),
        const NumberNode(0),
      ]);
      expect(result, const NumberValue(1));
    });

    test('k > n returns #NUM!', () {
      final result = eval(registry.get('PERMUT')!, [
        const NumberNode(2),
        const NumberNode(8),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('negative n returns #NUM!', () {
      final result = eval(registry.get('PERMUT')!, [
        const NumberNode(-1),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('PERMUTATIONA', () {
    test('basic permutation with repetition', () {
      final result = eval(registry.get('PERMUTATIONA')!, [
        const NumberNode(4),
        const NumberNode(3),
      ]);
      expect(result, const NumberValue(64));
    });

    test('k = 0 returns 1', () {
      final result = eval(registry.get('PERMUTATIONA')!, [
        const NumberNode(4),
        const NumberNode(0),
      ]);
      expect(result, const NumberValue(1));
    });

    test('negative n returns #NUM!', () {
      final result = eval(registry.get('PERMUTATIONA')!, [
        const NumberNode(-1),
        const NumberNode(2),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('DEVSQ', () {
    test('sum of squared deviations', () {
      final result = eval(registry.get('DEVSQ')!, [
        const NumberNode(2),
        const NumberNode(4),
        const NumberNode(6),
        const NumberNode(8),
      ]);
      expect((result as NumberValue).value, closeTo(20, 0.001));
    });

    test('single value returns 0', () {
      final result = eval(registry.get('DEVSQ')!, [
        const NumberNode(5),
      ]);
      expect((result as NumberValue).value, closeTo(0, 0.001));
    });

    test('with range', () {
      context.setRange('A1:A4', const FormulaValue.range([
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(6)],
        [NumberValue(8)],
      ]));
      final result = eval(registry.get('DEVSQ')!, [_range('A1:A4')]);
      expect((result as NumberValue).value, closeTo(20, 0.001));
    });
  });

  group('KURT', () {
    test('kurtosis of dataset', () {
      context.setRange('A1:A10', const FormulaValue.range([
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(6)],
        [NumberValue(4)],
        [NumberValue(7)],
      ]));
      final result = eval(registry.get('KURT')!, [_range('A1:A10')]);
      expect((result as NumberValue).value, closeTo(-0.1518, 0.01));
    });

    test('fewer than 4 values returns #DIV/0!', () {
      final result = eval(registry.get('KURT')!, [
        const NumberNode(1),
        const NumberNode(2),
        const NumberNode(3),
      ]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });

  group('SKEW', () {
    test('sample skewness', () {
      context.setRange('A1:A10', const FormulaValue.range([
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(6)],
        [NumberValue(4)],
        [NumberValue(7)],
      ]));
      final result = eval(registry.get('SKEW')!, [_range('A1:A10')]);
      expect((result as NumberValue).value, closeTo(0.3595, 0.01));
    });

    test('fewer than 3 values returns #DIV/0!', () {
      final result = eval(registry.get('SKEW')!, [
        const NumberNode(1),
        const NumberNode(2),
      ]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });

  group('SKEW.P', () {
    test('population skewness', () {
      context.setRange('A1:A10', const FormulaValue.range([
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(6)],
        [NumberValue(4)],
        [NumberValue(7)],
      ]));
      final result = eval(registry.get('SKEW.P')!, [_range('A1:A10')]);
      expect((result as NumberValue).value, closeTo(0.3033, 0.01));
    });

    test('fewer than 3 values returns #DIV/0!', () {
      final result = eval(registry.get('SKEW.P')!, [
        const NumberNode(1),
        const NumberNode(2),
      ]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });

  group('COVARIANCE.P', () {
    test('population covariance', () {
      context.setRange('A1:A5', const FormulaValue.range([
        [NumberValue(3)],
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(6)],
      ]));
      context.setRange('B1:B5', const FormulaValue.range([
        [NumberValue(9)],
        [NumberValue(7)],
        [NumberValue(12)],
        [NumberValue(15)],
        [NumberValue(17)],
      ]));
      final result = eval(registry.get('COVARIANCE.P')!, [
        _range('A1:A5'),
        _range('B1:B5'),
      ]);
      expect((result as NumberValue).value, closeTo(5.2, 0.01));
    });

    test('unequal array sizes returns #N/A', () {
      context.setRange('A1:A3', const FormulaValue.range([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]));
      context.setRange('B1:B5', const FormulaValue.range([
        [NumberValue(9)],
        [NumberValue(7)],
        [NumberValue(12)],
        [NumberValue(15)],
        [NumberValue(17)],
      ]));
      final result = eval(registry.get('COVARIANCE.P')!, [
        _range('A1:A3'),
        _range('B1:B5'),
      ]);
      expect(result, const ErrorValue(FormulaError.na));
    });
  });

  group('COVARIANCE.S', () {
    test('sample covariance', () {
      context.setRange('A1:A5', const FormulaValue.range([
        [NumberValue(3)],
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(6)],
      ]));
      context.setRange('B1:B5', const FormulaValue.range([
        [NumberValue(9)],
        [NumberValue(7)],
        [NumberValue(12)],
        [NumberValue(15)],
        [NumberValue(17)],
      ]));
      final result = eval(registry.get('COVARIANCE.S')!, [
        _range('A1:A5'),
        _range('B1:B5'),
      ]);
      expect((result as NumberValue).value, closeTo(6.5, 0.01));
    });

    test('single pair returns #DIV/0!', () {
      context.setRange('A1:A1', const FormulaValue.range([
        [NumberValue(1)],
      ]));
      context.setRange('B1:B1', const FormulaValue.range([
        [NumberValue(2)],
      ]));
      final result = eval(registry.get('COVARIANCE.S')!, [
        _range('A1:A1'),
        _range('B1:B1'),
      ]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });

  group('CORREL', () {
    test('Pearson correlation coefficient', () {
      context.setRange('A1:A5', const FormulaValue.range([
        [NumberValue(3)],
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(6)],
      ]));
      context.setRange('B1:B5', const FormulaValue.range([
        [NumberValue(9)],
        [NumberValue(7)],
        [NumberValue(12)],
        [NumberValue(15)],
        [NumberValue(17)],
      ]));
      final result = eval(registry.get('CORREL')!, [
        _range('A1:A5'),
        _range('B1:B5'),
      ]);
      expect((result as NumberValue).value, closeTo(0.9970, 0.01));
    });

    test('perfect correlation returns 1', () {
      context.setRange('A1:A3', const FormulaValue.range([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]));
      context.setRange('B1:B3', const FormulaValue.range([
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(6)],
      ]));
      final result = eval(registry.get('CORREL')!, [
        _range('A1:A3'),
        _range('B1:B3'),
      ]);
      expect((result as NumberValue).value, closeTo(1.0, 0.001));
    });

    test('unequal sizes returns #N/A', () {
      context.setRange('A1:A3', const FormulaValue.range([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]));
      context.setRange('B1:B2', const FormulaValue.range([
        [NumberValue(2)],
        [NumberValue(4)],
      ]));
      final result = eval(registry.get('CORREL')!, [
        _range('A1:A3'),
        _range('B1:B2'),
      ]);
      expect(result, const ErrorValue(FormulaError.na));
    });
  });

  group('PEARSON', () {
    test('same as CORREL', () {
      context.setRange('A1:A5', const FormulaValue.range([
        [NumberValue(3)],
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(6)],
      ]));
      context.setRange('B1:B5', const FormulaValue.range([
        [NumberValue(9)],
        [NumberValue(7)],
        [NumberValue(12)],
        [NumberValue(15)],
        [NumberValue(17)],
      ]));
      final result = eval(registry.get('PEARSON')!, [
        _range('A1:A5'),
        _range('B1:B5'),
      ]);
      expect((result as NumberValue).value, closeTo(0.9970, 0.01));
    });
  });

  group('RSQ', () {
    test('R-squared value', () {
      context.setRange('A1:A5', const FormulaValue.range([
        [NumberValue(3)],
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(6)],
      ]));
      context.setRange('B1:B5', const FormulaValue.range([
        [NumberValue(9)],
        [NumberValue(7)],
        [NumberValue(12)],
        [NumberValue(15)],
        [NumberValue(17)],
      ]));
      final result = eval(registry.get('RSQ')!, [
        _range('A1:A5'),
        _range('B1:B5'),
      ]);
      expect((result as NumberValue).value, closeTo(0.9940, 0.01));
    });

    test('perfect fit returns 1', () {
      context.setRange('A1:A3', const FormulaValue.range([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]));
      context.setRange('B1:B3', const FormulaValue.range([
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(6)],
      ]));
      final result = eval(registry.get('RSQ')!, [
        _range('A1:A3'),
        _range('B1:B3'),
      ]);
      expect((result as NumberValue).value, closeTo(1.0, 0.001));
    });
  });

  group('SLOPE', () {
    test('basic slope of regression line', () {
      context.setRange('A1:A5', const FormulaValue.range([
        [NumberValue(9)],
        [NumberValue(7)],
        [NumberValue(12)],
        [NumberValue(15)],
        [NumberValue(17)],
      ]));
      context.setRange('B1:B5', const FormulaValue.range([
        [NumberValue(3)],
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(6)],
      ]));
      final result = eval(registry.get('SLOPE')!, [
        _range('A1:A5'),
        _range('B1:B5'),
      ]);
      expect((result as NumberValue).value, closeTo(2.6, 0.01));
    });

    test('horizontal line returns 0', () {
      context.setRange('A1:A3', const FormulaValue.range([
        [NumberValue(5)],
        [NumberValue(5)],
        [NumberValue(5)],
      ]));
      context.setRange('B1:B3', const FormulaValue.range([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]));
      final result = eval(registry.get('SLOPE')!, [
        _range('A1:A3'),
        _range('B1:B3'),
      ]);
      expect((result as NumberValue).value, closeTo(0, 0.001));
    });
  });

  group('INTERCEPT', () {
    test('basic y-intercept', () {
      context.setRange('A1:A5', const FormulaValue.range([
        [NumberValue(9)],
        [NumberValue(7)],
        [NumberValue(12)],
        [NumberValue(15)],
        [NumberValue(17)],
      ]));
      context.setRange('B1:B5', const FormulaValue.range([
        [NumberValue(3)],
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(6)],
      ]));
      final result = eval(registry.get('INTERCEPT')!, [
        _range('A1:A5'),
        _range('B1:B5'),
      ]);
      expect((result as NumberValue).value, closeTo(1.6, 0.01));
    });

    test('data through origin', () {
      context.setRange('A1:A3', const FormulaValue.range([
        [NumberValue(0)],
        [NumberValue(2)],
        [NumberValue(4)],
      ]));
      context.setRange('B1:B3', const FormulaValue.range([
        [NumberValue(0)],
        [NumberValue(1)],
        [NumberValue(2)],
      ]));
      final result = eval(registry.get('INTERCEPT')!, [
        _range('A1:A3'),
        _range('B1:B3'),
      ]);
      expect((result as NumberValue).value, closeTo(0, 0.01));
    });
  });

  group('STEYX', () {
    test('standard error of regression', () {
      context.setRange('A1:A5', const FormulaValue.range([
        [NumberValue(9)],
        [NumberValue(7)],
        [NumberValue(12)],
        [NumberValue(15)],
        [NumberValue(17)],
      ]));
      context.setRange('B1:B5', const FormulaValue.range([
        [NumberValue(3)],
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(6)],
      ]));
      final result = eval(registry.get('STEYX')!, [
        _range('A1:A5'),
        _range('B1:B5'),
      ]);
      expect((result as NumberValue).value, closeTo(0.3651, 0.01));
    });

    test('fewer than 3 data points returns #DIV/0!', () {
      context.setRange('A1:A2', const FormulaValue.range([
        [NumberValue(1)],
        [NumberValue(2)],
      ]));
      context.setRange('B1:B2', const FormulaValue.range([
        [NumberValue(3)],
        [NumberValue(4)],
      ]));
      final result = eval(registry.get('STEYX')!, [
        _range('A1:A2'),
        _range('B1:B2'),
      ]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });

  group('FORECAST.LINEAR', () {
    test('basic linear forecast', () {
      context.setRange('A1:A5', const FormulaValue.range([
        [NumberValue(9)],
        [NumberValue(7)],
        [NumberValue(12)],
        [NumberValue(15)],
        [NumberValue(17)],
      ]));
      context.setRange('B1:B5', const FormulaValue.range([
        [NumberValue(3)],
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(6)],
      ]));
      final result = eval(registry.get('FORECAST.LINEAR')!, [
        const NumberNode(7),
        _range('A1:A5'),
        _range('B1:B5'),
      ]);
      expect((result as NumberValue).value, closeTo(19.8, 0.01));
    });

    test('forecast at mean x returns mean y', () {
      context.setRange('A1:A3', const FormulaValue.range([
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(6)],
      ]));
      context.setRange('B1:B3', const FormulaValue.range([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]));
      final result = eval(registry.get('FORECAST.LINEAR')!, [
        const NumberNode(2),
        _range('A1:A3'),
        _range('B1:B3'),
      ]);
      expect((result as NumberValue).value, closeTo(4, 0.01));
    });
  });

  group('PROB', () {
    test('probability within range', () {
      context.setRange('A1:A4', const FormulaValue.range([
        [NumberValue(0)],
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]));
      context.setRange('B1:B4', const FormulaValue.range([
        [NumberValue(0.25)],
        [NumberValue(0.25)],
        [NumberValue(0.25)],
        [NumberValue(0.25)],
      ]));
      final result = eval(registry.get('PROB')!, [
        _range('A1:A4'),
        _range('B1:B4'),
        const NumberNode(1),
        const NumberNode(2),
      ]);
      expect((result as NumberValue).value, closeTo(0.5, 0.01));
    });

    test('single value (no upper bound)', () {
      context.setRange('A1:A4', const FormulaValue.range([
        [NumberValue(0)],
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]));
      context.setRange('B1:B4', const FormulaValue.range([
        [NumberValue(0.25)],
        [NumberValue(0.25)],
        [NumberValue(0.25)],
        [NumberValue(0.25)],
      ]));
      final result = eval(registry.get('PROB')!, [
        _range('A1:A4'),
        _range('B1:B4'),
        const NumberNode(2),
      ]);
      expect((result as NumberValue).value, closeTo(0.25, 0.01));
    });

    test('probabilities not summing to 1 still works', () {
      context.setRange('A1:A2', const FormulaValue.range([
        [NumberValue(0)],
        [NumberValue(1)],
      ]));
      context.setRange('B1:B2', const FormulaValue.range([
        [NumberValue(0.3)],
        [NumberValue(0.3)],
      ]));
      final result = eval(registry.get('PROB')!, [
        _range('A1:A2'),
        _range('B1:B2'),
        const NumberNode(0),
        const NumberNode(1),
      ]);
      expect((result as NumberValue).value, closeTo(0.6, 0.01));
    });
  });

  group('MODE.MULT', () {
    test('returns multiple modes', () {
      context.setRange('A1:A6', const FormulaValue.range([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(3)],
        [NumberValue(4)],
      ]));
      final result = eval(registry.get('MODE.MULT')!, [_range('A1:A6')]);
      expect(result, isA<RangeValue>());
      final range = result as RangeValue;
      final values = range.values.map((row) => (row.first as NumberValue).value).toList();
      expect(values, containsAll([2, 3]));
    });

    test('single mode returns single value', () {
      context.setRange('A1:A4', const FormulaValue.range([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]));
      final result = eval(registry.get('MODE.MULT')!, [_range('A1:A4')]);
      expect(result, isA<RangeValue>());
      final range = result as RangeValue;
      expect((range.values.first.first as NumberValue).value, equals(2));
    });

    test('all unique returns #N/A', () {
      context.setRange('A1:A3', const FormulaValue.range([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]));
      final result = eval(registry.get('MODE.MULT')!, [_range('A1:A3')]);
      expect(result, const ErrorValue(FormulaError.na));
    });
  });

  group('STDEVA', () {
    test('basic with numeric values', () {
      final result = eval(registry.get('STDEVA')!, [
        const NumberNode(1),
        const NumberNode(2),
        const NumberNode(3),
        const NumberNode(4),
        const NumberNode(5),
      ]);
      expect((result as NumberValue).value, closeTo(1.5811, 0.01));
    });

    test('includes TRUE as 1', () {
      final result = eval(registry.get('STDEVA')!, [
        const NumberNode(1),
        const BooleanNode(true),
        const NumberNode(3),
      ]);
      expect(result, isA<NumberValue>());
    });
  });

  group('STDEVPA', () {
    test('basic population stdev with all types', () {
      final result = eval(registry.get('STDEVPA')!, [
        const NumberNode(1),
        const NumberNode(2),
        const NumberNode(3),
        const NumberNode(4),
        const NumberNode(5),
      ]);
      expect((result as NumberValue).value, closeTo(1.4142, 0.01));
    });
  });

  group('VARA', () {
    test('basic sample variance with all types', () {
      final result = eval(registry.get('VARA')!, [
        const NumberNode(1),
        const NumberNode(2),
        const NumberNode(3),
        const NumberNode(4),
        const NumberNode(5),
      ]);
      expect((result as NumberValue).value, closeTo(2.5, 0.01));
    });
  });

  group('VARPA', () {
    test('basic population variance with all types', () {
      final result = eval(registry.get('VARPA')!, [
        const NumberNode(1),
        const NumberNode(2),
        const NumberNode(3),
        const NumberNode(4),
        const NumberNode(5),
      ]);
      expect((result as NumberValue).value, closeTo(2.0, 0.01));
    });
  });

  // ═══════════════════════════════════════════════════════════════════
  // Wave 2 — Math Primitives (Gamma, Gauss, Phi)
  // ═══════════════════════════════════════════════════════════════════

  group('GAMMA', () {
    test('gamma(5) = 4! = 24', () {
      final result = eval(registry.get('GAMMA')!, [
        const NumberNode(5),
      ]);
      expect((result as NumberValue).value, closeTo(24.0, 0.01));
    });

    test('gamma(0.5) = sqrt(pi)', () {
      final result = eval(registry.get('GAMMA')!, [
        const NumberNode(0.5),
      ]);
      expect((result as NumberValue).value, closeTo(1.7725, 0.01));
    });

    test('gamma(1) = 1', () {
      final result = eval(registry.get('GAMMA')!, [
        const NumberNode(1),
      ]);
      expect((result as NumberValue).value, closeTo(1.0, 0.01));
    });

    test('zero returns #NUM!', () {
      final result = eval(registry.get('GAMMA')!, [
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('negative integer returns #NUM!', () {
      final result = eval(registry.get('GAMMA')!, [
        const NumberNode(-2),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('GAMMALN', () {
    test('gammaln(5) = ln(24)', () {
      final result = eval(registry.get('GAMMALN')!, [
        const NumberNode(5),
      ]);
      expect((result as NumberValue).value, closeTo(3.1781, 0.01));
    });

    test('gammaln(1) = 0', () {
      final result = eval(registry.get('GAMMALN')!, [
        const NumberNode(1),
      ]);
      expect((result as NumberValue).value, closeTo(0, 0.01));
    });

    test('x <= 0 returns #NUM!', () {
      final result = eval(registry.get('GAMMALN')!, [
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('GAMMALN.PRECISE', () {
    test('gammaln.precise(5) = ln(24)', () {
      final result = eval(registry.get('GAMMALN.PRECISE')!, [
        const NumberNode(5),
      ]);
      expect((result as NumberValue).value, closeTo(3.1781, 0.01));
    });

    test('x <= 0 returns #NUM!', () {
      final result = eval(registry.get('GAMMALN.PRECISE')!, [
        const NumberNode(-1),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('GAUSS', () {
    test('gauss(1)', () {
      final result = eval(registry.get('GAUSS')!, [
        const NumberNode(1),
      ]);
      expect((result as NumberValue).value, closeTo(0.3413, 0.01));
    });

    test('gauss(2)', () {
      final result = eval(registry.get('GAUSS')!, [
        const NumberNode(2),
      ]);
      expect((result as NumberValue).value, closeTo(0.4772, 0.01));
    });

    test('gauss(0) = 0', () {
      final result = eval(registry.get('GAUSS')!, [
        const NumberNode(0),
      ]);
      expect((result as NumberValue).value, closeTo(0, 0.01));
    });
  });

  group('PHI', () {
    test('phi(0) standard normal PDF at 0', () {
      final result = eval(registry.get('PHI')!, [
        const NumberNode(0),
      ]);
      expect((result as NumberValue).value, closeTo(0.3989, 0.01));
    });

    test('phi(1)', () {
      final result = eval(registry.get('PHI')!, [
        const NumberNode(1),
      ]);
      expect((result as NumberValue).value, closeTo(0.2420, 0.01));
    });

    test('phi(-1) = phi(1) by symmetry', () {
      final result = eval(registry.get('PHI')!, [
        const NumberNode(-1),
      ]);
      expect((result as NumberValue).value, closeTo(0.2420, 0.01));
    });
  });

  // ═══════════════════════════════════════════════════════════════════
  // Wave 3 — Normal Distribution
  // ═══════════════════════════════════════════════════════════════════

  group('NORM.S.DIST', () {
    test('CDF at 0 = 0.5', () {
      final result = eval(registry.get('NORM.S.DIST')!, [
        const NumberNode(0),
        const BooleanNode(true),
      ]);
      expect((result as NumberValue).value, closeTo(0.5, 0.01));
    });

    test('CDF at 1.96', () {
      final result = eval(registry.get('NORM.S.DIST')!, [
        const NumberNode(1.96),
        const BooleanNode(true),
      ]);
      expect((result as NumberValue).value, closeTo(0.9750, 0.01));
    });

    test('PDF at 0', () {
      final result = eval(registry.get('NORM.S.DIST')!, [
        const NumberNode(0),
        const BooleanNode(false),
      ]);
      expect((result as NumberValue).value, closeTo(0.3989, 0.01));
    });

    test('CDF at -1.96 by symmetry', () {
      final result = eval(registry.get('NORM.S.DIST')!, [
        const NumberNode(-1.96),
        const BooleanNode(true),
      ]);
      expect((result as NumberValue).value, closeTo(0.0250, 0.01));
    });
  });

  group('NORM.S.INV', () {
    test('inverse at 0.975', () {
      final result = eval(registry.get('NORM.S.INV')!, [
        const NumberNode(0.975),
      ]);
      expect((result as NumberValue).value, closeTo(1.9600, 0.01));
    });

    test('inverse at 0.5 = 0', () {
      final result = eval(registry.get('NORM.S.INV')!, [
        const NumberNode(0.5),
      ]);
      expect((result as NumberValue).value, closeTo(0.0, 0.01));
    });

    test('probability <= 0 returns #NUM!', () {
      final result = eval(registry.get('NORM.S.INV')!, [
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('probability >= 1 returns #NUM!', () {
      final result = eval(registry.get('NORM.S.INV')!, [
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('NORM.DIST', () {
    test('CDF of general normal', () {
      final result = eval(registry.get('NORM.DIST')!, [
        const NumberNode(42),
        const NumberNode(40),
        const NumberNode(1.5),
        const BooleanNode(true),
      ]);
      expect((result as NumberValue).value, closeTo(0.9088, 0.01));
    });

    test('PDF at mean', () {
      final result = eval(registry.get('NORM.DIST')!, [
        const NumberNode(40),
        const NumberNode(40),
        const NumberNode(1.5),
        const BooleanNode(false),
      ]);
      expect(result, isA<NumberValue>());
      expect((result as NumberValue).value, greaterThan(0));
    });

    test('sigma <= 0 returns #NUM!', () {
      final result = eval(registry.get('NORM.DIST')!, [
        const NumberNode(42),
        const NumberNode(40),
        const NumberNode(0),
        const BooleanNode(true),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('NORM.INV', () {
    test('inverse of general normal CDF', () {
      final result = eval(registry.get('NORM.INV')!, [
        const NumberNode(0.9088),
        const NumberNode(40),
        const NumberNode(1.5),
      ]);
      expect((result as NumberValue).value, closeTo(42.0, 0.01));
    });

    test('inverse at 0.5 returns mean', () {
      final result = eval(registry.get('NORM.INV')!, [
        const NumberNode(0.5),
        const NumberNode(40),
        const NumberNode(1.5),
      ]);
      expect((result as NumberValue).value, closeTo(40.0, 0.01));
    });

    test('probability <= 0 returns #NUM!', () {
      final result = eval(registry.get('NORM.INV')!, [
        const NumberNode(0),
        const NumberNode(40),
        const NumberNode(1.5),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('probability >= 1 returns #NUM!', () {
      final result = eval(registry.get('NORM.INV')!, [
        const NumberNode(1),
        const NumberNode(40),
        const NumberNode(1.5),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  // ═══════════════════════════════════════════════════════════════════
  // Wave 4 — Discrete Distributions
  // ═══════════════════════════════════════════════════════════════════

  group('BINOM.DIST', () {
    test('PMF', () {
      final result = eval(registry.get('BINOM.DIST')!, [
        const NumberNode(3),
        const NumberNode(10),
        const NumberNode(0.5),
        const BooleanNode(false),
      ]);
      expect((result as NumberValue).value, closeTo(0.1172, 0.01));
    });

    test('CDF', () {
      final result = eval(registry.get('BINOM.DIST')!, [
        const NumberNode(3),
        const NumberNode(10),
        const NumberNode(0.5),
        const BooleanNode(true),
      ]);
      expect((result as NumberValue).value, closeTo(0.1719, 0.01));
    });

    test('probability 0 and k=0 returns 1 for CDF', () {
      final result = eval(registry.get('BINOM.DIST')!, [
        const NumberNode(0),
        const NumberNode(10),
        const NumberNode(0),
        const BooleanNode(true),
      ]);
      expect((result as NumberValue).value, closeTo(1.0, 0.01));
    });

    test('p < 0 returns #NUM!', () {
      final result = eval(registry.get('BINOM.DIST')!, [
        const NumberNode(3),
        const NumberNode(10),
        const NumberNode(-0.1),
        const BooleanNode(false),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('p > 1 returns #NUM!', () {
      final result = eval(registry.get('BINOM.DIST')!, [
        const NumberNode(3),
        const NumberNode(10),
        const NumberNode(1.1),
        const BooleanNode(false),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('BINOM.INV', () {
    test('inverse binomial', () {
      final result = eval(registry.get('BINOM.INV')!, [
        const NumberNode(10),
        const NumberNode(0.5),
        const NumberNode(0.6),
      ]);
      expect(result, const NumberValue(5));
    });

    test('alpha 0 returns 0', () {
      final result = eval(registry.get('BINOM.INV')!, [
        const NumberNode(10),
        const NumberNode(0.5),
        const NumberNode(0),
      ]);
      expect((result as NumberValue).value, equals(0));
    });

    test('alpha < 0 returns #NUM!', () {
      final result = eval(registry.get('BINOM.INV')!, [
        const NumberNode(10),
        const NumberNode(0.5),
        const NumberNode(-0.1),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('BINOM.DIST.RANGE', () {
    test('probability over range', () {
      final result = eval(registry.get('BINOM.DIST.RANGE')!, [
        const NumberNode(10),
        const NumberNode(0.5),
        const NumberNode(3),
        const NumberNode(5),
      ]);
      expect((result as NumberValue).value, closeTo(0.5683, 0.01));
    });

    test('single value (no upper bound)', () {
      final result = eval(registry.get('BINOM.DIST.RANGE')!, [
        const NumberNode(10),
        const NumberNode(0.5),
        const NumberNode(5),
      ]);
      expect((result as NumberValue).value, closeTo(0.2461, 0.01));
    });
  });

  group('NEGBINOM.DIST', () {
    test('PMF', () {
      final result = eval(registry.get('NEGBINOM.DIST')!, [
        const NumberNode(3),
        const NumberNode(5),
        const NumberNode(0.4),
        const BooleanNode(false),
      ]);
      expect((result as NumberValue).value, closeTo(0.0774, 0.01));
    });

    test('CDF', () {
      final result = eval(registry.get('NEGBINOM.DIST')!, [
        const NumberNode(3),
        const NumberNode(5),
        const NumberNode(0.4),
        const BooleanNode(true),
      ]);
      expect(result, isA<NumberValue>());
      expect((result as NumberValue).value, greaterThan(0));
    });

    test('p <= 0 returns #NUM!', () {
      final result = eval(registry.get('NEGBINOM.DIST')!, [
        const NumberNode(3),
        const NumberNode(5),
        const NumberNode(0),
        const BooleanNode(false),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('p > 1 returns #NUM!', () {
      final result = eval(registry.get('NEGBINOM.DIST')!, [
        const NumberNode(3),
        const NumberNode(5),
        const NumberNode(1.1),
        const BooleanNode(false),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('HYPGEOM.DIST', () {
    test('PMF', () {
      final result = eval(registry.get('HYPGEOM.DIST')!, [
        const NumberNode(3),
        const NumberNode(10),
        const NumberNode(20),
        const NumberNode(50),
        const BooleanNode(false),
      ]);
      expect((result as NumberValue).value, closeTo(0.2259, 0.01));
    });

    test('CDF', () {
      final result = eval(registry.get('HYPGEOM.DIST')!, [
        const NumberNode(3),
        const NumberNode(10),
        const NumberNode(20),
        const NumberNode(50),
        const BooleanNode(true),
      ]);
      expect(result, isA<NumberValue>());
      expect((result as NumberValue).value, greaterThan(0));
    });

    test('sample > population returns #NUM!', () {
      final result = eval(registry.get('HYPGEOM.DIST')!, [
        const NumberNode(3),
        const NumberNode(60),
        const NumberNode(20),
        const NumberNode(50),
        const BooleanNode(false),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('POISSON.DIST', () {
    test('PMF', () {
      final result = eval(registry.get('POISSON.DIST')!, [
        const NumberNode(2),
        const NumberNode(5),
        const BooleanNode(false),
      ]);
      expect((result as NumberValue).value, closeTo(0.0842, 0.01));
    });

    test('CDF', () {
      final result = eval(registry.get('POISSON.DIST')!, [
        const NumberNode(2),
        const NumberNode(5),
        const BooleanNode(true),
      ]);
      expect((result as NumberValue).value, closeTo(0.1247, 0.01));
    });

    test('x = 0 PMF', () {
      final result = eval(registry.get('POISSON.DIST')!, [
        const NumberNode(0),
        const NumberNode(5),
        const BooleanNode(false),
      ]);
      expect(result, isA<NumberValue>());
      expect((result as NumberValue).value, greaterThan(0));
    });

    test('lambda < 0 returns #NUM!', () {
      final result = eval(registry.get('POISSON.DIST')!, [
        const NumberNode(2),
        const NumberNode(-1),
        const BooleanNode(false),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('x < 0 returns #NUM!', () {
      final result = eval(registry.get('POISSON.DIST')!, [
        const NumberNode(-1),
        const NumberNode(5),
        const BooleanNode(false),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  // ═══════════════════════════════════════════════════════════════════
  // Wave 5 — Continuous Distributions
  // ═══════════════════════════════════════════════════════════════════

  group('EXPON.DIST', () {
    test('CDF', () {
      final result = eval(registry.get('EXPON.DIST')!, [
        const NumberNode(1),
        const NumberNode(1.5),
        const BooleanNode(true),
      ]);
      expect((result as NumberValue).value, closeTo(0.7769, 0.01));
    });

    test('PDF', () {
      final result = eval(registry.get('EXPON.DIST')!, [
        const NumberNode(1),
        const NumberNode(1.5),
        const BooleanNode(false),
      ]);
      expect((result as NumberValue).value, closeTo(0.3347, 0.01));
    });

    test('x = 0 CDF returns 0', () {
      final result = eval(registry.get('EXPON.DIST')!, [
        const NumberNode(0),
        const NumberNode(1.5),
        const BooleanNode(true),
      ]);
      expect((result as NumberValue).value, closeTo(0, 0.01));
    });

    test('lambda <= 0 returns #NUM!', () {
      final result = eval(registry.get('EXPON.DIST')!, [
        const NumberNode(1),
        const NumberNode(0),
        const BooleanNode(true),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('x < 0 returns #NUM!', () {
      final result = eval(registry.get('EXPON.DIST')!, [
        const NumberNode(-1),
        const NumberNode(1.5),
        const BooleanNode(true),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('GAMMA.DIST', () {
    test('CDF', () {
      final result = eval(registry.get('GAMMA.DIST')!, [
        const NumberNode(2),
        const NumberNode(3),
        const NumberNode(2),
        const BooleanNode(true),
      ]);
      expect((result as NumberValue).value, closeTo(0.0803, 0.01));
    });

    test('PDF at x > 0', () {
      final result = eval(registry.get('GAMMA.DIST')!, [
        const NumberNode(2),
        const NumberNode(3),
        const NumberNode(2),
        const BooleanNode(false),
      ]);
      expect(result, isA<NumberValue>());
      expect((result as NumberValue).value, greaterThan(0));
    });

    test('alpha <= 0 returns #NUM!', () {
      final result = eval(registry.get('GAMMA.DIST')!, [
        const NumberNode(2),
        const NumberNode(0),
        const NumberNode(2),
        const BooleanNode(true),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('beta <= 0 returns #NUM!', () {
      final result = eval(registry.get('GAMMA.DIST')!, [
        const NumberNode(2),
        const NumberNode(3),
        const NumberNode(0),
        const BooleanNode(true),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('GAMMA.INV', () {
    test('inverse gamma CDF', () {
      final result = eval(registry.get('GAMMA.INV')!, [
        const NumberNode(0.0803),
        const NumberNode(3),
        const NumberNode(2),
      ]);
      expect((result as NumberValue).value, closeTo(2.0, 0.01));
    });

    test('probability 0.5', () {
      final result = eval(registry.get('GAMMA.INV')!, [
        const NumberNode(0.5),
        const NumberNode(3),
        const NumberNode(2),
      ]);
      expect(result, isA<NumberValue>());
      expect((result as NumberValue).value, greaterThan(0));
    });

    test('probability <= 0 returns #NUM!', () {
      final result = eval(registry.get('GAMMA.INV')!, [
        const NumberNode(0),
        const NumberNode(3),
        const NumberNode(2),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('probability >= 1 returns #NUM!', () {
      final result = eval(registry.get('GAMMA.INV')!, [
        const NumberNode(1),
        const NumberNode(3),
        const NumberNode(2),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('BETA.DIST', () {
    test('CDF', () {
      final result = eval(registry.get('BETA.DIST')!, [
        const NumberNode(0.5),
        const NumberNode(2),
        const NumberNode(5),
        const BooleanNode(true),
      ]);
      expect((result as NumberValue).value, closeTo(0.8906, 0.01));
    });

    test('PDF', () {
      final result = eval(registry.get('BETA.DIST')!, [
        const NumberNode(0.5),
        const NumberNode(2),
        const NumberNode(5),
        const BooleanNode(false),
      ]);
      expect(result, isA<NumberValue>());
      expect((result as NumberValue).value, greaterThan(0));
    });

    test('x outside [0,1] returns #NUM!', () {
      final result = eval(registry.get('BETA.DIST')!, [
        const NumberNode(1.5),
        const NumberNode(2),
        const NumberNode(5),
        const BooleanNode(true),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('alpha <= 0 returns #NUM!', () {
      final result = eval(registry.get('BETA.DIST')!, [
        const NumberNode(0.5),
        const NumberNode(0),
        const NumberNode(5),
        const BooleanNode(true),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('BETA.INV', () {
    test('inverse beta CDF', () {
      final result = eval(registry.get('BETA.INV')!, [
        const NumberNode(0.8906),
        const NumberNode(2),
        const NumberNode(5),
      ]);
      expect((result as NumberValue).value, closeTo(0.5, 0.01));
    });

    test('probability 0.5', () {
      final result = eval(registry.get('BETA.INV')!, [
        const NumberNode(0.5),
        const NumberNode(2),
        const NumberNode(5),
      ]);
      expect(result, isA<NumberValue>());
      final v = (result as NumberValue).value;
      expect(v, greaterThan(0));
      expect(v, lessThan(1));
    });

    test('probability <= 0 returns #NUM!', () {
      final result = eval(registry.get('BETA.INV')!, [
        const NumberNode(0),
        const NumberNode(2),
        const NumberNode(5),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('probability >= 1 returns #NUM!', () {
      final result = eval(registry.get('BETA.INV')!, [
        const NumberNode(1),
        const NumberNode(2),
        const NumberNode(5),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('CHISQ.DIST', () {
    test('CDF', () {
      final result = eval(registry.get('CHISQ.DIST')!, [
        const NumberNode(3.841),
        const NumberNode(1),
        const BooleanNode(true),
      ]);
      expect((result as NumberValue).value, closeTo(0.9500, 0.01));
    });

    test('PDF', () {
      final result = eval(registry.get('CHISQ.DIST')!, [
        const NumberNode(3.841),
        const NumberNode(1),
        const BooleanNode(false),
      ]);
      expect(result, isA<NumberValue>());
      expect((result as NumberValue).value, greaterThan(0));
    });

    test('x < 0 returns #NUM!', () {
      final result = eval(registry.get('CHISQ.DIST')!, [
        const NumberNode(-1),
        const NumberNode(1),
        const BooleanNode(true),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('df < 1 returns #NUM!', () {
      final result = eval(registry.get('CHISQ.DIST')!, [
        const NumberNode(3.841),
        const NumberNode(0),
        const BooleanNode(true),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('CHISQ.INV', () {
    test('inverse chi-squared CDF', () {
      final result = eval(registry.get('CHISQ.INV')!, [
        const NumberNode(0.95),
        const NumberNode(1),
      ]);
      expect((result as NumberValue).value, closeTo(3.841, 0.01));
    });

    test('probability 0.5 df=1', () {
      final result = eval(registry.get('CHISQ.INV')!, [
        const NumberNode(0.5),
        const NumberNode(1),
      ]);
      expect(result, isA<NumberValue>());
      expect((result as NumberValue).value, greaterThan(0));
    });

    test('probability <= 0 returns #NUM!', () {
      final result = eval(registry.get('CHISQ.INV')!, [
        const NumberNode(0),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('probability >= 1 returns #NUM!', () {
      final result = eval(registry.get('CHISQ.INV')!, [
        const NumberNode(1),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('CHISQ.DIST.RT', () {
    test('right-tail chi-squared CDF', () {
      final result = eval(registry.get('CHISQ.DIST.RT')!, [
        const NumberNode(3.841),
        const NumberNode(1),
      ]);
      expect((result as NumberValue).value, closeTo(0.0500, 0.01));
    });

    test('x = 0 returns 1', () {
      final result = eval(registry.get('CHISQ.DIST.RT')!, [
        const NumberNode(0),
        const NumberNode(1),
      ]);
      expect((result as NumberValue).value, closeTo(1.0, 0.01));
    });
  });

  group('CHISQ.INV.RT', () {
    test('inverse right-tail chi-squared', () {
      final result = eval(registry.get('CHISQ.INV.RT')!, [
        const NumberNode(0.05),
        const NumberNode(1),
      ]);
      expect((result as NumberValue).value, closeTo(3.841, 0.01));
    });

    test('probability <= 0 returns #NUM!', () {
      final result = eval(registry.get('CHISQ.INV.RT')!, [
        const NumberNode(0),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('probability >= 1 returns #NUM!', () {
      final result = eval(registry.get('CHISQ.INV.RT')!, [
        const NumberNode(1),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('T.DIST', () {
    test('CDF', () {
      final result = eval(registry.get('T.DIST')!, [
        const NumberNode(2.0),
        const NumberNode(10),
        const BooleanNode(true),
      ]);
      expect((result as NumberValue).value, closeTo(0.9633, 0.01));
    });

    test('PDF', () {
      final result = eval(registry.get('T.DIST')!, [
        const NumberNode(0),
        const NumberNode(10),
        const BooleanNode(false),
      ]);
      expect(result, isA<NumberValue>());
      expect((result as NumberValue).value, greaterThan(0));
    });

    test('CDF at 0 = 0.5', () {
      final result = eval(registry.get('T.DIST')!, [
        const NumberNode(0),
        const NumberNode(10),
        const BooleanNode(true),
      ]);
      expect((result as NumberValue).value, closeTo(0.5, 0.01));
    });

    test('df < 1 returns #NUM!', () {
      final result = eval(registry.get('T.DIST')!, [
        const NumberNode(2),
        const NumberNode(0),
        const BooleanNode(true),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('T.INV', () {
    test('inverse t-distribution', () {
      final result = eval(registry.get('T.INV')!, [
        const NumberNode(0.975),
        const NumberNode(10),
      ]);
      expect((result as NumberValue).value, closeTo(2.2281, 0.01));
    });

    test('inverse at 0.5 = 0', () {
      final result = eval(registry.get('T.INV')!, [
        const NumberNode(0.5),
        const NumberNode(10),
      ]);
      expect((result as NumberValue).value, closeTo(0, 0.01));
    });

    test('probability <= 0 returns #NUM!', () {
      final result = eval(registry.get('T.INV')!, [
        const NumberNode(0),
        const NumberNode(10),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('probability >= 1 returns #NUM!', () {
      final result = eval(registry.get('T.INV')!, [
        const NumberNode(1),
        const NumberNode(10),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('T.DIST.2T', () {
    test('two-tailed t-distribution', () {
      final result = eval(registry.get('T.DIST.2T')!, [
        const NumberNode(2.228),
        const NumberNode(10),
      ]);
      expect((result as NumberValue).value, closeTo(0.0500, 0.01));
    });

    test('x = 0 returns 1', () {
      final result = eval(registry.get('T.DIST.2T')!, [
        const NumberNode(0),
        const NumberNode(10),
      ]);
      expect((result as NumberValue).value, closeTo(1.0, 0.01));
    });

    test('x < 0 returns #NUM!', () {
      final result = eval(registry.get('T.DIST.2T')!, [
        const NumberNode(-1),
        const NumberNode(10),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('T.INV.2T', () {
    test('inverse two-tailed t', () {
      final result = eval(registry.get('T.INV.2T')!, [
        const NumberNode(0.05),
        const NumberNode(10),
      ]);
      expect((result as NumberValue).value, closeTo(2.2281, 0.01));
    });

    test('probability = 1 returns 0', () {
      final result = eval(registry.get('T.INV.2T')!, [
        const NumberNode(1),
        const NumberNode(10),
      ]);
      expect((result as NumberValue).value, closeTo(0, 0.01));
    });

    test('probability <= 0 returns #NUM!', () {
      final result = eval(registry.get('T.INV.2T')!, [
        const NumberNode(0),
        const NumberNode(10),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('probability > 1 returns #NUM!', () {
      final result = eval(registry.get('T.INV.2T')!, [
        const NumberNode(1.5),
        const NumberNode(10),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('T.DIST.RT', () {
    test('right-tail t-distribution', () {
      final result = eval(registry.get('T.DIST.RT')!, [
        const NumberNode(2.0),
        const NumberNode(10),
      ]);
      expect((result as NumberValue).value, closeTo(0.0367, 0.01));
    });

    test('at 0 returns 0.5', () {
      final result = eval(registry.get('T.DIST.RT')!, [
        const NumberNode(0),
        const NumberNode(10),
      ]);
      expect((result as NumberValue).value, closeTo(0.5, 0.01));
    });

    test('df < 1 returns #NUM!', () {
      final result = eval(registry.get('T.DIST.RT')!, [
        const NumberNode(2.0),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  // ═══════════════════════════════════════════════════════════════════
  // Wave 6 — F, Weibull, Lognormal Distributions
  // ═══════════════════════════════════════════════════════════════════

  group('F.DIST', () {
    test('CDF', () {
      final result = eval(registry.get('F.DIST')!, [
        const NumberNode(3.0),
        const NumberNode(5),
        const NumberNode(10),
        const BooleanNode(true),
      ]);
      expect((result as NumberValue).value, closeTo(0.9344, 0.01));
    });

    test('PDF', () {
      final result = eval(registry.get('F.DIST')!, [
        const NumberNode(3.0),
        const NumberNode(5),
        const NumberNode(10),
        const BooleanNode(false),
      ]);
      expect(result, isA<NumberValue>());
      expect((result as NumberValue).value, greaterThan(0));
    });

    test('x < 0 returns #NUM!', () {
      final result = eval(registry.get('F.DIST')!, [
        const NumberNode(-1),
        const NumberNode(5),
        const NumberNode(10),
        const BooleanNode(true),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('df1 < 1 returns #NUM!', () {
      final result = eval(registry.get('F.DIST')!, [
        const NumberNode(3.0),
        const NumberNode(0),
        const NumberNode(10),
        const BooleanNode(true),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('df2 < 1 returns #NUM!', () {
      final result = eval(registry.get('F.DIST')!, [
        const NumberNode(3.0),
        const NumberNode(5),
        const NumberNode(0),
        const BooleanNode(true),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('F.INV', () {
    test('inverse F CDF', () {
      final result = eval(registry.get('F.INV')!, [
        const NumberNode(0.95),
        const NumberNode(5),
        const NumberNode(10),
      ]);
      expect((result as NumberValue).value, closeTo(3.3258, 0.01));
    });

    test('probability 0.5', () {
      final result = eval(registry.get('F.INV')!, [
        const NumberNode(0.5),
        const NumberNode(5),
        const NumberNode(10),
      ]);
      expect(result, isA<NumberValue>());
      expect((result as NumberValue).value, greaterThan(0));
    });

    test('probability <= 0 returns #NUM!', () {
      final result = eval(registry.get('F.INV')!, [
        const NumberNode(0),
        const NumberNode(5),
        const NumberNode(10),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('probability >= 1 returns #NUM!', () {
      final result = eval(registry.get('F.INV')!, [
        const NumberNode(1),
        const NumberNode(5),
        const NumberNode(10),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('F.DIST.RT', () {
    test('right-tail F-distribution', () {
      final result = eval(registry.get('F.DIST.RT')!, [
        const NumberNode(3.3258),
        const NumberNode(5),
        const NumberNode(10),
      ]);
      expect((result as NumberValue).value, closeTo(0.0500, 0.01));
    });

    test('x = 0 returns 1', () {
      final result = eval(registry.get('F.DIST.RT')!, [
        const NumberNode(0),
        const NumberNode(5),
        const NumberNode(10),
      ]);
      expect((result as NumberValue).value, closeTo(1.0, 0.01));
    });
  });

  group('F.INV.RT', () {
    test('inverse right-tail F', () {
      final result = eval(registry.get('F.INV.RT')!, [
        const NumberNode(0.05),
        const NumberNode(5),
        const NumberNode(10),
      ]);
      expect((result as NumberValue).value, closeTo(3.3258, 0.01));
    });

    test('probability <= 0 returns #NUM!', () {
      final result = eval(registry.get('F.INV.RT')!, [
        const NumberNode(0),
        const NumberNode(5),
        const NumberNode(10),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('probability >= 1 returns #NUM!', () {
      final result = eval(registry.get('F.INV.RT')!, [
        const NumberNode(1),
        const NumberNode(5),
        const NumberNode(10),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('WEIBULL.DIST', () {
    test('CDF', () {
      final result = eval(registry.get('WEIBULL.DIST')!, [
        const NumberNode(1),
        const NumberNode(2),
        const NumberNode(1),
        const BooleanNode(true),
      ]);
      expect((result as NumberValue).value, closeTo(0.6321, 0.01));
    });

    test('PDF', () {
      final result = eval(registry.get('WEIBULL.DIST')!, [
        const NumberNode(1),
        const NumberNode(2),
        const NumberNode(1),
        const BooleanNode(false),
      ]);
      expect(result, isA<NumberValue>());
      expect((result as NumberValue).value, greaterThan(0));
    });

    test('x < 0 returns #NUM!', () {
      final result = eval(registry.get('WEIBULL.DIST')!, [
        const NumberNode(-1),
        const NumberNode(2),
        const NumberNode(1),
        const BooleanNode(true),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('alpha <= 0 returns #NUM!', () {
      final result = eval(registry.get('WEIBULL.DIST')!, [
        const NumberNode(1),
        const NumberNode(0),
        const NumberNode(1),
        const BooleanNode(true),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('beta <= 0 returns #NUM!', () {
      final result = eval(registry.get('WEIBULL.DIST')!, [
        const NumberNode(1),
        const NumberNode(2),
        const NumberNode(0),
        const BooleanNode(true),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('LOGNORM.DIST', () {
    test('CDF', () {
      final result = eval(registry.get('LOGNORM.DIST')!, [
        const NumberNode(4),
        const NumberNode(3.5),
        const NumberNode(1.2),
        const BooleanNode(true),
      ]);
      expect((result as NumberValue).value, closeTo(0.0390, 0.01));
    });

    test('PDF', () {
      final result = eval(registry.get('LOGNORM.DIST')!, [
        const NumberNode(4),
        const NumberNode(3.5),
        const NumberNode(1.2),
        const BooleanNode(false),
      ]);
      expect(result, isA<NumberValue>());
      expect((result as NumberValue).value, greaterThan(0));
    });

    test('x <= 0 returns #NUM!', () {
      final result = eval(registry.get('LOGNORM.DIST')!, [
        const NumberNode(0),
        const NumberNode(3.5),
        const NumberNode(1.2),
        const BooleanNode(true),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('sigma <= 0 returns #NUM!', () {
      final result = eval(registry.get('LOGNORM.DIST')!, [
        const NumberNode(4),
        const NumberNode(3.5),
        const NumberNode(0),
        const BooleanNode(true),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('LOGNORM.INV', () {
    test('inverse lognormal CDF', () {
      final result = eval(registry.get('LOGNORM.INV')!, [
        const NumberNode(0.0390),
        const NumberNode(3.5),
        const NumberNode(1.2),
      ]);
      expect((result as NumberValue).value, closeTo(4.0, 0.01));
    });

    test('probability 0.5', () {
      final result = eval(registry.get('LOGNORM.INV')!, [
        const NumberNode(0.5),
        const NumberNode(3.5),
        const NumberNode(1.2),
      ]);
      expect(result, isA<NumberValue>());
      expect((result as NumberValue).value, greaterThan(0));
    });

    test('probability <= 0 returns #NUM!', () {
      final result = eval(registry.get('LOGNORM.INV')!, [
        const NumberNode(0),
        const NumberNode(3.5),
        const NumberNode(1.2),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('probability >= 1 returns #NUM!', () {
      final result = eval(registry.get('LOGNORM.INV')!, [
        const NumberNode(1),
        const NumberNode(3.5),
        const NumberNode(1.2),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  // ═══════════════════════════════════════════════════════════════════
  // Wave 7 — Confidence, Tests, Regression
  // ═══════════════════════════════════════════════════════════════════

  group('CONFIDENCE.NORM', () {
    test('normal confidence interval half-width', () {
      final result = eval(registry.get('CONFIDENCE.NORM')!, [
        const NumberNode(0.05),
        const NumberNode(2.5),
        const NumberNode(50),
      ]);
      expect((result as NumberValue).value, closeTo(0.6930, 0.01));
    });

    test('alpha = 0.01', () {
      final result = eval(registry.get('CONFIDENCE.NORM')!, [
        const NumberNode(0.01),
        const NumberNode(2.5),
        const NumberNode(50),
      ]);
      expect(result, isA<NumberValue>());
      expect((result as NumberValue).value, greaterThan(0));
    });

    test('alpha <= 0 returns #NUM!', () {
      final result = eval(registry.get('CONFIDENCE.NORM')!, [
        const NumberNode(0),
        const NumberNode(2.5),
        const NumberNode(50),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('alpha >= 1 returns #NUM!', () {
      final result = eval(registry.get('CONFIDENCE.NORM')!, [
        const NumberNode(1),
        const NumberNode(2.5),
        const NumberNode(50),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('sigma <= 0 returns #NUM!', () {
      final result = eval(registry.get('CONFIDENCE.NORM')!, [
        const NumberNode(0.05),
        const NumberNode(0),
        const NumberNode(50),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('n < 1 returns #NUM!', () {
      final result = eval(registry.get('CONFIDENCE.NORM')!, [
        const NumberNode(0.05),
        const NumberNode(2.5),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('CONFIDENCE.T', () {
    test('t confidence interval half-width', () {
      final result = eval(registry.get('CONFIDENCE.T')!, [
        const NumberNode(0.05),
        const NumberNode(2.5),
        const NumberNode(50),
      ]);
      expect((result as NumberValue).value, closeTo(0.7091, 0.01));
    });

    test('larger than CONFIDENCE.NORM', () {
      final normResult = eval(registry.get('CONFIDENCE.NORM')!, [
        const NumberNode(0.05),
        const NumberNode(2.5),
        const NumberNode(50),
      ]);
      final tResult = eval(registry.get('CONFIDENCE.T')!, [
        const NumberNode(0.05),
        const NumberNode(2.5),
        const NumberNode(50),
      ]);
      expect(
        (tResult as NumberValue).value,
        greaterThan((normResult as NumberValue).value),
      );
    });

    test('alpha <= 0 returns #NUM!', () {
      final result = eval(registry.get('CONFIDENCE.T')!, [
        const NumberNode(0),
        const NumberNode(2.5),
        const NumberNode(50),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('alpha >= 1 returns #NUM!', () {
      final result = eval(registry.get('CONFIDENCE.T')!, [
        const NumberNode(1),
        const NumberNode(2.5),
        const NumberNode(50),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('n = 1 returns #DIV/0!', () {
      final result = eval(registry.get('CONFIDENCE.T')!, [
        const NumberNode(0.05),
        const NumberNode(2.5),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });

  group('Z.TEST', () {
    test('one-sided p-value', () {
      context.setRange('A1:A5', const FormulaValue.range([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(5)],
      ]));
      final result = eval(registry.get('Z.TEST')!, [
        _range('A1:A5'),
        const NumberNode(4),
      ]);
      expect(result, isA<NumberValue>());
      final val = (result as NumberValue).value;
      expect(val, greaterThanOrEqualTo(0));
      expect(val, lessThanOrEqualTo(1));
    });

    test('test at mean returns 0.5', () {
      context.setRange('A1:A5', const FormulaValue.range([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(5)],
      ]));
      final result = eval(registry.get('Z.TEST')!, [
        _range('A1:A5'),
        const NumberNode(3),
      ]);
      expect((result as NumberValue).value, closeTo(0.5, 0.01));
    });

    test('with explicit sigma', () {
      context.setRange('A1:A5', const FormulaValue.range([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(5)],
      ]));
      final result = eval(registry.get('Z.TEST')!, [
        _range('A1:A5'),
        const NumberNode(4),
        const NumberNode(1),
      ]);
      expect(result, isA<NumberValue>());
      final val = (result as NumberValue).value;
      expect(val, greaterThanOrEqualTo(0));
      expect(val, lessThanOrEqualTo(1));
    });
  });

  group('T.TEST', () {
    test('paired two-tailed test', () {
      context.setRange('A1:A5', const FormulaValue.range([
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(8)],
        [NumberValue(9)],
      ]));
      context.setRange('B1:B5', const FormulaValue.range([
        [NumberValue(6)],
        [NumberValue(19)],
        [NumberValue(3)],
        [NumberValue(2)],
        [NumberValue(14)],
      ]));
      final result = eval(registry.get('T.TEST')!, [
        _range('A1:A5'),
        _range('B1:B5'),
        const NumberNode(2),
        const NumberNode(1),
      ]);
      expect(result, isA<NumberValue>());
      final pValue = (result as NumberValue).value;
      expect(pValue, greaterThan(0));
      expect(pValue, lessThanOrEqualTo(1));
    });

    test('two-sample equal variance two-tailed', () {
      context.setRange('A1:A5', const FormulaValue.range([
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(8)],
        [NumberValue(9)],
      ]));
      context.setRange('B1:B5', const FormulaValue.range([
        [NumberValue(6)],
        [NumberValue(19)],
        [NumberValue(3)],
        [NumberValue(2)],
        [NumberValue(14)],
      ]));
      final result = eval(registry.get('T.TEST')!, [
        _range('A1:A5'),
        _range('B1:B5'),
        const NumberNode(2),
        const NumberNode(2),
      ]);
      expect(result, isA<NumberValue>());
      final pValue = (result as NumberValue).value;
      expect(pValue, greaterThan(0));
      expect(pValue, lessThanOrEqualTo(1));
    });

    test('two-sample unequal variance', () {
      context.setRange('A1:A5', const FormulaValue.range([
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(8)],
        [NumberValue(9)],
      ]));
      context.setRange('B1:B5', const FormulaValue.range([
        [NumberValue(6)],
        [NumberValue(19)],
        [NumberValue(3)],
        [NumberValue(2)],
        [NumberValue(14)],
      ]));
      final result = eval(registry.get('T.TEST')!, [
        _range('A1:A5'),
        _range('B1:B5'),
        const NumberNode(2),
        const NumberNode(3),
      ]);
      expect(result, isA<NumberValue>());
      final pValue = (result as NumberValue).value;
      expect(pValue, greaterThan(0));
      expect(pValue, lessThanOrEqualTo(1));
    });

    test('invalid tails returns #NUM!', () {
      context.setRange('A1:A3', const FormulaValue.range([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]));
      context.setRange('B1:B3', const FormulaValue.range([
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(6)],
      ]));
      final result = eval(registry.get('T.TEST')!, [
        _range('A1:A3'),
        _range('B1:B3'),
        const NumberNode(3),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('invalid type returns #NUM!', () {
      context.setRange('A1:A3', const FormulaValue.range([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]));
      context.setRange('B1:B3', const FormulaValue.range([
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(6)],
      ]));
      final result = eval(registry.get('T.TEST')!, [
        _range('A1:A3'),
        _range('B1:B3'),
        const NumberNode(2),
        const NumberNode(4),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('LINEST', () {
    test('returns slope and intercept', () {
      context.setRange('A1:A5', const FormulaValue.range([
        [NumberValue(9)],
        [NumberValue(7)],
        [NumberValue(12)],
        [NumberValue(15)],
        [NumberValue(17)],
      ]));
      context.setRange('B1:B5', const FormulaValue.range([
        [NumberValue(3)],
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(6)],
      ]));
      final result = eval(registry.get('LINEST')!, [
        _range('A1:A5'),
        _range('B1:B5'),
      ]);
      expect(result, isA<RangeValue>());
      final range = result as RangeValue;
      expect(range.columnCount, greaterThanOrEqualTo(2));
      final slope = (range.values.first[0] as NumberValue).value;
      final intercept = (range.values.first[1] as NumberValue).value;
      expect(slope, closeTo(2.6, 0.01));
      expect(intercept, closeTo(1.6, 0.01));
    });

    test('with only ys (uses 1,2,3... as xs)', () {
      context.setRange('A1:A3', const FormulaValue.range([
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(6)],
      ]));
      final result = eval(registry.get('LINEST')!, [
        _range('A1:A3'),
      ]);
      expect(result, isA<RangeValue>());
      final range = result as RangeValue;
      final slope = (range.values.first[0] as NumberValue).value;
      expect(slope, closeTo(2, 0.01));
    });
  });

  group('TREND', () {
    test('returns predicted values', () {
      context.setRange('A1:A5', const FormulaValue.range([
        [NumberValue(9)],
        [NumberValue(7)],
        [NumberValue(12)],
        [NumberValue(15)],
        [NumberValue(17)],
      ]));
      context.setRange('B1:B5', const FormulaValue.range([
        [NumberValue(3)],
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(6)],
      ]));
      final result = eval(registry.get('TREND')!, [
        _range('A1:A5'),
        _range('B1:B5'),
      ]);
      expect(result, isA<RangeValue>());
      final range = result as RangeValue;
      expect(range.rowCount, equals(5));
    });

    test('with new x values', () {
      context.setRange('A1:A5', const FormulaValue.range([
        [NumberValue(9)],
        [NumberValue(7)],
        [NumberValue(12)],
        [NumberValue(15)],
        [NumberValue(17)],
      ]));
      context.setRange('B1:B5', const FormulaValue.range([
        [NumberValue(3)],
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(6)],
      ]));
      context.setRange('C1:C2', const FormulaValue.range([
        [NumberValue(7)],
        [NumberValue(8)],
      ]));
      final result = eval(registry.get('TREND')!, [
        _range('A1:A5'),
        _range('B1:B5'),
        _range('C1:C2'),
      ]);
      expect(result, isA<RangeValue>());
      final range = result as RangeValue;
      expect(range.rowCount, equals(2));
      // x=7 should give y = 2.6*7 + 1.6 = 19.8
      final predicted0 = (range.values[0][0] as NumberValue).value;
      expect(predicted0, closeTo(19.8, 0.01));
    });
  });

  group('GROWTH', () {
    test('returns predicted exponential values', () {
      context.setRange('A1:A4', const FormulaValue.range([
        [NumberValue(33100)],
        [NumberValue(47300)],
        [NumberValue(69000)],
        [NumberValue(102000)],
      ]));
      context.setRange('B1:B4', const FormulaValue.range([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
      ]));
      final result = eval(registry.get('GROWTH')!, [
        _range('A1:A4'),
        _range('B1:B4'),
      ]);
      expect(result, isA<RangeValue>());
      final range = result as RangeValue;
      expect(range.rowCount, equals(4));
    });

    test('with new x values for prediction', () {
      context.setRange('A1:A4', const FormulaValue.range([
        [NumberValue(33100)],
        [NumberValue(47300)],
        [NumberValue(69000)],
        [NumberValue(102000)],
      ]));
      context.setRange('B1:B4', const FormulaValue.range([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
      ]));
      context.setRange('C1:C2', const FormulaValue.range([
        [NumberValue(5)],
        [NumberValue(6)],
      ]));
      final result = eval(registry.get('GROWTH')!, [
        _range('A1:A4'),
        _range('B1:B4'),
        _range('C1:C2'),
      ]);
      expect(result, isA<RangeValue>());
      final range = result as RangeValue;
      expect(range.rowCount, equals(2));
      // Values should be increasing (exponential growth)
      final val1 = (range.values[0][0] as NumberValue).value;
      final val2 = (range.values[1][0] as NumberValue).value;
      expect(val2, greaterThan(val1));
    });

    test('with only ys (uses 1,2,3... as xs)', () {
      context.setRange('A1:A4', const FormulaValue.range([
        [NumberValue(33100)],
        [NumberValue(47300)],
        [NumberValue(69000)],
        [NumberValue(102000)],
      ]));
      final result = eval(registry.get('GROWTH')!, [
        _range('A1:A4'),
      ]);
      expect(result, isA<RangeValue>());
      final range = result as RangeValue;
      expect(range.rowCount, equals(4));
    });
  });
}
