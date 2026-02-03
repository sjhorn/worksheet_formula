import 'package:a1/a1.dart';
import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/functions/function.dart';
import 'package:worksheet_formula/src/functions/registry.dart';
import 'package:worksheet_formula/src/functions/financial.dart';

class _TestContext implements EvaluationContext {
  @override
  A1 get currentCell => 'A1'.a1;
  @override
  String? get currentSheet => null;
  @override
  bool get isCancelled => false;

  final FunctionRegistry _registry;
  /// Map from range string (e.g. 'A1:F1') to override value.
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

/// Helper to create a RangeRefNode from a string like 'A1:F1'.
RangeRefNode _range(String ref) => RangeRefNode(A1Reference.parse(ref));

void main() {
  late FunctionRegistry registry;
  late _TestContext context;

  setUp(() {
    registry = FunctionRegistry(registerBuiltIns: false);
    registerFinancialFunctions(registry);
    context = _TestContext(registry);
  });

  FormulaValue eval(FormulaFunction fn, List<FormulaNode> args) =>
      fn.call(args, context);

  // ─── Wave 1 — Simple Closed-Form ─────────────────────────────

  group('EFFECT', () {
    test('annual compounding', () {
      final result = eval(registry.get('EFFECT')!, [
        const NumberNode(0.10),
        const NumberNode(1),
      ]);
      expect((result as NumberValue).value, closeTo(0.10, 0.0001));
    });

    test('quarterly compounding', () {
      final result = eval(registry.get('EFFECT')!, [
        const NumberNode(0.10),
        const NumberNode(4),
      ]);
      expect((result as NumberValue).value, closeTo(0.10381, 0.0001));
    });

    test('monthly compounding', () {
      final result = eval(registry.get('EFFECT')!, [
        const NumberNode(0.10),
        const NumberNode(12),
      ]);
      expect((result as NumberValue).value, closeTo(0.10471, 0.0001));
    });

    test('rate <= 0 returns #NUM!', () {
      final result = eval(registry.get('EFFECT')!, [
        const NumberNode(0),
        const NumberNode(4),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('npery < 1 returns #NUM!', () {
      final result = eval(registry.get('EFFECT')!, [
        const NumberNode(0.10),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('NOMINAL', () {
    test('from annual effective rate', () {
      final result = eval(registry.get('NOMINAL')!, [
        const NumberNode(0.10381289),
        const NumberNode(4),
      ]);
      expect((result as NumberValue).value, closeTo(0.10, 0.0001));
    });

    test('from monthly effective rate', () {
      final result = eval(registry.get('NOMINAL')!, [
        const NumberNode(0.10471307),
        const NumberNode(12),
      ]);
      expect((result as NumberValue).value, closeTo(0.10, 0.0001));
    });

    test('rate <= 0 returns #NUM!', () {
      final result = eval(registry.get('NOMINAL')!, [
        const NumberNode(0),
        const NumberNode(4),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('npery < 1 returns #NUM!', () {
      final result = eval(registry.get('NOMINAL')!, [
        const NumberNode(0.10),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('PDURATION', () {
    test('basic calculation', () {
      final result = eval(registry.get('PDURATION')!, [
        const NumberNode(0.025),
        const NumberNode(2000),
        const NumberNode(2200),
      ]);
      expect((result as NumberValue).value, closeTo(3.8599, 0.01));
    });

    test('rate <= 0 returns #NUM!', () {
      final result = eval(registry.get('PDURATION')!, [
        const NumberNode(0),
        const NumberNode(2000),
        const NumberNode(2200),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('pv or fv <= 0 returns #NUM!', () {
      final result = eval(registry.get('PDURATION')!, [
        const NumberNode(0.05),
        const NumberNode(-100),
        const NumberNode(200),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('RRI', () {
    test('basic calculation', () {
      final result = eval(registry.get('RRI')!, [
        const NumberNode(96),
        const NumberNode(10000),
        const NumberNode(11000),
      ]);
      expect((result as NumberValue).value, closeTo(0.001009, 0.0001));
    });

    test('nper <= 0 returns #NUM!', () {
      final result = eval(registry.get('RRI')!, [
        const NumberNode(0),
        const NumberNode(10000),
        const NumberNode(11000),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('pv = 0 returns #NUM!', () {
      final result = eval(registry.get('RRI')!, [
        const NumberNode(10),
        const NumberNode(0),
        const NumberNode(11000),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('ISPMT', () {
    test('basic calculation', () {
      final result = eval(registry.get('ISPMT')!, [
        NumberNode(0.10 / 12),
        const NumberNode(1),
        const NumberNode(36),
        const NumberNode(8000000),
      ]);
      expect((result as NumberValue).value, closeTo(-64814.81, 0.01));
    });

    test('period at end of loan', () {
      final result = eval(registry.get('ISPMT')!, [
        NumberNode(0.10 / 12),
        const NumberNode(35),
        const NumberNode(36),
        const NumberNode(8000000),
      ]);
      expect((result as NumberValue).value, closeTo(-1851.85, 0.01));
    });
  });

  group('SLN', () {
    test('basic depreciation', () {
      final result = eval(registry.get('SLN')!, [
        const NumberNode(30000),
        const NumberNode(7500),
        const NumberNode(10),
      ]);
      expect(result, const NumberValue(2250));
    });

    test('life = 0 returns #DIV/0!', () {
      final result = eval(registry.get('SLN')!, [
        const NumberNode(30000),
        const NumberNode(7500),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });

  group('SYD', () {
    test('first year depreciation', () {
      final result = eval(registry.get('SYD')!, [
        const NumberNode(30000),
        const NumberNode(7500),
        const NumberNode(10),
        const NumberNode(1),
      ]);
      expect((result as NumberValue).value, closeTo(4090.91, 0.01));
    });

    test('last year depreciation', () {
      final result = eval(registry.get('SYD')!, [
        const NumberNode(30000),
        const NumberNode(7500),
        const NumberNode(10),
        const NumberNode(10),
      ]);
      expect((result as NumberValue).value, closeTo(409.09, 0.01));
    });

    test('life <= 0 returns #NUM!', () {
      final result = eval(registry.get('SYD')!, [
        const NumberNode(30000),
        const NumberNode(7500),
        const NumberNode(0),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('per < 1 returns #NUM!', () {
      final result = eval(registry.get('SYD')!, [
        const NumberNode(30000),
        const NumberNode(7500),
        const NumberNode(10),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('DOLLARDE', () {
    test('basic conversion', () {
      final result = eval(registry.get('DOLLARDE')!, [
        const NumberNode(1.02),
        const NumberNode(16),
      ]);
      expect((result as NumberValue).value, closeTo(1.125, 0.0001));
    });

    test('negative value', () {
      final result = eval(registry.get('DOLLARDE')!, [
        const NumberNode(-1.02),
        const NumberNode(16),
      ]);
      expect((result as NumberValue).value, closeTo(-1.125, 0.0001));
    });

    test('fraction < 1 returns #NUM!', () {
      final result = eval(registry.get('DOLLARDE')!, [
        const NumberNode(1.02),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('DOLLARFR', () {
    test('basic conversion', () {
      final result = eval(registry.get('DOLLARFR')!, [
        const NumberNode(1.125),
        const NumberNode(16),
      ]);
      expect((result as NumberValue).value, closeTo(1.02, 0.0001));
    });

    test('fraction < 1 returns #NUM!', () {
      final result = eval(registry.get('DOLLARFR')!, [
        const NumberNode(1.125),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('FVSCHEDULE', () {
    test('basic schedule', () {
      context.setRange('A1:C1', const FormulaValue.range([
        [NumberValue(0.09), NumberValue(0.11), NumberValue(0.10)]
      ]));
      final result = eval(registry.get('FVSCHEDULE')!, [
        const NumberNode(1),
        _range('A1:C1'),
      ]);
      expect((result as NumberValue).value, closeTo(1.33089, 0.0001));
    });

    test('single rate', () {
      context.setRange('A1:A1', const FormulaValue.range([
        [NumberValue(0.05)]
      ]));
      final result = eval(registry.get('FVSCHEDULE')!, [
        const NumberNode(100),
        _range('A1:A1'),
      ]);
      expect((result as NumberValue).value, closeTo(105, 0.01));
    });
  });

  group('NPV', () {
    test('basic net present value', () {
      final result = eval(registry.get('NPV')!, [
        const NumberNode(0.10),
        const NumberNode(-10000),
        const NumberNode(3000),
        const NumberNode(4200),
        const NumberNode(6800),
      ]);
      expect((result as NumberValue).value, closeTo(1188.44, 0.01));
    });

    test('with range argument', () {
      context.setRange('A1:F1', const FormulaValue.range([
        [
          NumberValue(-40000), NumberValue(8000), NumberValue(9200),
          NumberValue(10000), NumberValue(12000), NumberValue(14500),
        ]
      ]));
      final result = eval(registry.get('NPV')!, [
        const NumberNode(0.08),
        _range('A1:F1'),
      ]);
      expect((result as NumberValue).value, closeTo(1779.69, 0.01));
    });
  });

  group('TBILLEQ', () {
    test('basic T-Bill equivalent yield', () {
      final result = eval(registry.get('TBILLEQ')!, [
        const NumberNode(45382), // 2024-03-31
        const NumberNode(45444), // 2024-06-01
        const NumberNode(0.0914),
      ]);
      expect((result as NumberValue).value, closeTo(0.09415, 0.001));
    });

    test('DSM > 365 returns #NUM!', () {
      final result = eval(registry.get('TBILLEQ')!, [
        const NumberNode(45292),
        const NumberNode(45658),
        const NumberNode(0.05),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('TBILLPRICE', () {
    test('basic T-Bill price', () {
      final result = eval(registry.get('TBILLPRICE')!, [
        const NumberNode(45382),
        const NumberNode(45444),
        const NumberNode(0.0914),
      ]);
      expect((result as NumberValue).value, closeTo(98.4258, 0.01));
    });

    test('DSM > 365 returns #NUM!', () {
      final result = eval(registry.get('TBILLPRICE')!, [
        const NumberNode(45292),
        const NumberNode(45658),
        const NumberNode(0.05),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('TBILLYIELD', () {
    test('basic T-Bill yield', () {
      final result = eval(registry.get('TBILLYIELD')!, [
        const NumberNode(45382),
        const NumberNode(45444),
        const NumberNode(98.45),
      ]);
      expect((result as NumberValue).value, closeTo(0.09141, 0.001));
    });

    test('price <= 0 returns #NUM!', () {
      final result = eval(registry.get('TBILLYIELD')!, [
        const NumberNode(45382),
        const NumberNode(45444),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  // ─── Wave 2 — TVM Annuity ────────────────────────────────────

  group('PMT', () {
    test('basic loan payment', () {
      final result = eval(registry.get('PMT')!, [
        NumberNode(0.08 / 12),
        const NumberNode(10),
        const NumberNode(10000),
      ]);
      expect((result as NumberValue).value, closeTo(-1037.03, 0.01));
    });

    test('with future value', () {
      final result = eval(registry.get('PMT')!, [
        NumberNode(0.06 / 12),
        NumberNode(18 * 12),
        const NumberNode(0),
        const NumberNode(50000),
      ]);
      expect((result as NumberValue).value, closeTo(-129.08, 0.01));
    });

    test('type 1 (beginning of period)', () {
      final result = eval(registry.get('PMT')!, [
        NumberNode(0.06 / 12),
        NumberNode(18 * 12),
        const NumberNode(0),
        const NumberNode(50000),
        const NumberNode(1),
      ]);
      expect((result as NumberValue).value, closeTo(-128.43, 0.01));
    });

    test('zero rate', () {
      final result = eval(registry.get('PMT')!, [
        const NumberNode(0),
        const NumberNode(10),
        const NumberNode(1000),
      ]);
      expect((result as NumberValue).value, closeTo(-100, 0.01));
    });
  });

  group('FV', () {
    test('basic future value', () {
      final result = eval(registry.get('FV')!, [
        NumberNode(0.06 / 12),
        const NumberNode(10),
        const NumberNode(-200),
        const NumberNode(-500),
      ]);
      expect((result as NumberValue).value, closeTo(2571.18, 0.01));
    });

    test('type 1', () {
      final result = eval(registry.get('FV')!, [
        NumberNode(0.06 / 12),
        const NumberNode(10),
        const NumberNode(-200),
        const NumberNode(-500),
        const NumberNode(1),
      ]);
      expect((result as NumberValue).value, closeTo(2581.40, 0.01));
    });

    test('zero rate', () {
      final result = eval(registry.get('FV')!, [
        const NumberNode(0),
        const NumberNode(10),
        const NumberNode(-200),
        const NumberNode(-500),
      ]);
      expect((result as NumberValue).value, closeTo(2500, 0.01));
    });
  });

  group('PV', () {
    test('basic present value', () {
      final result = eval(registry.get('PV')!, [
        NumberNode(0.08 / 12),
        NumberNode(20 * 12),
        const NumberNode(500),
      ]);
      expect((result as NumberValue).value, closeTo(-59777.15, 0.01));
    });

    test('with future value', () {
      final result = eval(registry.get('PV')!, [
        const NumberNode(0.10),
        const NumberNode(5),
        const NumberNode(0),
        const NumberNode(100000),
      ]);
      expect((result as NumberValue).value, closeTo(-62092.13, 0.01));
    });

    test('zero rate', () {
      final result = eval(registry.get('PV')!, [
        const NumberNode(0),
        const NumberNode(10),
        const NumberNode(100),
      ]);
      expect((result as NumberValue).value, closeTo(-1000, 0.01));
    });
  });

  group('NPER', () {
    test('basic number of periods', () {
      final result = eval(registry.get('NPER')!, [
        NumberNode(0.12 / 12),
        const NumberNode(-100),
        const NumberNode(-1000),
        const NumberNode(10000),
      ]);
      expect((result as NumberValue).value, closeTo(60.08, 0.01));
    });

    test('zero rate', () {
      final result = eval(registry.get('NPER')!, [
        const NumberNode(0),
        const NumberNode(-100),
        const NumberNode(1000),
      ]);
      expect((result as NumberValue).value, closeTo(10, 0.01));
    });

    test('non-numeric returns #VALUE!', () {
      final result = eval(registry.get('NPER')!, [
        const TextNode('abc'),
        const NumberNode(-100),
        const NumberNode(1000),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('IPMT', () {
    test('first period interest', () {
      final result = eval(registry.get('IPMT')!, [
        NumberNode(0.10 / 12),
        const NumberNode(1),
        const NumberNode(36),
        const NumberNode(8000),
      ]);
      expect((result as NumberValue).value, closeTo(-66.67, 0.01));
    });

    test('last period interest', () {
      final result = eval(registry.get('IPMT')!, [
        NumberNode(0.10 / 12),
        const NumberNode(36),
        const NumberNode(36),
        const NumberNode(8000),
      ]);
      expect((result as NumberValue).value, closeTo(-2.13, 0.02));
    });

    test('per < 1 returns #NUM!', () {
      final result = eval(registry.get('IPMT')!, [
        const NumberNode(0.05),
        const NumberNode(0),
        const NumberNode(12),
        const NumberNode(1000),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('per > nper returns #NUM!', () {
      final result = eval(registry.get('IPMT')!, [
        const NumberNode(0.05),
        const NumberNode(13),
        const NumberNode(12),
        const NumberNode(1000),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('PPMT', () {
    test('first period principal', () {
      final result = eval(registry.get('PPMT')!, [
        NumberNode(0.10 / 12),
        const NumberNode(1),
        const NumberNode(36),
        const NumberNode(8000),
      ]);
      expect((result as NumberValue).value, closeTo(-191.47, 0.01));
    });

    test('last period principal', () {
      final result = eval(registry.get('PPMT')!, [
        NumberNode(0.10 / 12),
        const NumberNode(36),
        const NumberNode(36),
        const NumberNode(8000),
      ]);
      expect((result as NumberValue).value, closeTo(-256.00, 0.02));
    });
  });

  group('CUMIPMT', () {
    test('cumulative interest for first year', () {
      final result = eval(registry.get('CUMIPMT')!, [
        NumberNode(0.09 / 12),
        NumberNode(30 * 12),
        const NumberNode(125000),
        const NumberNode(1),
        const NumberNode(12),
        const NumberNode(0),
      ]);
      expect((result as NumberValue).value, closeTo(-11215.34, 0.01));
    });

    test('single period', () {
      final result = eval(registry.get('CUMIPMT')!, [
        NumberNode(0.09 / 12),
        NumberNode(30 * 12),
        const NumberNode(125000),
        const NumberNode(1),
        const NumberNode(1),
        const NumberNode(0),
      ]);
      expect((result as NumberValue).value, closeTo(-937.50, 0.01));
    });

    test('rate <= 0 returns #NUM!', () {
      final result = eval(registry.get('CUMIPMT')!, [
        const NumberNode(0),
        const NumberNode(12),
        const NumberNode(1000),
        const NumberNode(1),
        const NumberNode(12),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('start > end returns #NUM!', () {
      final result = eval(registry.get('CUMIPMT')!, [
        const NumberNode(0.05),
        const NumberNode(12),
        const NumberNode(1000),
        const NumberNode(5),
        const NumberNode(3),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('CUMPRINC', () {
    test('cumulative principal for first year', () {
      final result = eval(registry.get('CUMPRINC')!, [
        NumberNode(0.09 / 12),
        NumberNode(30 * 12),
        const NumberNode(125000),
        const NumberNode(1),
        const NumberNode(12),
        const NumberNode(0),
      ]);
      expect((result as NumberValue).value, closeTo(-854.00, 0.01));
    });

    test('rate <= 0 returns #NUM!', () {
      final result = eval(registry.get('CUMPRINC')!, [
        const NumberNode(0),
        const NumberNode(12),
        const NumberNode(1000),
        const NumberNode(1),
        const NumberNode(12),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  // ─── Wave 3 — Iterative Solvers ──────────────────────────────

  group('RATE', () {
    test('basic rate calculation', () {
      final result = eval(registry.get('RATE')!, [
        const NumberNode(48),
        const NumberNode(-200),
        const NumberNode(8000),
      ]);
      expect((result as NumberValue).value, closeTo(0.007701, 0.0001));
    });

    test('with future value', () {
      final result = eval(registry.get('RATE')!, [
        NumberNode(4 * 12),
        const NumberNode(-200),
        const NumberNode(8000),
        const NumberNode(0),
        const NumberNode(0),
        const NumberNode(0.1),
      ]);
      expect((result as NumberValue).value, closeTo(0.007701, 0.0001));
    });

    test('convergence failure returns #NUM!', () {
      // nper=1, pmt=0, pv=-1, fv=1: f(rate) = -1*(1+rate) + 0 + 1 = -rate
      // df(rate) = -1, so Newton gives rate = rate - (-rate)/(-1) = 0
      // f(0) = 0 => converges to 0. We need a truly pathological case.
      // With nper=10, pmt=1, pv=1, fv=1e15: huge fv makes it impossible.
      final result = eval(registry.get('RATE')!, [
        const NumberNode(10),
        const NumberNode(1),
        const NumberNode(1),
        const NumberNode(1e15),
        const NumberNode(0),
        const NumberNode(0.1),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('IRR', () {
    test('basic IRR', () {
      context.setRange('A1:F1', const FormulaValue.range([
        [
          NumberValue(-70000),
          NumberValue(12000),
          NumberValue(15000),
          NumberValue(18000),
          NumberValue(21000),
          NumberValue(26000),
        ]
      ]));
      final result = eval(registry.get('IRR')!, [_range('A1:F1')]);
      expect((result as NumberValue).value, closeTo(0.0866, 0.001));
    });

    test('with guess', () {
      context.setRange('A1:F1', const FormulaValue.range([
        [
          NumberValue(-70000),
          NumberValue(12000),
          NumberValue(15000),
          NumberValue(18000),
          NumberValue(21000),
          NumberValue(26000),
        ]
      ]));
      final result = eval(registry.get('IRR')!, [
        _range('A1:F1'),
        const NumberNode(-0.1),
      ]);
      expect((result as NumberValue).value, closeTo(0.0866, 0.001));
    });

    test('simple IRR', () {
      context.setRange('A1:B1', const FormulaValue.range([
        [NumberValue(-100), NumberValue(110)]
      ]));
      final result = eval(registry.get('IRR')!, [_range('A1:B1')]);
      expect((result as NumberValue).value, closeTo(0.10, 0.001));
    });
  });

  group('XNPV', () {
    test('basic XNPV', () {
      context.setRange('A1:E1', const FormulaValue.range([
        [
          NumberValue(-10000),
          NumberValue(2750),
          NumberValue(4250),
          NumberValue(3250),
          NumberValue(2750),
        ]
      ]));
      context.setRange('A2:E2', const FormulaValue.range([
        [
          NumberValue(39448),
          NumberValue(39531),
          NumberValue(39722),
          NumberValue(40057),
          NumberValue(40148),
        ]
      ]));
      final result = eval(registry.get('XNPV')!, [
        const NumberNode(0.09),
        _range('A1:E1'),
        _range('A2:E2'),
      ]);
      expect((result as NumberValue).value, closeTo(1826.20, 1.0));
    });
  });

  group('XIRR', () {
    test('basic XIRR', () {
      context.setRange('A1:E1', const FormulaValue.range([
        [
          NumberValue(-10000),
          NumberValue(2750),
          NumberValue(4250),
          NumberValue(3250),
          NumberValue(2750),
        ]
      ]));
      context.setRange('A2:E2', const FormulaValue.range([
        [
          NumberValue(39448),
          NumberValue(39531),
          NumberValue(39722),
          NumberValue(40057),
          NumberValue(40148),
        ]
      ]));
      final result = eval(registry.get('XIRR')!, [
        _range('A1:E1'),
        _range('A2:E2'),
      ]);
      expect((result as NumberValue).value, closeTo(0.2796, 0.01));
    });

    test('with guess', () {
      context.setRange('A1:B1', const FormulaValue.range([
        [NumberValue(-100), NumberValue(110)]
      ]));
      context.setRange('A2:B2', const FormulaValue.range([
        [
          NumberValue(45292),
          NumberValue(45657),
        ]
      ]));
      final result = eval(registry.get('XIRR')!, [
        _range('A1:B1'),
        _range('A2:B2'),
        const NumberNode(0.1),
      ]);
      expect((result as NumberValue).value, closeTo(0.1003, 0.01));
    });
  });

  // ─── Wave 4 — MIRR + Complex Depreciation ────────────────────

  group('MIRR', () {
    test('basic MIRR', () {
      context.setRange('A1:F1', const FormulaValue.range([
        [
          NumberValue(-120000),
          NumberValue(39000),
          NumberValue(30000),
          NumberValue(21000),
          NumberValue(37000),
          NumberValue(46000),
        ]
      ]));
      final result = eval(registry.get('MIRR')!, [
        _range('A1:F1'),
        const NumberNode(0.10),
        const NumberNode(0.12),
      ]);
      expect((result as NumberValue).value, closeTo(0.1261, 0.001));
    });

    test('all positive returns #DIV/0!', () {
      context.setRange('A1:B1', const FormulaValue.range([
        [NumberValue(100), NumberValue(200)]
      ]));
      final result = eval(registry.get('MIRR')!, [
        _range('A1:B1'),
        const NumberNode(0.10),
        const NumberNode(0.12),
      ]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });

    test('all negative returns #DIV/0!', () {
      context.setRange('A1:B1', const FormulaValue.range([
        [NumberValue(-100), NumberValue(-200)]
      ]));
      final result = eval(registry.get('MIRR')!, [
        _range('A1:B1'),
        const NumberNode(0.10),
        const NumberNode(0.12),
      ]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });

  group('DB', () {
    test('first year', () {
      final result = eval(registry.get('DB')!, [
        const NumberNode(1000000),
        const NumberNode(100000),
        const NumberNode(6),
        const NumberNode(1),
      ]);
      expect((result as NumberValue).value, closeTo(319000, 1));
    });

    test('second year', () {
      final result = eval(registry.get('DB')!, [
        const NumberNode(1000000),
        const NumberNode(100000),
        const NumberNode(6),
        const NumberNode(2),
      ]);
      expect((result as NumberValue).value, closeTo(217239, 1));
    });

    test('with fractional first year month', () {
      final result = eval(registry.get('DB')!, [
        const NumberNode(1000000),
        const NumberNode(100000),
        const NumberNode(6),
        const NumberNode(1),
        const NumberNode(7),
      ]);
      expect((result as NumberValue).value, closeTo(186083.33, 1));
    });

    test('cost = 0 returns 0', () {
      final result = eval(registry.get('DB')!, [
        const NumberNode(0),
        const NumberNode(0),
        const NumberNode(6),
        const NumberNode(1),
      ]);
      expect(result, const NumberValue(0));
    });
  });

  group('DDB', () {
    test('first year', () {
      final result = eval(registry.get('DDB')!, [
        const NumberNode(2400),
        const NumberNode(300),
        const NumberNode(10),
        const NumberNode(1),
      ]);
      expect((result as NumberValue).value, closeTo(480, 0.01));
    });

    test('second year', () {
      final result = eval(registry.get('DDB')!, [
        const NumberNode(2400),
        const NumberNode(300),
        const NumberNode(10),
        const NumberNode(2),
      ]);
      expect((result as NumberValue).value, closeTo(384, 0.01));
    });

    test('tenth year', () {
      final result = eval(registry.get('DDB')!, [
        const NumberNode(2400),
        const NumberNode(300),
        const NumberNode(10),
        const NumberNode(10),
      ]);
      expect((result as NumberValue).value, closeTo(22.12, 0.1));
    });

    test('with custom factor', () {
      final result = eval(registry.get('DDB')!, [
        const NumberNode(2400),
        const NumberNode(300),
        const NumberNode(10),
        const NumberNode(1),
        const NumberNode(1.5),
      ]);
      expect((result as NumberValue).value, closeTo(360, 0.01));
    });
  });

  group('VDB', () {
    test('single period same as DDB', () {
      final result = eval(registry.get('VDB')!, [
        const NumberNode(2400),
        const NumberNode(300),
        const NumberNode(10),
        const NumberNode(0),
        const NumberNode(1),
      ]);
      expect((result as NumberValue).value, closeTo(480, 0.01));
    });

    test('multiple periods', () {
      final result = eval(registry.get('VDB')!, [
        const NumberNode(2400),
        const NumberNode(300),
        const NumberNode(10),
        const NumberNode(0),
        const NumberNode(3),
      ]);
      expect((result as NumberValue).value, closeTo(1171.2, 0.1));
    });

    test('fractional periods', () {
      final result = eval(registry.get('VDB')!, [
        const NumberNode(2400),
        const NumberNode(300),
        const NumberNode(10),
        const NumberNode(0),
        const NumberNode(0.5),
      ]);
      expect((result as NumberValue).value, closeTo(240, 0.01));
    });

    test('no switch to SLN', () {
      final result = eval(registry.get('VDB')!, [
        const NumberNode(2400),
        const NumberNode(300),
        const NumberNode(10),
        const NumberNode(0),
        const NumberNode(1),
        const NumberNode(2),
        const BooleanNode(true),
      ]);
      expect((result as NumberValue).value, closeTo(480, 0.01));
    });
  });

  // ─── Wave 5 — Simple Bond/Security Functions ─────────────────

  group('DISC', () {
    test('basic discount rate', () {
      final result = eval(registry.get('DISC')!, [
        const NumberNode(39838),
        const NumberNode(40055),
        const NumberNode(97.975),
        const NumberNode(100),
        const NumberNode(2),
      ]);
      expect((result as NumberValue).value, closeTo(0.03348, 0.001));
    });
  });

  group('INTRATE', () {
    test('basic interest rate', () {
      final result = eval(registry.get('INTRATE')!, [
        const NumberNode(39838),
        const NumberNode(40055),
        const NumberNode(1000000),
        const NumberNode(1014420),
        const NumberNode(2),
      ]);
      expect((result as NumberValue).value, closeTo(0.02385, 0.001));
    });
  });

  group('RECEIVED', () {
    test('basic received amount', () {
      final result = eval(registry.get('RECEIVED')!, [
        const NumberNode(39838),
        const NumberNode(40055),
        const NumberNode(1000000),
        const NumberNode(0.0575),
        const NumberNode(2),
      ]);
      expect((result as NumberValue).value, closeTo(1035904.15, 1));
    });
  });

  group('PRICEDISC', () {
    test('basic discounted price', () {
      final result = eval(registry.get('PRICEDISC')!, [
        const NumberNode(39838),
        const NumberNode(40055),
        const NumberNode(0.0525),
        const NumberNode(100),
        const NumberNode(2),
      ]);
      expect((result as NumberValue).value, closeTo(96.83, 0.01));
    });
  });

  group('PRICEMAT', () {
    test('basic price at maturity', () {
      final result = eval(registry.get('PRICEMAT')!, [
        const NumberNode(39829),
        const NumberNode(40211),
        const NumberNode(39768),
        const NumberNode(0.0575),
        const NumberNode(0.0625),
        const NumberNode(0),
      ]);
      expect((result as NumberValue).value, closeTo(99.48, 0.1));
    });
  });

  group('ACCRINT', () {
    test('basic accrued interest', () {
      final result = eval(registry.get('ACCRINT')!, [
        const NumberNode(39508),
        const NumberNode(39691),
        const NumberNode(39600),
        const NumberNode(0.10),
        const NumberNode(1000),
        const NumberNode(2),
        const NumberNode(0),
      ]);
      expect((result as NumberValue).value, closeTo(25.0, 0.01));
    });

    test('invalid frequency returns #NUM!', () {
      final result = eval(registry.get('ACCRINT')!, [
        const NumberNode(39508),
        const NumberNode(39691),
        const NumberNode(39600),
        const NumberNode(0.10),
        const NumberNode(1000),
        const NumberNode(5),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  // ─── Wave 6 — Complex Bond Functions ─────────────────────────

  group('PRICE', () {
    test('basic bond price', () {
      final result = eval(registry.get('PRICE')!, [
        const NumberNode(39829),
        const NumberNode(43789),
        const NumberNode(0.0575),
        const NumberNode(0.065),
        const NumberNode(100),
        const NumberNode(2),
        const NumberNode(0),
      ]);
      expect((result as NumberValue).value, closeTo(94.22, 0.1));
    });

    test('at par when rate == yield', () {
      final result = eval(registry.get('PRICE')!, [
        const NumberNode(45292),
        const NumberNode(49071),
        const NumberNode(0.05),
        const NumberNode(0.05),
        const NumberNode(100),
        const NumberNode(2),
        const NumberNode(0),
      ]);
      expect((result as NumberValue).value, closeTo(100, 1));
    });
  });

  group('YIELD', () {
    test('basic bond yield', () {
      final result = eval(registry.get('YIELD')!, [
        const NumberNode(39829),
        const NumberNode(43789),
        const NumberNode(0.0575),
        const NumberNode(94.21788),
        const NumberNode(100),
        const NumberNode(2),
        const NumberNode(0),
      ]);
      expect((result as NumberValue).value, closeTo(0.065, 0.001));
    });

    test('at par price yields coupon rate', () {
      final result = eval(registry.get('YIELD')!, [
        const NumberNode(45292),
        const NumberNode(49071),
        const NumberNode(0.05),
        const NumberNode(100),
        const NumberNode(100),
        const NumberNode(2),
        const NumberNode(0),
      ]);
      expect((result as NumberValue).value, closeTo(0.05, 0.001));
    });
  });

  group('DURATION', () {
    test('basic Macaulay duration', () {
      final result = eval(registry.get('DURATION')!, [
        const NumberNode(39829),
        const NumberNode(43789),
        const NumberNode(0.08),
        const NumberNode(0.09),
        const NumberNode(2),
        const NumberNode(1),
      ]);
      expect((result as NumberValue).value, closeTo(7.22, 0.1));
    });

    test('zero coupon has duration equal to maturity', () {
      final result = eval(registry.get('DURATION')!, [
        const NumberNode(45292),
        const NumberNode(45658),
        const NumberNode(0),
        const NumberNode(0.05),
        const NumberNode(2),
        const NumberNode(0),
      ]);
      expect((result as NumberValue).value, closeTo(1.0, 0.05));
    });

    test('invalid frequency returns #NUM!', () {
      final result = eval(registry.get('DURATION')!, [
        const NumberNode(45292),
        const NumberNode(45658),
        const NumberNode(0.05),
        const NumberNode(0.05),
        const NumberNode(3),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('MDURATION', () {
    test('basic modified duration', () {
      final result = eval(registry.get('MDURATION')!, [
        const NumberNode(39829),
        const NumberNode(43789),
        const NumberNode(0.08),
        const NumberNode(0.09),
        const NumberNode(2),
        const NumberNode(1),
      ]);
      expect((result as NumberValue).value, closeTo(6.91, 0.1));
    });

    test('invalid frequency returns #NUM!', () {
      final result = eval(registry.get('MDURATION')!, [
        const NumberNode(45292),
        const NumberNode(45658),
        const NumberNode(0.05),
        const NumberNode(0.05),
        const NumberNode(3),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });
}
