import 'package:a1/a1.dart';
import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/functions/function.dart';
import 'package:worksheet_formula/src/functions/registry.dart';
import 'package:worksheet_formula/src/functions/database.dart';

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

/// Helper to create a RangeRefNode for a given reference string.
RangeRefNode _rangeRef(String ref) => RangeRefNode(A1Reference.parse(ref));

void main() {
  late FunctionRegistry registry;
  late _TestContext context;

  // Test database:
  // | Name    | Department  | Salary | Age |
  // | Alice   | Sales       | 50000  | 30  |
  // | Bob     | Sales       | 60000  | 45  |
  // | Charlie | Engineering | 80000  | 35  |
  // | Diana   | Sales       | 55000  | 28  |
  // | Eve     | Engineering | 90000  | 40  |
  // | Frank   | HR          | 45000  | 50  |
  const testDatabase = RangeValue([
    [TextValue('Name'), TextValue('Department'), TextValue('Salary'), TextValue('Age')],
    [TextValue('Alice'), TextValue('Sales'), NumberValue(50000), NumberValue(30)],
    [TextValue('Bob'), TextValue('Sales'), NumberValue(60000), NumberValue(45)],
    [TextValue('Charlie'), TextValue('Engineering'), NumberValue(80000), NumberValue(35)],
    [TextValue('Diana'), TextValue('Sales'), NumberValue(55000), NumberValue(28)],
    [TextValue('Eve'), TextValue('Engineering'), NumberValue(90000), NumberValue(40)],
    [TextValue('Frank'), TextValue('HR'), NumberValue(45000), NumberValue(50)],
  ]);

  // Criteria constants
  const critSales = RangeValue([
    [TextValue('Department')],
    [TextValue('Sales')],
  ]);

  const critEngineering = RangeValue([
    [TextValue('Department')],
    [TextValue('Engineering')],
  ]);

  const critSalaryGt50k = RangeValue([
    [TextValue('Salary')],
    [TextValue('>50000')],
  ]);

  const critSalesAndSalaryGt50k = RangeValue([
    [TextValue('Department'), TextValue('Salary')],
    [TextValue('Sales'), TextValue('>50000')],
  ]);

  const critSalesOrEngineering = RangeValue([
    [TextValue('Department')],
    [TextValue('Sales')],
    [TextValue('Engineering')],
  ]);

  const critHR = RangeValue([
    [TextValue('Department')],
    [TextValue('HR')],
  ]);

  const critNoMatch = RangeValue([
    [TextValue('Department')],
    [TextValue('Marketing')],
  ]);

  const critAgeNot50 = RangeValue([
    [TextValue('Age')],
    [TextValue('<>50')],
  ]);

  const critEmpty = RangeValue([
    [TextValue('Department')],
    [EmptyValue()],
  ]);

  // Range references used as node args
  // DB = A1:D7, criteria get various ranges
  const dbRef = 'A1:D7';
  const crRef = 'F1:F2'; // single-column, 2-row criteria
  const crRef2 = 'F1:G2'; // 2-column criteria (AND)
  const crRef3 = 'F1:F3'; // single-column, 3-row criteria (OR)

  setUp(() {
    registry = FunctionRegistry(registerBuiltIns: false);
    registerDatabaseFunctions(registry);
    context = _TestContext(registry);
    // Always set the database
    context.setRange(dbRef, testDatabase);
  });

  FormulaValue eval(FormulaFunction fn, List<FormulaNode> args) =>
      fn.call(args, context);

  /// Helper: set criteria and return a list of 3 args [dbNode, fieldNode, critNode].
  List<FormulaNode> dbArgs(String field, RangeValue criteria,
      {String critRange = crRef, bool numericField = false}) {
    context.setRange(critRange, criteria);
    return [
      _rangeRef(dbRef),
      if (numericField)
        NumberNode(int.parse(field))
      else
        TextNode(field),
      _rangeRef(critRange),
    ];
  }

  // =========================================================================
  // DSUM
  // =========================================================================
  group('DSUM', () {
    test('sums Salary for Sales department', () {
      final result = eval(
          registry.get('DSUM')!, dbArgs('Salary', critSales));
      expect(result, const NumberValue(165000));
    });

    test('sums with operator criteria >50000', () {
      final result = eval(
          registry.get('DSUM')!, dbArgs('Salary', critSalaryGt50k));
      expect(result, const NumberValue(285000));
    });

    test('sums with AND criteria (Sales AND Salary>50000)', () {
      final result = eval(registry.get('DSUM')!,
          dbArgs('Salary', critSalesAndSalaryGt50k, critRange: crRef2));
      expect(result, const NumberValue(115000));
    });

    test('sums with OR criteria (Sales OR Engineering)', () {
      final result = eval(registry.get('DSUM')!,
          dbArgs('Salary', critSalesOrEngineering, critRange: crRef3));
      expect(result, const NumberValue(335000));
    });

    test('numeric field identifier (column 3 = Salary)', () {
      final result = eval(registry.get('DSUM')!,
          dbArgs('3', critSales, numericField: true));
      expect(result, const NumberValue(165000));
    });

    test('no matching rows returns 0', () {
      final result = eval(
          registry.get('DSUM')!, dbArgs('Salary', critNoMatch));
      expect(result, const NumberValue(0));
    });

    test('empty criteria matches all rows', () {
      final result = eval(
          registry.get('DSUM')!, dbArgs('Salary', critEmpty));
      expect(result, const NumberValue(380000));
    });
  });

  // =========================================================================
  // DAVERAGE
  // =========================================================================
  group('DAVERAGE', () {
    test('averages Salary for Sales department', () {
      final result = eval(
          registry.get('DAVERAGE')!, dbArgs('Salary', critSales));
      expect(result, const NumberValue(55000));
    });

    test('averages with operator criteria', () {
      final result = eval(
          registry.get('DAVERAGE')!, dbArgs('Salary', critSalaryGt50k));
      expect(result, const NumberValue(71250));
    });

    test('averages with AND criteria', () {
      final result = eval(registry.get('DAVERAGE')!,
          dbArgs('Salary', critSalesAndSalaryGt50k, critRange: crRef2));
      expect(result, const NumberValue(57500));
    });

    test('averages with OR criteria', () {
      final result = eval(registry.get('DAVERAGE')!,
          dbArgs('Salary', critSalesOrEngineering, critRange: crRef3));
      expect(result, const NumberValue(67000));
    });

    test('numeric field identifier', () {
      final result = eval(registry.get('DAVERAGE')!,
          dbArgs('3', critSales, numericField: true));
      expect(result, const NumberValue(55000));
    });

    test('no matching rows returns #DIV/0!', () {
      final result = eval(
          registry.get('DAVERAGE')!, dbArgs('Salary', critNoMatch));
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });

  // =========================================================================
  // DCOUNT
  // =========================================================================
  group('DCOUNT', () {
    test('counts numeric cells in Salary for Sales', () {
      final result = eval(
          registry.get('DCOUNT')!, dbArgs('Salary', critSales));
      expect(result, const NumberValue(3));
    });

    test('counts numeric cells in Name for Sales (text column = 0)', () {
      final result = eval(
          registry.get('DCOUNT')!, dbArgs('Name', critSales));
      expect(result, const NumberValue(0));
    });

    test('counts with operator criteria', () {
      final result = eval(
          registry.get('DCOUNT')!, dbArgs('Salary', critSalaryGt50k));
      expect(result, const NumberValue(4));
    });

    test('counts with AND criteria', () {
      final result = eval(registry.get('DCOUNT')!,
          dbArgs('Salary', critSalesAndSalaryGt50k, critRange: crRef2));
      expect(result, const NumberValue(2));
    });

    test('counts with OR criteria', () {
      final result = eval(registry.get('DCOUNT')!,
          dbArgs('Salary', critSalesOrEngineering, critRange: crRef3));
      expect(result, const NumberValue(5));
    });

    test('no matching rows returns 0', () {
      final result = eval(
          registry.get('DCOUNT')!, dbArgs('Salary', critNoMatch));
      expect(result, const NumberValue(0));
    });
  });

  // =========================================================================
  // DCOUNTA
  // =========================================================================
  group('DCOUNTA', () {
    test('counts non-empty cells in Name for Sales', () {
      final result = eval(
          registry.get('DCOUNTA')!, dbArgs('Name', critSales));
      expect(result, const NumberValue(3));
    });

    test('counts non-empty cells in Salary for Sales', () {
      final result = eval(
          registry.get('DCOUNTA')!, dbArgs('Salary', critSales));
      expect(result, const NumberValue(3));
    });

    test('counts with operator criteria', () {
      final result = eval(
          registry.get('DCOUNTA')!, dbArgs('Name', critSalaryGt50k));
      expect(result, const NumberValue(4));
    });

    test('counts with AND criteria', () {
      final result = eval(registry.get('DCOUNTA')!,
          dbArgs('Name', critSalesAndSalaryGt50k, critRange: crRef2));
      expect(result, const NumberValue(2));
    });

    test('no matching rows returns 0', () {
      final result = eval(
          registry.get('DCOUNTA')!, dbArgs('Name', critNoMatch));
      expect(result, const NumberValue(0));
    });

    test('numeric field identifier', () {
      final result = eval(registry.get('DCOUNTA')!,
          dbArgs('1', critSales, numericField: true));
      expect(result, const NumberValue(3));
    });
  });

  // =========================================================================
  // DMAX
  // =========================================================================
  group('DMAX', () {
    test('max Salary for Sales department', () {
      final result = eval(
          registry.get('DMAX')!, dbArgs('Salary', critSales));
      expect(result, const NumberValue(60000));
    });

    test('max with operator criteria', () {
      final result = eval(
          registry.get('DMAX')!, dbArgs('Salary', critSalaryGt50k));
      expect(result, const NumberValue(90000));
    });

    test('max with AND criteria', () {
      final result = eval(registry.get('DMAX')!,
          dbArgs('Salary', critSalesAndSalaryGt50k, critRange: crRef2));
      expect(result, const NumberValue(60000));
    });

    test('max with OR criteria', () {
      final result = eval(registry.get('DMAX')!,
          dbArgs('Salary', critSalesOrEngineering, critRange: crRef3));
      expect(result, const NumberValue(90000));
    });

    test('numeric field identifier', () {
      final result = eval(registry.get('DMAX')!,
          dbArgs('4', critSales, numericField: true));
      expect(result, const NumberValue(45));
    });

    test('no matching rows returns 0', () {
      final result = eval(
          registry.get('DMAX')!, dbArgs('Salary', critNoMatch));
      expect(result, const NumberValue(0));
    });
  });

  // =========================================================================
  // DMIN
  // =========================================================================
  group('DMIN', () {
    test('min Salary for Sales department', () {
      final result = eval(
          registry.get('DMIN')!, dbArgs('Salary', critSales));
      expect(result, const NumberValue(50000));
    });

    test('min with operator criteria', () {
      final result = eval(
          registry.get('DMIN')!, dbArgs('Salary', critSalaryGt50k));
      expect(result, const NumberValue(55000));
    });

    test('min with AND criteria', () {
      final result = eval(registry.get('DMIN')!,
          dbArgs('Salary', critSalesAndSalaryGt50k, critRange: crRef2));
      expect(result, const NumberValue(55000));
    });

    test('min with OR criteria', () {
      final result = eval(registry.get('DMIN')!,
          dbArgs('Age', critSalesOrEngineering, critRange: crRef3));
      expect(result, const NumberValue(28));
    });

    test('numeric field identifier', () {
      final result = eval(registry.get('DMIN')!,
          dbArgs('4', critEngineering, numericField: true));
      expect(result, const NumberValue(35));
    });

    test('no matching rows returns 0', () {
      final result = eval(
          registry.get('DMIN')!, dbArgs('Salary', critNoMatch));
      expect(result, const NumberValue(0));
    });
  });

  // =========================================================================
  // DGET
  // =========================================================================
  group('DGET', () {
    test('returns single matching value', () {
      final result = eval(
          registry.get('DGET')!, dbArgs('Salary', critHR));
      expect(result, const NumberValue(45000));
    });

    test('returns text value for single match', () {
      final result = eval(
          registry.get('DGET')!, dbArgs('Name', critHR));
      expect(result, const TextValue('Frank'));
    });

    test('0 matches returns #VALUE!', () {
      final result = eval(
          registry.get('DGET')!, dbArgs('Salary', critNoMatch));
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('2+ matches returns #NUM!', () {
      final result = eval(
          registry.get('DGET')!, dbArgs('Salary', critSales));
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('numeric field identifier', () {
      final result = eval(registry.get('DGET')!,
          dbArgs('3', critHR, numericField: true));
      expect(result, const NumberValue(45000));
    });

    test('returns value with AND criteria narrowing to one', () {
      const critSalesAge45 = RangeValue([
        [TextValue('Department'), TextValue('Age')],
        [TextValue('Sales'), NumberValue(45)],
      ]);
      final result = eval(registry.get('DGET')!,
          dbArgs('Name', critSalesAge45, critRange: crRef2));
      expect(result, const TextValue('Bob'));
    });
  });

  // =========================================================================
  // DPRODUCT
  // =========================================================================
  group('DPRODUCT', () {
    test('product of Age for Engineering department', () {
      final result = eval(
          registry.get('DPRODUCT')!, dbArgs('Age', critEngineering));
      expect(result, const NumberValue(1400));
    });

    test('product of Age for HR department (single value)', () {
      final result = eval(
          registry.get('DPRODUCT')!, dbArgs('Age', critHR));
      expect(result, const NumberValue(50));
    });

    test('product with operator criteria', () {
      const critAgeGt40 = RangeValue([
        [TextValue('Age')],
        [TextValue('>40')],
      ]);
      final result = eval(
          registry.get('DPRODUCT')!, dbArgs('Age', critAgeGt40));
      expect(result, const NumberValue(2250));
    });

    test('no matching rows returns 0', () {
      final result = eval(
          registry.get('DPRODUCT')!, dbArgs('Salary', critNoMatch));
      expect(result, const NumberValue(0));
    });

    test('numeric field identifier', () {
      final result = eval(registry.get('DPRODUCT')!,
          dbArgs('4', critEngineering, numericField: true));
      expect(result, const NumberValue(1400));
    });
  });

  // =========================================================================
  // DSTDEV
  // =========================================================================
  group('DSTDEV', () {
    test('sample std deviation of Salary for Sales', () {
      final result = eval(
          registry.get('DSTDEV')!, dbArgs('Salary', critSales));
      expect((result as NumberValue).value, closeTo(5000, 0.01));
    });

    test('sample std deviation of Salary for Engineering', () {
      final result = eval(
          registry.get('DSTDEV')!, dbArgs('Salary', critEngineering));
      expect((result as NumberValue).value, closeTo(7071.07, 0.1));
    });

    test('single value returns #DIV/0! (need 2+ for sample)', () {
      final result = eval(
          registry.get('DSTDEV')!, dbArgs('Salary', critHR));
      expect(result, const ErrorValue(FormulaError.divZero));
    });

    test('no matching rows returns #DIV/0!', () {
      final result = eval(
          registry.get('DSTDEV')!, dbArgs('Salary', critNoMatch));
      expect(result, const ErrorValue(FormulaError.divZero));
    });

    test('numeric field identifier', () {
      final result = eval(registry.get('DSTDEV')!,
          dbArgs('3', critSales, numericField: true));
      expect((result as NumberValue).value, closeTo(5000, 0.01));
    });
  });

  // =========================================================================
  // DSTDEVP
  // =========================================================================
  group('DSTDEVP', () {
    test('population std deviation of Salary for Sales', () {
      final result = eval(
          registry.get('DSTDEVP')!, dbArgs('Salary', critSales));
      expect((result as NumberValue).value, closeTo(4082.48, 0.1));
    });

    test('population std deviation of Salary for Engineering', () {
      final result = eval(
          registry.get('DSTDEVP')!, dbArgs('Salary', critEngineering));
      expect((result as NumberValue).value, closeTo(5000, 0.01));
    });

    test('single value returns 0 (population stdev of one value)', () {
      final result = eval(
          registry.get('DSTDEVP')!, dbArgs('Salary', critHR));
      expect((result as NumberValue).value, closeTo(0, 0.01));
    });

    test('no matching rows returns #DIV/0!', () {
      final result = eval(
          registry.get('DSTDEVP')!, dbArgs('Salary', critNoMatch));
      expect(result, const ErrorValue(FormulaError.divZero));
    });

    test('numeric field identifier', () {
      final result = eval(registry.get('DSTDEVP')!,
          dbArgs('3', critSales, numericField: true));
      expect((result as NumberValue).value, closeTo(4082.48, 0.1));
    });
  });

  // =========================================================================
  // DVAR
  // =========================================================================
  group('DVAR', () {
    test('sample variance of Salary for Sales', () {
      final result = eval(
          registry.get('DVAR')!, dbArgs('Salary', critSales));
      expect(result, const NumberValue(25000000));
    });

    test('sample variance of Salary for Engineering', () {
      final result = eval(
          registry.get('DVAR')!, dbArgs('Salary', critEngineering));
      expect(result, const NumberValue(50000000));
    });

    test('single value returns #DIV/0!', () {
      final result = eval(
          registry.get('DVAR')!, dbArgs('Salary', critHR));
      expect(result, const ErrorValue(FormulaError.divZero));
    });

    test('no matching rows returns #DIV/0!', () {
      final result = eval(
          registry.get('DVAR')!, dbArgs('Salary', critNoMatch));
      expect(result, const ErrorValue(FormulaError.divZero));
    });

    test('numeric field identifier', () {
      final result = eval(registry.get('DVAR')!,
          dbArgs('3', critSales, numericField: true));
      expect(result, const NumberValue(25000000));
    });

    test('with OR criteria', () {
      final result = eval(registry.get('DVAR')!,
          dbArgs('Salary', critSalesOrEngineering, critRange: crRef3));
      expect(result, const NumberValue(295000000));
    });
  });

  // =========================================================================
  // DVARP
  // =========================================================================
  group('DVARP', () {
    test('population variance of Salary for Sales', () {
      final result = eval(
          registry.get('DVARP')!, dbArgs('Salary', critSales));
      final v = (result as NumberValue).value;
      expect(v, closeTo(16666666.67, 0.1));
    });

    test('population variance of Salary for Engineering', () {
      final result = eval(
          registry.get('DVARP')!, dbArgs('Salary', critEngineering));
      expect(result, const NumberValue(25000000));
    });

    test('single value returns 0', () {
      final result = eval(
          registry.get('DVARP')!, dbArgs('Salary', critHR));
      expect((result as NumberValue).value, closeTo(0, 0.01));
    });

    test('no matching rows returns #DIV/0!', () {
      final result = eval(
          registry.get('DVARP')!, dbArgs('Salary', critNoMatch));
      expect(result, const ErrorValue(FormulaError.divZero));
    });

    test('numeric field identifier', () {
      final result = eval(registry.get('DVARP')!,
          dbArgs('3', critSales, numericField: true));
      final v = (result as NumberValue).value;
      expect(v, closeTo(16666666.67, 0.1));
    });
  });

  // =========================================================================
  // Cross-cutting error tests
  // =========================================================================
  group('Error handling', () {
    test('non-range database returns #VALUE!', () {
      context.setRange(crRef, critSales);
      final result = eval(registry.get('DSUM')!, [
        const NumberNode(42),
        const TextNode('Salary'),
        _rangeRef(crRef),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('non-range criteria returns #VALUE!', () {
      final result = eval(registry.get('DSUM')!, [
        _rangeRef(dbRef),
        const TextNode('Salary'),
        const NumberNode(42),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('invalid field name returns #VALUE!', () {
      final result = eval(
          registry.get('DSUM')!, dbArgs('NonExistent', critSales));
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('field number 0 (out of range) returns #VALUE!', () {
      final result = eval(registry.get('DSUM')!,
          dbArgs('0', critSales, numericField: true));
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('field number > column count returns #VALUE!', () {
      final result = eval(registry.get('DSUM')!,
          dbArgs('5', critSales, numericField: true));
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('empty database (headers only) returns 0 for DSUM', () {
      const emptyDb = RangeValue([
        [TextValue('Name'), TextValue('Salary')],
      ]);
      const crit = RangeValue([
        [TextValue('Name')],
        [TextValue('Alice')],
      ]);
      context.setRange('A1:B1', emptyDb);
      context.setRange(crRef, crit);
      final result = eval(registry.get('DSUM')!, [
        _rangeRef('A1:B1'),
        const TextNode('Salary'),
        _rangeRef(crRef),
      ]);
      expect(result, const NumberValue(0));
    });

    test('unmatched criteria header is ignored', () {
      const critUnmatched = RangeValue([
        [TextValue('Department'), TextValue('NonExistent')],
        [TextValue('Sales'), TextValue('xyz')],
      ]);
      final result = eval(registry.get('DSUM')!,
          dbArgs('Salary', critUnmatched, critRange: crRef2));
      expect(result, const NumberValue(165000));
    });

    test('case-insensitive header matching in database', () {
      final result = eval(
          registry.get('DSUM')!, dbArgs('salary', critSales));
      expect(result, const NumberValue(165000));
    });

    test('case-insensitive header matching in criteria', () {
      const critLowerCase = RangeValue([
        [TextValue('department')],
        [TextValue('Sales')],
      ]);
      final result = eval(
          registry.get('DSUM')!, dbArgs('Salary', critLowerCase));
      expect(result, const NumberValue(165000));
    });

    test('<> operator criteria works', () {
      final result = eval(
          registry.get('DSUM')!, dbArgs('Salary', critAgeNot50));
      expect(result, const NumberValue(335000));
    });

    test('DAVERAGE with empty database returns #DIV/0!', () {
      const emptyDb = RangeValue([
        [TextValue('Name'), TextValue('Salary')],
      ]);
      const crit = RangeValue([
        [TextValue('Name')],
        [TextValue('Alice')],
      ]);
      context.setRange('A1:B1', emptyDb);
      context.setRange(crRef, crit);
      final result = eval(registry.get('DAVERAGE')!, [
        _rangeRef('A1:B1'),
        const TextNode('Salary'),
        _rangeRef(crRef),
      ]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });

    test('DGET with empty database returns #VALUE!', () {
      const emptyDb = RangeValue([
        [TextValue('Name'), TextValue('Salary')],
      ]);
      const crit = RangeValue([
        [TextValue('Name')],
        [TextValue('Alice')],
      ]);
      context.setRange('A1:B1', emptyDb);
      context.setRange(crRef, crit);
      final result = eval(registry.get('DGET')!, [
        _rangeRef('A1:B1'),
        const TextNode('Salary'),
        _rangeRef(crRef),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('boolean field value returns #VALUE!', () {
      context.setRange(crRef, critSales);
      final result = eval(registry.get('DSUM')!, [
        _rangeRef(dbRef),
        const BooleanNode(true),
        _rangeRef(crRef),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });
}
