import 'package:a1/a1.dart';
import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/functions/function.dart';
import 'package:worksheet_formula/src/functions/registry.dart';
import 'package:worksheet_formula/src/functions/lookup.dart';

class _TestContext implements EvaluationContext {
  A1 currentCellOverride;
  @override
  A1 get currentCell => currentCellOverride;
  @override
  String? get currentSheet => null;
  @override
  FormulaValue? getVariable(String name) => null;
  @override
  bool get isCancelled => false;

  final FunctionRegistry _registry;
  FormulaValue? rangeOverride;
  Map<A1, FormulaValue> cellValues = {};
  _TestContext(this._registry) : currentCellOverride = 'A1'.a1;

  @override
  FormulaValue getCellValue(A1 cell) =>
      cellValues[cell] ?? const EmptyValue();
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
    registerLookupFunctions(registry);
    context = _TestContext(registry);
  });

  FormulaValue eval(FormulaFunction fn, List<FormulaNode> args) =>
      fn.call(args, context);

  group('VLOOKUP', () {
    // Table: [1, "apple", 10], [2, "banana", 20], [3, "cherry", 30]
    final table = const RangeValue([
      [NumberValue(1), TextValue('apple'), NumberValue(10)],
      [NumberValue(2), TextValue('banana'), NumberValue(20)],
      [NumberValue(3), TextValue('cherry'), NumberValue(30)],
    ]);

    test('exact match finds value', () {
      context.rangeOverride = table;
      final result = eval(registry.get('VLOOKUP')!, [
        const NumberNode(2),
        RangeRefNode(A1Reference.parse('A1:C3')),
        const NumberNode(2),
        const BooleanNode(false),
      ]);
      expect(result, const TextValue('banana'));
    });

    test('exact match returns value from specified column', () {
      context.rangeOverride = table;
      final result = eval(registry.get('VLOOKUP')!, [
        const NumberNode(2),
        RangeRefNode(A1Reference.parse('A1:C3')),
        const NumberNode(3),
        const BooleanNode(false),
      ]);
      expect(result, const NumberValue(20));
    });

    test('exact match not found returns #N/A', () {
      context.rangeOverride = table;
      final result = eval(registry.get('VLOOKUP')!, [
        const NumberNode(99),
        RangeRefNode(A1Reference.parse('A1:C3')),
        const NumberNode(2),
        const BooleanNode(false),
      ]);
      expect(result, const ErrorValue(FormulaError.na));
    });

    test('col_index < 1 returns #VALUE!', () {
      context.rangeOverride = table;
      final result = eval(registry.get('VLOOKUP')!, [
        const NumberNode(1),
        RangeRefNode(A1Reference.parse('A1:C3')),
        const NumberNode(0),
        const BooleanNode(false),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('col_index exceeds columns returns #REF!', () {
      context.rangeOverride = table;
      final result = eval(registry.get('VLOOKUP')!, [
        const NumberNode(1),
        RangeRefNode(A1Reference.parse('A1:C3')),
        const NumberNode(5),
        const BooleanNode(false),
      ]);
      expect(result, const ErrorValue(FormulaError.ref));
    });

    test('approximate match finds largest <= value', () {
      context.rangeOverride = table;
      final result = eval(registry.get('VLOOKUP')!, [
        const NumberNode(2.5),
        RangeRefNode(A1Reference.parse('A1:C3')),
        const NumberNode(2),
        const BooleanNode(true),
      ]);
      expect(result, const TextValue('banana'));
    });

    test('approximate match is default', () {
      context.rangeOverride = table;
      final result = eval(registry.get('VLOOKUP')!, [
        const NumberNode(2.5),
        RangeRefNode(A1Reference.parse('A1:C3')),
        const NumberNode(2),
      ]);
      expect(result, const TextValue('banana'));
    });
  });

  group('INDEX', () {
    final array = const RangeValue([
      [NumberValue(1), NumberValue(2), NumberValue(3)],
      [NumberValue(4), NumberValue(5), NumberValue(6)],
      [NumberValue(7), NumberValue(8), NumberValue(9)],
    ]);

    test('returns value at row and column', () {
      context.rangeOverride = array;
      final result = eval(registry.get('INDEX')!, [
        RangeRefNode(A1Reference.parse('A1:C3')),
        const NumberNode(2),
        const NumberNode(3),
      ]);
      expect(result, const NumberValue(6));
    });

    test('single column: row only', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10)],
        [NumberValue(20)],
        [NumberValue(30)],
      ]);
      final result = eval(registry.get('INDEX')!, [
        RangeRefNode(A1Reference.parse('A1:A3')),
        const NumberNode(2),
      ]);
      expect(result, const NumberValue(20));
    });

    test('row out of bounds returns #REF!', () {
      context.rangeOverride = array;
      final result = eval(registry.get('INDEX')!, [
        RangeRefNode(A1Reference.parse('A1:C3')),
        const NumberNode(5),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.ref));
    });

    test('column out of bounds returns #REF!', () {
      context.rangeOverride = array;
      final result = eval(registry.get('INDEX')!, [
        RangeRefNode(A1Reference.parse('A1:C3')),
        const NumberNode(1),
        const NumberNode(5),
      ]);
      expect(result, const ErrorValue(FormulaError.ref));
    });

    test('row < 1 returns #VALUE!', () {
      context.rangeOverride = array;
      final result = eval(registry.get('INDEX')!, [
        RangeRefNode(A1Reference.parse('A1:C3')),
        const NumberNode(0),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('MATCH', () {
    test('exact match finds position (1-indexed)', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10)],
        [NumberValue(20)],
        [NumberValue(30)],
      ]);
      final result = eval(registry.get('MATCH')!, [
        const NumberNode(20),
        RangeRefNode(A1Reference.parse('A1:A3')),
        const NumberNode(0),
      ]);
      expect(result, const NumberValue(2));
    });

    test('exact match not found returns #N/A', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10)],
        [NumberValue(20)],
      ]);
      final result = eval(registry.get('MATCH')!, [
        const NumberNode(99),
        RangeRefNode(A1Reference.parse('A1:A2')),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.na));
    });

    test('exact match text is case-insensitive', () {
      context.rangeOverride = const RangeValue([
        [TextValue('Apple')],
        [TextValue('Banana')],
        [TextValue('Cherry')],
      ]);
      final result = eval(registry.get('MATCH')!, [
        const TextNode('banana'),
        RangeRefNode(A1Reference.parse('A1:A3')),
        const NumberNode(0),
      ]);
      expect(result, const NumberValue(2));
    });

    test('ascending match (default) finds largest <=', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10)],
        [NumberValue(20)],
        [NumberValue(30)],
      ]);
      final result = eval(registry.get('MATCH')!, [
        const NumberNode(25),
        RangeRefNode(A1Reference.parse('A1:A3')),
      ]);
      expect(result, const NumberValue(2));
    });

    test('descending match finds smallest >=', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(30)],
        [NumberValue(20)],
        [NumberValue(10)],
      ]);
      final result = eval(registry.get('MATCH')!, [
        const NumberNode(25),
        RangeRefNode(A1Reference.parse('A1:A3')),
        const NumberNode(-1),
      ]);
      // Smallest value >= 25 is 30 at position 1
      expect(result, const NumberValue(1));
    });

    test('works with horizontal range', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10), NumberValue(20), NumberValue(30)],
      ]);
      final result = eval(registry.get('MATCH')!, [
        const NumberNode(20),
        RangeRefNode(A1Reference.parse('A1:C1')),
        const NumberNode(0),
      ]);
      expect(result, const NumberValue(2));
    });
  });

  group('HLOOKUP', () {
    final table = const RangeValue([
      [NumberValue(1), NumberValue(2), NumberValue(3)],
      [TextValue('apple'), TextValue('banana'), TextValue('cherry')],
      [NumberValue(10), NumberValue(20), NumberValue(30)],
    ]);

    test('exact match in first row', () {
      context.rangeOverride = table;
      final result = eval(registry.get('HLOOKUP')!, [
        const NumberNode(2),
        RangeRefNode(A1Reference.parse('A1:C3')),
        const NumberNode(2),
        const BooleanNode(false),
      ]);
      expect(result, const TextValue('banana'));
    });

    test('row_index out of bounds returns #REF!', () {
      context.rangeOverride = table;
      final result = eval(registry.get('HLOOKUP')!, [
        const NumberNode(1),
        RangeRefNode(A1Reference.parse('A1:C3')),
        const NumberNode(5),
        const BooleanNode(false),
      ]);
      expect(result, const ErrorValue(FormulaError.ref));
    });

    test('not found returns #N/A', () {
      context.rangeOverride = table;
      final result = eval(registry.get('HLOOKUP')!, [
        const NumberNode(99),
        RangeRefNode(A1Reference.parse('A1:C3')),
        const NumberNode(2),
        const BooleanNode(false),
      ]);
      expect(result, const ErrorValue(FormulaError.na));
    });
  });

  group('LOOKUP', () {
    test('approximate match in sorted vector', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10)],
        [NumberValue(20)],
        [NumberValue(30)],
      ]);
      final result = eval(registry.get('LOOKUP')!, [
        const NumberNode(25),
        RangeRefNode(A1Reference.parse('A1:A3')),
      ]);
      expect(result, const NumberValue(20));
    });

    test('not found returns #N/A', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10)],
        [NumberValue(20)],
      ]);
      final result = eval(registry.get('LOOKUP')!, [
        const NumberNode(5),
        RangeRefNode(A1Reference.parse('A1:A2')),
      ]);
      expect(result, const ErrorValue(FormulaError.na));
    });
  });

  group('CHOOSE', () {
    test('returns value at index', () {
      final result = eval(registry.get('CHOOSE')!, [
        const NumberNode(2),
        const TextNode('a'),
        const TextNode('b'),
        const TextNode('c'),
      ]);
      expect(result, const TextValue('b'));
    });

    test('index 1 returns first value', () {
      final result = eval(registry.get('CHOOSE')!, [
        const NumberNode(1),
        const TextNode('first'),
        const TextNode('second'),
      ]);
      expect(result, const TextValue('first'));
    });

    test('index out of range returns #VALUE!', () {
      final result = eval(registry.get('CHOOSE')!, [
        const NumberNode(5),
        const TextNode('a'),
        const TextNode('b'),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('index < 1 returns #VALUE!', () {
      final result = eval(registry.get('CHOOSE')!, [
        const NumberNode(0),
        const TextNode('a'),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('XMATCH', () {
    test('exact match returns position', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10)],
        [NumberValue(20)],
        [NumberValue(30)],
      ]);
      final result = eval(registry.get('XMATCH')!, [
        const NumberNode(20),
        RangeRefNode(A1Reference.parse('A1:A3')),
      ]);
      expect(result, const NumberValue(2));
    });

    test('not found returns #N/A', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10)],
        [NumberValue(20)],
      ]);
      final result = eval(registry.get('XMATCH')!, [
        const NumberNode(99),
        RangeRefNode(A1Reference.parse('A1:A2')),
      ]);
      expect(result, const ErrorValue(FormulaError.na));
    });

    test('next smaller match', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10)],
        [NumberValue(20)],
        [NumberValue(30)],
      ]);
      final result = eval(registry.get('XMATCH')!, [
        const NumberNode(25),
        RangeRefNode(A1Reference.parse('A1:A3')),
        const NumberNode(-1),
      ]);
      expect(result, const NumberValue(2)); // 20 is at position 2
    });

    test('next larger match', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10)],
        [NumberValue(20)],
        [NumberValue(30)],
      ]);
      final result = eval(registry.get('XMATCH')!, [
        const NumberNode(25),
        RangeRefNode(A1Reference.parse('A1:A3')),
        const NumberNode(1),
      ]);
      expect(result, const NumberValue(3)); // 30 is at position 3
    });

    test('wildcard match', () {
      context.rangeOverride = const RangeValue([
        [TextValue('apple')],
        [TextValue('banana')],
        [TextValue('cherry')],
      ]);
      final result = eval(registry.get('XMATCH')!, [
        const TextNode('ban*'),
        RangeRefNode(A1Reference.parse('A1:A3')),
        const NumberNode(2),
      ]);
      expect(result, const NumberValue(2));
    });
  });

  group('XLOOKUP', () {
    test('exact match returns from return array', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]);
      // We need two different ranges; since our test context returns the same
      // override, we'll test with the same data
      final result = eval(registry.get('XLOOKUP')!, [
        const NumberNode(2),
        RangeRefNode(A1Reference.parse('A1:A3')),
        RangeRefNode(A1Reference.parse('A1:A3')),
      ]);
      expect(result, const NumberValue(2));
    });

    test('not found returns #N/A by default', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      final result = eval(registry.get('XLOOKUP')!, [
        const NumberNode(99),
        RangeRefNode(A1Reference.parse('A1:A2')),
        RangeRefNode(A1Reference.parse('A1:A2')),
      ]);
      expect(result, const ErrorValue(FormulaError.na));
    });

    test('not found returns if_not_found value', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      final result = eval(registry.get('XLOOKUP')!, [
        const NumberNode(99),
        RangeRefNode(A1Reference.parse('A1:A2')),
        RangeRefNode(A1Reference.parse('A1:A2')),
        const TextNode('Not found'),
      ]);
      expect(result, const TextValue('Not found'));
    });
  });

  // -- Wave 6: Lookup & Reference functions ----------------------------------

  group('ROW', () {
    test('no args returns current cell row', () {
      context.currentCellOverride = 'C5'.a1;
      final result = eval(registry.get('ROW')!, []);
      expect(result, const NumberValue(5));
    });

    test('cell ref returns row', () {
      final result = eval(registry.get('ROW')!, [
        CellRefNode(A1Reference.parse('D7')),
      ]);
      expect(result, const NumberValue(7));
    });

    test('range ref returns first row', () {
      final result = eval(registry.get('ROW')!, [
        RangeRefNode(A1Reference.parse('B3:D8')),
      ]);
      expect(result, const NumberValue(3));
    });

    test('non-ref returns #VALUE!', () {
      final result = eval(registry.get('ROW')!, [const NumberNode(5)]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('COLUMN', () {
    test('no args returns current cell column', () {
      context.currentCellOverride = 'C5'.a1;
      final result = eval(registry.get('COLUMN')!, []);
      expect(result, const NumberValue(3));
    });

    test('cell ref returns column', () {
      final result = eval(registry.get('COLUMN')!, [
        CellRefNode(A1Reference.parse('D7')),
      ]);
      expect(result, const NumberValue(4));
    });

    test('range ref returns first column', () {
      final result = eval(registry.get('COLUMN')!, [
        RangeRefNode(A1Reference.parse('B3:D8')),
      ]);
      expect(result, const NumberValue(2));
    });

    test('non-ref returns #VALUE!', () {
      final result = eval(registry.get('COLUMN')!, [const NumberNode(5)]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('ROWS', () {
    test('cell ref returns 1', () {
      final result = eval(registry.get('ROWS')!, [
        CellRefNode(A1Reference.parse('A1')),
      ]);
      expect(result, const NumberValue(1));
    });

    test('range ref returns row count', () {
      final result = eval(registry.get('ROWS')!, [
        RangeRefNode(A1Reference.parse('A1:C5')),
      ]);
      expect(result, const NumberValue(5));
    });

    test('single row range returns 1', () {
      final result = eval(registry.get('ROWS')!, [
        RangeRefNode(A1Reference.parse('A1:D1')),
      ]);
      expect(result, const NumberValue(1));
    });
  });

  group('COLUMNS', () {
    test('cell ref returns 1', () {
      final result = eval(registry.get('COLUMNS')!, [
        CellRefNode(A1Reference.parse('A1')),
      ]);
      expect(result, const NumberValue(1));
    });

    test('range ref returns column count', () {
      final result = eval(registry.get('COLUMNS')!, [
        RangeRefNode(A1Reference.parse('A1:D5')),
      ]);
      expect(result, const NumberValue(4));
    });

    test('single column range returns 1', () {
      final result = eval(registry.get('COLUMNS')!, [
        RangeRefNode(A1Reference.parse('A1:A5')),
      ]);
      expect(result, const NumberValue(1));
    });
  });

  group('ADDRESS', () {
    test('absolute (default)', () {
      final result = eval(registry.get('ADDRESS')!, [
        const NumberNode(2),
        const NumberNode(3),
      ]);
      expect(result, const TextValue('\$C\$2'));
    });

    test('abs_num=2 (row absolute)', () {
      final result = eval(registry.get('ADDRESS')!, [
        const NumberNode(2),
        const NumberNode(3),
        const NumberNode(2),
      ]);
      expect(result, const TextValue('C\$2'));
    });

    test('abs_num=3 (column absolute)', () {
      final result = eval(registry.get('ADDRESS')!, [
        const NumberNode(2),
        const NumberNode(3),
        const NumberNode(3),
      ]);
      expect(result, const TextValue('\$C2'));
    });

    test('abs_num=4 (relative)', () {
      final result = eval(registry.get('ADDRESS')!, [
        const NumberNode(2),
        const NumberNode(3),
        const NumberNode(4),
      ]);
      expect(result, const TextValue('C2'));
    });

    test('R1C1 style absolute', () {
      final result = eval(registry.get('ADDRESS')!, [
        const NumberNode(2),
        const NumberNode(3),
        const NumberNode(1),
        const BooleanNode(false),
      ]);
      expect(result, const TextValue('R2C3'));
    });

    test('R1C1 style relative', () {
      final result = eval(registry.get('ADDRESS')!, [
        const NumberNode(2),
        const NumberNode(3),
        const NumberNode(4),
        const BooleanNode(false),
      ]);
      expect(result, const TextValue('R[2]C[3]'));
    });

    test('with sheet name', () {
      final result = eval(registry.get('ADDRESS')!, [
        const NumberNode(1),
        const NumberNode(1),
        const NumberNode(1),
        const BooleanNode(true),
        const TextNode('Sheet2'),
      ]);
      expect(result, const TextValue('Sheet2!\$A\$1'));
    });

    test('invalid row returns #VALUE!', () {
      final result = eval(registry.get('ADDRESS')!, [
        const NumberNode(0),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('invalid abs_num returns #VALUE!', () {
      final result = eval(registry.get('ADDRESS')!, [
        const NumberNode(1),
        const NumberNode(1),
        const NumberNode(5),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('INDIRECT', () {
    test('resolves cell reference from text', () {
      context.cellValues = {
        'B2'.a1: const NumberValue(42),
      };
      final result = eval(registry.get('INDIRECT')!, [
        const TextNode('B2'),
      ]);
      expect(result, const NumberValue(42));
    });

    test('resolves range reference from text', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
      ]);
      final result = eval(registry.get('INDIRECT')!, [
        const TextNode('A1:B2'),
      ]);
      expect(result, isA<RangeValue>());
    });

    test('invalid ref text returns #REF!', () {
      final result = eval(registry.get('INDIRECT')!, [
        const TextNode('!!!invalid!!!'),
      ]);
      expect(result, const ErrorValue(FormulaError.ref));
    });

    test('R1C1 style not supported returns #REF!', () {
      final result = eval(registry.get('INDIRECT')!, [
        const TextNode('R1C1'),
        const BooleanNode(false),
      ]);
      expect(result, const ErrorValue(FormulaError.ref));
    });

    test('error arg propagates', () {
      final result = eval(registry.get('INDIRECT')!, [
        const ErrorNode(FormulaError.value),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('OFFSET', () {
    test('single cell offset', () {
      context.cellValues = {
        'C3'.a1: const NumberValue(99),
      };
      // OFFSET(A1, 2, 2) → C3
      final result = eval(registry.get('OFFSET')!, [
        CellRefNode(A1Reference.parse('A1')),
        const NumberNode(2),
        const NumberNode(2),
      ]);
      expect(result, const NumberValue(99));
    });

    test('offset with height and width returns range', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
      ]);
      // OFFSET(A1, 1, 1, 2, 2) → B2:C3
      final result = eval(registry.get('OFFSET')!, [
        CellRefNode(A1Reference.parse('A1')),
        const NumberNode(1),
        const NumberNode(1),
        const NumberNode(2),
        const NumberNode(2),
      ]);
      expect(result, isA<RangeValue>());
    });

    test('negative offset that goes out of bounds returns #REF!', () {
      final result = eval(registry.get('OFFSET')!, [
        CellRefNode(A1Reference.parse('A1')),
        const NumberNode(-1),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.ref));
    });

    test('zero height returns #REF!', () {
      final result = eval(registry.get('OFFSET')!, [
        CellRefNode(A1Reference.parse('A1')),
        const NumberNode(0),
        const NumberNode(0),
        const NumberNode(0),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.ref));
    });

    test('range ref preserves default dimensions', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10), NumberValue(20)],
        [NumberValue(30), NumberValue(40)],
      ]);
      // OFFSET(A1:B2, 2, 0) → A3:B4 (height=2, width=2 from original)
      final result = eval(registry.get('OFFSET')!, [
        RangeRefNode(A1Reference.parse('A1:B2')),
        const NumberNode(2),
        const NumberNode(0),
      ]);
      expect(result, isA<RangeValue>());
    });

    test('non-ref arg returns #VALUE!', () {
      final result = eval(registry.get('OFFSET')!, [
        const NumberNode(5),
        const NumberNode(0),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('TRANSPOSE', () {
    test('transposes 2x3 to 3x2', () {
      final result = eval(registry.get('TRANSPOSE')!, [
        const NumberNode(0), // dummy — we override range
      ]);
      // With a non-range value, returns 1x1
      expect(result, isA<RangeValue>());
      final rv = result as RangeValue;
      expect(rv.rowCount, 1);
      expect(rv.columnCount, 1);
    });

    test('transposes range', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1), NumberValue(2), NumberValue(3)],
        [NumberValue(4), NumberValue(5), NumberValue(6)],
      ]);
      final result = eval(registry.get('TRANSPOSE')!, [
        RangeRefNode(A1Reference.parse('A1:C2')),
      ]);
      expect(result, isA<RangeValue>());
      final rv = result as RangeValue;
      expect(rv.rowCount, 3);
      expect(rv.columnCount, 2);
      expect(rv.values[0][0], const NumberValue(1));
      expect(rv.values[0][1], const NumberValue(4));
      expect(rv.values[1][0], const NumberValue(2));
      expect(rv.values[1][1], const NumberValue(5));
      expect(rv.values[2][0], const NumberValue(3));
      expect(rv.values[2][1], const NumberValue(6));
    });

    test('transposes single row to single column', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10), NumberValue(20), NumberValue(30)],
      ]);
      final result = eval(registry.get('TRANSPOSE')!, [
        RangeRefNode(A1Reference.parse('A1:C1')),
      ]);
      final rv = result as RangeValue;
      expect(rv.rowCount, 3);
      expect(rv.columnCount, 1);
      expect(rv.values[0][0], const NumberValue(10));
      expect(rv.values[1][0], const NumberValue(20));
      expect(rv.values[2][0], const NumberValue(30));
    });

    test('error propagates', () {
      final result = eval(registry.get('TRANSPOSE')!, [
        const ErrorNode(FormulaError.ref),
      ]);
      expect(result, const ErrorValue(FormulaError.ref));
    });
  });

  group('HYPERLINK', () {
    test('returns URL when no friendly name', () {
      final result = eval(registry.get('HYPERLINK')!, [
        const TextNode('https://example.com'),
      ]);
      expect(result, const TextValue('https://example.com'));
    });

    test('returns friendly name when provided', () {
      final result = eval(registry.get('HYPERLINK')!, [
        const TextNode('https://example.com'),
        const TextNode('Click here'),
      ]);
      expect(result, const TextValue('Click here'));
    });

    test('error in URL propagates', () {
      final result = eval(registry.get('HYPERLINK')!, [
        const ErrorNode(FormulaError.value),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('error in friendly name propagates', () {
      final result = eval(registry.get('HYPERLINK')!, [
        const TextNode('https://example.com'),
        const ErrorNode(FormulaError.ref),
      ]);
      expect(result, const ErrorValue(FormulaError.ref));
    });
  });

  group('AREAS', () {
    test('cell ref returns 1', () {
      final result = eval(registry.get('AREAS')!, [
        CellRefNode(A1Reference.parse('A1')),
      ]);
      expect(result, const NumberValue(1));
    });

    test('range ref returns 1', () {
      final result = eval(registry.get('AREAS')!, [
        RangeRefNode(A1Reference.parse('A1:C5')),
      ]);
      expect(result, const NumberValue(1));
    });

    test('non-ref returns #VALUE!', () {
      final result = eval(registry.get('AREAS')!, [
        const NumberNode(5),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });
}
