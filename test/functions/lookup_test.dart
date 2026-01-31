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
}
