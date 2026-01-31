import 'package:a1/a1.dart';
import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/functions/function.dart';
import 'package:worksheet_formula/src/functions/registry.dart';
import 'package:worksheet_formula/src/functions/statistical.dart';

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
    registerStatisticalFunctions(registry);
    context = _TestContext(registry);
  });

  FormulaValue eval(FormulaFunction fn, List<FormulaNode> args) =>
      fn.call(args, context);

  group('COUNT', () {
    test('counts numeric values', () {
      final result = eval(registry.get('COUNT')!, [
        const NumberNode(1),
        const NumberNode(2),
        const NumberNode(3),
      ]);
      expect(result, const NumberValue(3));
    });

    test('skips non-numeric values', () {
      final result = eval(registry.get('COUNT')!, [
        const NumberNode(1),
        const TextNode('text'),
        const NumberNode(3),
      ]);
      expect(result, const NumberValue(2));
    });

    test('counts numbers in range', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10), TextValue('skip')],
        [NumberValue(20), EmptyValue()],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:B2'));
      final result = eval(registry.get('COUNT')!, [rangeNode]);
      expect(result, const NumberValue(2));
    });

    test('returns 0 for no numeric values', () {
      final result = eval(registry.get('COUNT')!, [const TextNode('a')]);
      expect(result, const NumberValue(0));
    });
  });

  group('COUNTA', () {
    test('counts non-empty values', () {
      final result = eval(registry.get('COUNTA')!, [
        const NumberNode(1),
        const TextNode('text'),
        const BooleanNode(true),
      ]);
      expect(result, const NumberValue(3));
    });

    test('skips empty values in range', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10), EmptyValue()],
        [TextValue('hi'), EmptyValue()],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:B2'));
      final result = eval(registry.get('COUNTA')!, [rangeNode]);
      expect(result, const NumberValue(2));
    });

    test('returns 0 for all empty', () {
      context.rangeOverride = const RangeValue([
        [EmptyValue(), EmptyValue()],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:B1'));
      final result = eval(registry.get('COUNTA')!, [rangeNode]);
      expect(result, const NumberValue(0));
    });
  });

  group('COUNTBLANK', () {
    test('counts empty cells in range', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10), EmptyValue()],
        [EmptyValue(), TextValue('hi')],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:B2'));
      final result = eval(registry.get('COUNTBLANK')!, [rangeNode]);
      expect(result, const NumberValue(2));
    });

    test('returns 0 when no empty cells', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1), NumberValue(2)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:B1'));
      final result = eval(registry.get('COUNTBLANK')!, [rangeNode]);
      expect(result, const NumberValue(0));
    });
  });

  group('COUNTIF', () {
    test('counts cells greater than value', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(5)],
        [NumberValue(10)],
        [NumberValue(15)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A4'));
      final result =
          eval(registry.get('COUNTIF')!, [rangeNode, const TextNode('>5')]);
      expect(result, const NumberValue(2));
    });

    test('counts cells less than value', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(5)],
        [NumberValue(10)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A3'));
      final result =
          eval(registry.get('COUNTIF')!, [rangeNode, const TextNode('<10')]);
      expect(result, const NumberValue(2));
    });

    test('counts cells equal to value', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(5)],
        [NumberValue(10)],
        [NumberValue(5)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A3'));
      final result =
          eval(registry.get('COUNTIF')!, [rangeNode, const NumberNode(5)]);
      expect(result, const NumberValue(2));
    });

    test('counts text matches case-insensitively', () {
      context.rangeOverride = const RangeValue([
        [TextValue('Apple')],
        [TextValue('banana')],
        [TextValue('APPLE')],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A3'));
      final result = eval(
          registry.get('COUNTIF')!, [rangeNode, const TextNode('apple')]);
      expect(result, const NumberValue(2));
    });

    test('counts with not equal operator', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A3'));
      final result =
          eval(registry.get('COUNTIF')!, [rangeNode, const TextNode('<>2')]);
      expect(result, const NumberValue(2));
    });

    test('counts with >= operator', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(5)],
        [NumberValue(10)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A3'));
      final result =
          eval(registry.get('COUNTIF')!, [rangeNode, const TextNode('>=5')]);
      expect(result, const NumberValue(2));
    });

    test('counts with <= operator', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(5)],
        [NumberValue(10)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A3'));
      final result =
          eval(registry.get('COUNTIF')!, [rangeNode, const TextNode('<=5')]);
      expect(result, const NumberValue(2));
    });
  });

  group('SUMIF', () {
    test('sums cells matching criteria', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(5)],
        [NumberValue(10)],
        [NumberValue(15)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A4'));
      final result =
          eval(registry.get('SUMIF')!, [rangeNode, const TextNode('>5')]);
      expect(result, const NumberValue(25));
    });

    test('sums with exact match', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(5)],
        [NumberValue(10)],
        [NumberValue(5)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A3'));
      final result =
          eval(registry.get('SUMIF')!, [rangeNode, const NumberNode(5)]);
      expect(result, const NumberValue(10));
    });

    test('returns 0 when no matches', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A2'));
      final result =
          eval(registry.get('SUMIF')!, [rangeNode, const TextNode('>100')]);
      expect(result, const NumberValue(0));
    });
  });

  group('AVERAGEIF', () {
    test('averages cells matching criteria', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(5)],
        [NumberValue(10)],
        [NumberValue(15)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A4'));
      final result = eval(
          registry.get('AVERAGEIF')!, [rangeNode, const TextNode('>5')]);
      expect(result, const NumberValue(12.5));
    });

    test('returns #DIV/0! when no matches', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A2'));
      final result = eval(
          registry.get('AVERAGEIF')!, [rangeNode, const TextNode('>100')]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });
}
