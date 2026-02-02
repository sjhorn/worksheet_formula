import 'dart:math' as math;

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

  group('SUMIFS', () {
    test('sums with single criteria pair', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10)],
        [NumberValue(20)],
        [NumberValue(30)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A3'));
      final result = eval(registry.get('SUMIFS')!, [
        rangeNode,
        rangeNode,
        const TextNode('>10'),
      ]);
      expect(result, const NumberValue(50));
    });

    test('returns 0 when no matches', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A2'));
      final result = eval(registry.get('SUMIFS')!, [
        rangeNode,
        rangeNode,
        const TextNode('>100'),
      ]);
      expect(result, const NumberValue(0));
    });
  });

  group('COUNTIFS', () {
    test('counts with single criteria pair', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(5)],
        [NumberValue(10)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A3'));
      final result = eval(registry.get('COUNTIFS')!, [
        rangeNode,
        const TextNode('>2'),
      ]);
      expect(result, const NumberValue(2));
    });
  });

  group('AVERAGEIFS', () {
    test('averages with criteria', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10)],
        [NumberValue(20)],
        [NumberValue(30)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A3'));
      final result = eval(registry.get('AVERAGEIFS')!, [
        rangeNode,
        rangeNode,
        const TextNode('>10'),
      ]);
      expect(result, const NumberValue(25));
    });

    test('returns #DIV/0! when no matches', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A1'));
      final result = eval(registry.get('AVERAGEIFS')!, [
        rangeNode,
        rangeNode,
        const TextNode('>100'),
      ]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });

  group('MEDIAN', () {
    test('returns middle value for odd count', () {
      final result = eval(registry.get('MEDIAN')!, [
        const NumberNode(3),
        const NumberNode(1),
        const NumberNode(2),
      ]);
      expect(result, const NumberValue(2));
    });

    test('returns average of middle two for even count', () {
      final result = eval(registry.get('MEDIAN')!, [
        const NumberNode(1),
        const NumberNode(2),
        const NumberNode(3),
        const NumberNode(4),
      ]);
      expect(result, const NumberValue(2.5));
    });

    test('single value returns itself', () {
      final result = eval(registry.get('MEDIAN')!, [const NumberNode(5)]);
      expect(result, const NumberValue(5));
    });

    test('no numbers returns #NUM!', () {
      final result =
          eval(registry.get('MEDIAN')!, [const TextNode('abc')]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('MODE.SNGL', () {
    test('returns most frequent value', () {
      final result = eval(registry.get('MODE.SNGL')!, [
        const NumberNode(1),
        const NumberNode(2),
        const NumberNode(2),
        const NumberNode(3),
      ]);
      expect(result, const NumberValue(2));
    });

    test('returns #N/A when all unique', () {
      final result = eval(registry.get('MODE.SNGL')!, [
        const NumberNode(1),
        const NumberNode(2),
        const NumberNode(3),
      ]);
      expect(result, const ErrorValue(FormulaError.na));
    });

    test('MODE alias works', () {
      final result = eval(registry.get('MODE')!, [
        const NumberNode(5),
        const NumberNode(5),
        const NumberNode(3),
      ]);
      expect(result, const NumberValue(5));
    });
  });

  group('LARGE', () {
    test('returns k-th largest', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(3)],
        [NumberValue(1)],
        [NumberValue(4)],
        [NumberValue(1)],
        [NumberValue(5)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A5'));
      final result = eval(registry.get('LARGE')!, [
        rangeNode,
        const NumberNode(2),
      ]);
      expect(result, const NumberValue(4));
    });

    test('k out of range returns #NUM!', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A2'));
      final result = eval(registry.get('LARGE')!, [
        rangeNode,
        const NumberNode(5),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('SMALL', () {
    test('returns k-th smallest', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(3)],
        [NumberValue(1)],
        [NumberValue(4)],
        [NumberValue(1)],
        [NumberValue(5)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A5'));
      final result = eval(registry.get('SMALL')!, [
        rangeNode,
        const NumberNode(2),
      ]);
      expect(result, const NumberValue(1));
    });

    test('k out of range returns #NUM!', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A1'));
      final result = eval(registry.get('SMALL')!, [
        rangeNode,
        const NumberNode(5),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('RANK.EQ', () {
    test('descending rank (default)', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(3)],
        [NumberValue(1)],
        [NumberValue(4)],
        [NumberValue(1)],
        [NumberValue(5)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A5'));
      // 5 is the largest, rank 1
      final result = eval(registry.get('RANK.EQ')!, [
        const NumberNode(5),
        rangeNode,
      ]);
      expect(result, const NumberValue(1));
    });

    test('ascending rank', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(3)],
        [NumberValue(1)],
        [NumberValue(4)],
        [NumberValue(5)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A4'));
      // 1 is the smallest, rank 1
      final result = eval(registry.get('RANK.EQ')!, [
        const NumberNode(1),
        rangeNode,
        const NumberNode(1),
      ]);
      expect(result, const NumberValue(1));
    });

    test('not found returns #N/A', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A2'));
      final result = eval(registry.get('RANK.EQ')!, [
        const NumberNode(99),
        rangeNode,
      ]);
      expect(result, const ErrorValue(FormulaError.na));
    });

    test('RANK alias works', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]);
      final rangeNode = RangeRefNode(A1Reference.parse('A1:A3'));
      final result = eval(registry.get('RANK')!, [
        const NumberNode(3),
        rangeNode,
      ]);
      expect(result, const NumberValue(1));
    });
  });

  group('STDEV.S', () {
    test('sample standard deviation', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(4)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(5)],
        [NumberValue(7)],
        [NumberValue(9)],
      ]);
      final result = eval(registry.get('STDEV.S')!, [
        RangeRefNode(A1Reference.parse('A1:A8')),
      ]);
      expect((result as NumberValue).value, closeTo(2.138, 0.01));
    });

    test('less than 2 values returns #DIV/0!', () {
      final result = eval(registry.get('STDEV.S')!, [const NumberNode(5)]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });
  });

  group('STDEV.P', () {
    test('population standard deviation', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(4)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(5)],
        [NumberValue(7)],
        [NumberValue(9)],
      ]);
      final result = eval(registry.get('STDEV.P')!, [
        RangeRefNode(A1Reference.parse('A1:A8')),
      ]);
      expect((result as NumberValue).value, closeTo(math.sqrt(4.0), 0.01));
    });
  });

  group('VAR.S', () {
    test('sample variance', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(4)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(5)],
        [NumberValue(7)],
        [NumberValue(9)],
      ]);
      final result = eval(registry.get('VAR.S')!, [
        RangeRefNode(A1Reference.parse('A1:A8')),
      ]);
      expect((result as NumberValue).value, closeTo(4.571, 0.01));
    });
  });

  group('VAR.P', () {
    test('population variance', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(2)],
        [NumberValue(4)],
        [NumberValue(4)],
        [NumberValue(4)],
        [NumberValue(5)],
        [NumberValue(5)],
        [NumberValue(7)],
        [NumberValue(9)],
      ]);
      final result = eval(registry.get('VAR.P')!, [
        RangeRefNode(A1Reference.parse('A1:A8')),
      ]);
      expect((result as NumberValue).value, closeTo(4.0, 0.01));
    });
  });

  group('PERCENTILE.INC', () {
    test('returns median at k=0.5', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
      ]);
      final result = eval(registry.get('PERCENTILE.INC')!, [
        RangeRefNode(A1Reference.parse('A1:A4')),
        const NumberNode(0.5),
      ]);
      expect(result, const NumberValue(2.5));
    });

    test('k=0 returns min', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]);
      final result = eval(registry.get('PERCENTILE.INC')!, [
        RangeRefNode(A1Reference.parse('A1:A3')),
        const NumberNode(0),
      ]);
      expect(result, const NumberValue(1));
    });

    test('k=1 returns max', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]);
      final result = eval(registry.get('PERCENTILE.INC')!, [
        RangeRefNode(A1Reference.parse('A1:A3')),
        const NumberNode(1),
      ]);
      expect(result, const NumberValue(3));
    });

    test('k out of range returns #NUM!', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
      ]);
      final result = eval(registry.get('PERCENTILE.INC')!, [
        RangeRefNode(A1Reference.parse('A1:A1')),
        const NumberNode(1.5),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('PERCENTILE.EXC', () {
    test('returns value at k', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
      ]);
      final result = eval(registry.get('PERCENTILE.EXC')!, [
        RangeRefNode(A1Reference.parse('A1:A4')),
        const NumberNode(0.5),
      ]);
      expect(result, const NumberValue(2.5));
    });

    test('k=0 returns #NUM!', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
      ]);
      final result = eval(registry.get('PERCENTILE.EXC')!, [
        RangeRefNode(A1Reference.parse('A1:A1')),
        const NumberNode(0),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('PERCENTRANK.INC', () {
    test('returns percent rank', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
      ]);
      final result = eval(registry.get('PERCENTRANK.INC')!, [
        RangeRefNode(A1Reference.parse('A1:A4')),
        const NumberNode(3),
      ]);
      expect((result as NumberValue).value, closeTo(0.666, 0.001));
    });

    test('value outside range returns #N/A', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      final result = eval(registry.get('PERCENTRANK.INC')!, [
        RangeRefNode(A1Reference.parse('A1:A2')),
        const NumberNode(5),
      ]);
      expect(result, const ErrorValue(FormulaError.na));
    });
  });

  group('PERCENTRANK.EXC', () {
    test('returns exclusive percent rank', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
      ]);
      final result = eval(registry.get('PERCENTRANK.EXC')!, [
        RangeRefNode(A1Reference.parse('A1:A4')),
        const NumberNode(3),
      ]);
      // rank = (2+1)/(4+1) = 0.6
      expect((result as NumberValue).value, closeTo(0.6, 0.001));
    });
  });

  group('RANK.AVG', () {
    test('no ties same as RANK.EQ', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
      ]);
      final result = eval(registry.get('RANK.AVG')!, [
        const NumberNode(2),
        RangeRefNode(A1Reference.parse('A1:A3')),
      ]);
      expect(result, const NumberValue(2));
    });

    test('ties are averaged', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(3)],
        [NumberValue(3)],
        [NumberValue(1)],
      ]);
      final result = eval(registry.get('RANK.AVG')!, [
        const NumberNode(3),
        RangeRefNode(A1Reference.parse('A1:A3')),
      ]);
      // descending: rank 1 and 2 for the two 3s -> avg = 1.5
      expect(result, const NumberValue(1.5));
    });
  });

  group('FREQUENCY', () {
    test('distributes into bins', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(5)],
      ]);
      // Use same override for both args
      final result = eval(registry.get('FREQUENCY')!, [
        RangeRefNode(A1Reference.parse('A1:A5')),
        RangeRefNode(A1Reference.parse('A1:A5')),
      ]);
      expect(result, isA<RangeValue>());
      final range = result as RangeValue;
      // bins = [1,2,3,4,5], so 6 entries: <=1, <=2, <=3, <=4, <=5, >5
      expect(range.rowCount, 6);
    });
  });

  group('AVEDEV', () {
    test('average absolute deviation', () {
      final result = eval(registry.get('AVEDEV')!, [
        const NumberNode(2),
        const NumberNode(4),
        const NumberNode(8),
        const NumberNode(16),
      ]);
      // mean = 7.5, deviations: 5.5, 3.5, 0.5, 8.5, avg = 4.5
      expect((result as NumberValue).value, closeTo(4.5, 0.001));
    });
  });

  group('AVERAGEA', () {
    test('includes booleans as 1/0', () {
      final result = eval(registry.get('AVERAGEA')!, [
        const NumberNode(1),
        const BooleanNode(true),
        const BooleanNode(false),
      ]);
      // (1 + 1 + 0) / 3 = 0.666...
      expect((result as NumberValue).value, closeTo(0.666, 0.01));
    });
  });

  group('MAXA', () {
    test('includes booleans', () {
      final result = eval(registry.get('MAXA')!, [
        const NumberNode(-5),
        const BooleanNode(true),
        const BooleanNode(false),
      ]);
      // max of -5, 1, 0 = 1
      expect(result, const NumberValue(1));
    });
  });

  group('MINA', () {
    test('includes booleans', () {
      final result = eval(registry.get('MINA')!, [
        const NumberNode(5),
        const BooleanNode(true),
        const BooleanNode(false),
      ]);
      // min of 5, 1, 0 = 0
      expect(result, const NumberValue(0));
    });
  });

  group('TRIMMEAN', () {
    test('trims outliers and averages', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
        [NumberValue(3)],
        [NumberValue(4)],
        [NumberValue(100)],
      ]);
      // 20% trim: trim 1 from each end (floor(5 * 0.2 / 2) = 0)
      // Actually floor(5*0.4/2) = floor(1) = 1 from each end
      final result = eval(registry.get('TRIMMEAN')!, [
        RangeRefNode(A1Reference.parse('A1:A5')),
        const NumberNode(0.4),
      ]);
      // trimmed: [2, 3, 4], mean = 3
      expect(result, const NumberValue(3));
    });

    test('percent >= 1 returns #NUM!', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
      ]);
      final result = eval(registry.get('TRIMMEAN')!, [
        RangeRefNode(A1Reference.parse('A1:A1')),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('GEOMEAN', () {
    test('geometric mean', () {
      final result = eval(registry.get('GEOMEAN')!, [
        const NumberNode(4),
        const NumberNode(9),
      ]);
      expect((result as NumberValue).value, closeTo(6, 0.001));
    });

    test('negative value returns #NUM!', () {
      final result = eval(registry.get('GEOMEAN')!, [
        const NumberNode(4),
        const NumberNode(-1),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('HARMEAN', () {
    test('harmonic mean', () {
      final result = eval(registry.get('HARMEAN')!, [
        const NumberNode(1),
        const NumberNode(2),
        const NumberNode(4),
      ]);
      // 3 / (1 + 0.5 + 0.25) = 3 / 1.75 = 1.714...
      expect((result as NumberValue).value, closeTo(1.714, 0.01));
    });

    test('zero returns #NUM!', () {
      final result = eval(registry.get('HARMEAN')!, [
        const NumberNode(0),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('MAXIFS', () {
    test('returns max of matching values', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10)],
        [NumberValue(20)],
        [NumberValue(30)],
      ]);
      final result = eval(registry.get('MAXIFS')!, [
        RangeRefNode(A1Reference.parse('A1:A3')),
        RangeRefNode(A1Reference.parse('A1:A3')),
        const TextNode('>5'),
      ]);
      expect(result, const NumberValue(30));
    });

    test('no matches returns 0', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(1)],
        [NumberValue(2)],
      ]);
      final result = eval(registry.get('MAXIFS')!, [
        RangeRefNode(A1Reference.parse('A1:A2')),
        RangeRefNode(A1Reference.parse('A1:A2')),
        const TextNode('>100'),
      ]);
      expect(result, const NumberValue(0));
    });
  });

  group('MINIFS', () {
    test('returns min of matching values', () {
      context.rangeOverride = const RangeValue([
        [NumberValue(10)],
        [NumberValue(20)],
        [NumberValue(30)],
      ]);
      final result = eval(registry.get('MINIFS')!, [
        RangeRefNode(A1Reference.parse('A1:A3')),
        RangeRefNode(A1Reference.parse('A1:A3')),
        const TextNode('>5'),
      ]);
      expect(result, const NumberValue(10));
    });
  });
}
