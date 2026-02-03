import 'package:a1/a1.dart';
import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/ast/operators.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/functions/function.dart';

/// Simple mock context for testing AST node evaluation.
class MockContext implements EvaluationContext {
  final Map<String, FormulaValue> cells;
  final Map<String, FormulaFunction> functions;

  MockContext({
    this.cells = const {},
    this.functions = const {},
  });

  @override
  A1 get currentCell => 'A1'.a1;

  @override
  String? get currentSheet => null;

  @override
  FormulaValue getCellValue(A1 cell) {
    return cells[cell.toString()] ?? const EmptyValue();
  }

  @override
  FormulaValue getRangeValues(A1Range range) {
    final fromA1 = range.from.a1;
    final toA1 = range.to.a1;
    if (fromA1 == null || toA1 == null) {
      return const FormulaValue.error(FormulaError.ref);
    }
    final values = <List<FormulaValue>>[];
    for (var row = fromA1.row; row <= toA1.row; row++) {
      final rowValues = <FormulaValue>[];
      for (var col = fromA1.column; col <= toA1.column; col++) {
        final cell = A1.fromVector(col, row);
        rowValues.add(getCellValue(cell));
      }
      values.add(rowValues);
    }
    return FormulaValue.range(values);
  }

  @override
  FormulaFunction? getFunction(String name) =>
      functions[name.toUpperCase()];

  @override
  FormulaValue? getVariable(String name) => null;

  @override
  bool get isCancelled => false;
}

void main() {
  late MockContext context;

  setUp(() {
    context = MockContext(cells: {
      'A1': const NumberValue(10),
      'B1': const NumberValue(20),
      'A2': const TextValue('hello'),
    });
  });

  group('NumberNode', () {
    test('evaluates to NumberValue', () {
      const node = NumberNode(42);
      expect(node.evaluate(context), const NumberValue(42));
    });

    test('has no cell references', () {
      const node = NumberNode(42);
      expect(node.cellReferences, isEmpty);
    });

    test('toFormulaString', () {
      const node = NumberNode(42);
      expect(node.toFormulaString(), '42');
    });

    test('equality', () {
      expect(const NumberNode(42), const NumberNode(42));
      expect(const NumberNode(42), isNot(const NumberNode(43)));
    });
  });

  group('TextNode', () {
    test('evaluates to TextValue', () {
      const node = TextNode('hello');
      expect(node.evaluate(context), const TextValue('hello'));
    });

    test('has no cell references', () {
      const node = TextNode('hello');
      expect(node.cellReferences, isEmpty);
    });

    test('toFormulaString wraps in quotes', () {
      const node = TextNode('hello');
      expect(node.toFormulaString(), '"hello"');
    });

    test('toFormulaString escapes inner quotes', () {
      const node = TextNode('say "hi"');
      expect(node.toFormulaString(), '"say ""hi"""');
    });

    test('equality', () {
      expect(const TextNode('a'), const TextNode('a'));
      expect(const TextNode('a'), isNot(const TextNode('b')));
    });
  });

  group('BooleanNode', () {
    test('evaluates to BooleanValue', () {
      const nodeT = BooleanNode(true);
      const nodeF = BooleanNode(false);
      expect(nodeT.evaluate(context), const BooleanValue(true));
      expect(nodeF.evaluate(context), const BooleanValue(false));
    });

    test('has no cell references', () {
      expect(const BooleanNode(true).cellReferences, isEmpty);
    });

    test('toFormulaString', () {
      expect(const BooleanNode(true).toFormulaString(), 'TRUE');
      expect(const BooleanNode(false).toFormulaString(), 'FALSE');
    });

    test('equality', () {
      expect(const BooleanNode(true), const BooleanNode(true));
      expect(const BooleanNode(true), isNot(const BooleanNode(false)));
    });
  });

  group('ErrorNode', () {
    test('evaluates to ErrorValue', () {
      const node = ErrorNode(FormulaError.ref);
      expect(node.evaluate(context), const ErrorValue(FormulaError.ref));
    });

    test('has no cell references', () {
      expect(const ErrorNode(FormulaError.ref).cellReferences, isEmpty);
    });

    test('toFormulaString', () {
      expect(const ErrorNode(FormulaError.ref).toFormulaString(), '#REF!');
      expect(
        const ErrorNode(FormulaError.divZero).toFormulaString(),
        '#DIV/0!',
      );
    });
  });

  group('CellRefNode', () {
    test('evaluates by looking up cell value', () {
      final node = CellRefNode(A1Reference.parse('A1'));
      expect(node.evaluate(context), const NumberValue(10));
    });

    test('returns EmptyValue for empty cell', () {
      final node = CellRefNode(A1Reference.parse('Z99'));
      expect(node.evaluate(context), const EmptyValue());
    });

    test('has the cell in cellReferences', () {
      final ref = A1Reference.parse('A1');
      final node = CellRefNode(ref);
      expect(node.cellReferences, ['A1'.a1]);
    });

    test('toFormulaString without sheet', () {
      final node = CellRefNode(A1Reference.parse('A1'));
      expect(node.toFormulaString(), 'A1');
    });

    test('toFormulaString with sheet', () {
      final node = CellRefNode(A1Reference.parse('Sheet1!A1'));
      expect(node.toFormulaString(), contains('A1'));
    });

    test('equality', () {
      final ref = A1Reference.parse('A1');
      expect(CellRefNode(ref), CellRefNode(ref));
    });
  });

  group('RangeRefNode', () {
    test('evaluates to RangeValue', () {
      final node = RangeRefNode(A1Reference.parse('A1:B1'));
      final result = node.evaluate(context);
      expect(result, isA<RangeValue>());
      final rv = result as RangeValue;
      expect(rv.rowCount, 1);
      expect(rv.columnCount, 2);
    });

    test('has all cells in the range as cellReferences', () {
      final node = RangeRefNode(A1Reference.parse('A1:B2'));
      expect(node.cellReferences.length, 4);
    });

    test('toFormulaString', () {
      final node = RangeRefNode(A1Reference.parse('A1:B2'));
      expect(node.toFormulaString(), contains('A1'));
    });

    test('equality', () {
      final ref = A1Reference.parse('A1:B2');
      expect(RangeRefNode(ref), RangeRefNode(ref));
    });
  });

  group('BinaryOpNode', () {
    test('evaluates arithmetic', () {
      const node = BinaryOpNode(
        NumberNode(10),
        BinaryOperator.add,
        NumberNode(5),
      );
      expect(node.evaluate(context), const NumberValue(15));
    });

    test('short-circuits on left error', () {
      const node = BinaryOpNode(
        ErrorNode(FormulaError.ref),
        BinaryOperator.add,
        NumberNode(5),
      );
      expect(node.evaluate(context), const ErrorValue(FormulaError.ref));
    });

    test('short-circuits on right error', () {
      const node = BinaryOpNode(
        NumberNode(5),
        BinaryOperator.multiply,
        ErrorNode(FormulaError.na),
      );
      expect(node.evaluate(context), const ErrorValue(FormulaError.na));
    });

    test('collects cell references from both sides', () {
      final node = BinaryOpNode(
        CellRefNode(A1Reference.parse('A1')),
        BinaryOperator.add,
        CellRefNode(A1Reference.parse('B1')),
      );
      expect(node.cellReferences.length, 2);
    });

    test('toFormulaString', () {
      const node = BinaryOpNode(
        NumberNode(1),
        BinaryOperator.add,
        NumberNode(2),
      );
      expect(node.toFormulaString(), '1+2');
    });
  });

  group('UnaryOpNode', () {
    test('evaluates negate', () {
      const node = UnaryOpNode(UnaryOperator.negate, NumberNode(5));
      expect(node.evaluate(context), const NumberValue(-5));
    });

    test('evaluates percent', () {
      const node = UnaryOpNode(UnaryOperator.percent, NumberNode(50));
      expect(node.evaluate(context), const NumberValue(0.5));
    });

    test('propagates error', () {
      const node = UnaryOpNode(
        UnaryOperator.negate,
        ErrorNode(FormulaError.ref),
      );
      expect(node.evaluate(context), const ErrorValue(FormulaError.ref));
    });

    test('collects cell references from operand', () {
      final node = UnaryOpNode(
        UnaryOperator.negate,
        CellRefNode(A1Reference.parse('A1')),
      );
      expect(node.cellReferences.length, 1);
    });

    test('toFormulaString', () {
      const node = UnaryOpNode(UnaryOperator.negate, NumberNode(5));
      expect(node.toFormulaString(), '-5');
    });
  });

  group('FunctionCallNode', () {
    test('returns #NAME? for unknown function', () {
      const node = FunctionCallNode('UNKNOWN', [NumberNode(1)]);
      expect(node.evaluate(context), const ErrorValue(FormulaError.name));
    });

    test('collects cell references from arguments', () {
      final node = FunctionCallNode('SUM', [
        CellRefNode(A1Reference.parse('A1')),
        CellRefNode(A1Reference.parse('B1')),
      ]);
      expect(node.cellReferences.length, 2);
    });

    test('toFormulaString', () {
      const node = FunctionCallNode('SUM', [
        NumberNode(1),
        NumberNode(2),
      ]);
      expect(node.toFormulaString(), 'SUM(1,2)');
    });
  });

  group('ParenthesizedNode', () {
    test('evaluates to inner value', () {
      const node = ParenthesizedNode(NumberNode(42));
      expect(node.evaluate(context), const NumberValue(42));
    });

    test('collects cell references from inner', () {
      final node = ParenthesizedNode(
        CellRefNode(A1Reference.parse('A1')),
      );
      expect(node.cellReferences.length, 1);
    });

    test('toFormulaString wraps in parens', () {
      const node = ParenthesizedNode(NumberNode(42));
      expect(node.toFormulaString(), '(42)');
    });
  });

  group('NameNode', () {
    test('evaluates to #NAME? without variable scope', () {
      const node = NameNode('x');
      expect(node.evaluate(context), const ErrorValue(FormulaError.name));
    });

    test('has no cell references', () {
      const node = NameNode('count');
      expect(node.cellReferences, isEmpty);
    });

    test('toFormulaString returns the name', () {
      const node = NameNode('myVar');
      expect(node.toFormulaString(), 'myVar');
    });

    test('equality', () {
      expect(const NameNode('x'), const NameNode('x'));
      expect(const NameNode('x'), isNot(const NameNode('y')));
    });

    test('hashCode is consistent with equality', () {
      expect(const NameNode('x').hashCode, const NameNode('x').hashCode);
    });
  });

  group('FormulaNode sealed class', () {
    test('pattern matching works on all subtypes', () {
      final nodes = <FormulaNode>[
        const NumberNode(1),
        const TextNode('hi'),
        const BooleanNode(true),
        const ErrorNode(FormulaError.na),
        CellRefNode(A1Reference.parse('A1')),
        RangeRefNode(A1Reference.parse('A1:B1')),
        const BinaryOpNode(NumberNode(1), BinaryOperator.add, NumberNode(2)),
        const UnaryOpNode(UnaryOperator.negate, NumberNode(1)),
        const FunctionCallNode('SUM', [NumberNode(1)]),
        const CallExpressionNode(
            FunctionCallNode('LAMBDA', [NameNode('x'), NameNode('x')]),
            [NumberNode(5)]),
        const ParenthesizedNode(NumberNode(1)),
        const NameNode('x'),
      ];

      final types = nodes.map((n) => switch (n) {
        NumberNode() => 'number',
        TextNode() => 'text',
        BooleanNode() => 'boolean',
        ErrorNode() => 'error',
        CellRefNode() => 'cellRef',
        RangeRefNode() => 'rangeRef',
        BinaryOpNode() => 'binaryOp',
        UnaryOpNode() => 'unaryOp',
        FunctionCallNode() => 'functionCall',
        CallExpressionNode() => 'callExpression',
        ParenthesizedNode() => 'parenthesized',
        NameNode() => 'name',
      }).toList();

      expect(types, [
        'number', 'text', 'boolean', 'error', 'cellRef', 'rangeRef',
        'binaryOp', 'unaryOp', 'functionCall', 'callExpression',
        'parenthesized', 'name',
      ]);
    });
  });
}
