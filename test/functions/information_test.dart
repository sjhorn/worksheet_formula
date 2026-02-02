import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/functions/function.dart';
import 'package:worksheet_formula/src/functions/information.dart';
import 'package:worksheet_formula/src/functions/registry.dart';
import 'package:a1/a1.dart';

class _TestContext implements EvaluationContext {
  @override
  A1 get currentCell => 'A1'.a1;
  @override
  String? get currentSheet => null;
  @override
  bool get isCancelled => false;

  final FunctionRegistry _registry;
  _TestContext(this._registry);

  @override
  FormulaValue getCellValue(A1 cell) => const EmptyValue();
  @override
  FormulaValue getRangeValues(A1Range range) =>
      const FormulaValue.error(FormulaError.ref);
  @override
  FormulaFunction? getFunction(String name) => _registry.get(name);
}

void main() {
  late FunctionRegistry registry;
  late _TestContext context;

  FormulaValue eval(FormulaFunction fn, List<FormulaNode> args) =>
      fn.call(args, context);

  setUp(() {
    registry = FunctionRegistry(registerBuiltIns: false);
    registerInformationFunctions(registry);
    context = _TestContext(registry);
  });

  group('ISBLANK', () {
    test('returns true for empty value', () {
      final result = eval(registry.get('ISBLANK')!, [
        CellRefNode(A1Reference.parse('A1')),
      ]);
      expect(result, const BooleanValue(true));
    });

    test('returns false for number', () {
      final result = eval(registry.get('ISBLANK')!, [const NumberNode(5)]);
      expect(result, const BooleanValue(false));
    });

    test('returns false for text', () {
      final result =
          eval(registry.get('ISBLANK')!, [const TextNode('hello')]);
      expect(result, const BooleanValue(false));
    });
  });

  group('ISERROR', () {
    test('returns true for error', () {
      final result = eval(
          registry.get('ISERROR')!, [const ErrorNode(FormulaError.value)]);
      expect(result, const BooleanValue(true));
    });

    test('returns true for #N/A', () {
      final result =
          eval(registry.get('ISERROR')!, [const ErrorNode(FormulaError.na)]);
      expect(result, const BooleanValue(true));
    });

    test('returns false for number', () {
      final result = eval(registry.get('ISERROR')!, [const NumberNode(5)]);
      expect(result, const BooleanValue(false));
    });
  });

  group('ISNUMBER', () {
    test('returns true for number', () {
      final result = eval(registry.get('ISNUMBER')!, [const NumberNode(42)]);
      expect(result, const BooleanValue(true));
    });

    test('returns false for text', () {
      final result =
          eval(registry.get('ISNUMBER')!, [const TextNode('hello')]);
      expect(result, const BooleanValue(false));
    });

    test('returns false for boolean', () {
      final result =
          eval(registry.get('ISNUMBER')!, [const BooleanNode(true)]);
      expect(result, const BooleanValue(false));
    });
  });

  group('ISTEXT', () {
    test('returns true for text', () {
      final result =
          eval(registry.get('ISTEXT')!, [const TextNode('hello')]);
      expect(result, const BooleanValue(true));
    });

    test('returns false for number', () {
      final result = eval(registry.get('ISTEXT')!, [const NumberNode(5)]);
      expect(result, const BooleanValue(false));
    });
  });

  group('ISLOGICAL', () {
    test('returns true for boolean', () {
      final result =
          eval(registry.get('ISLOGICAL')!, [const BooleanNode(true)]);
      expect(result, const BooleanValue(true));
    });

    test('returns false for number', () {
      final result = eval(registry.get('ISLOGICAL')!, [const NumberNode(1)]);
      expect(result, const BooleanValue(false));
    });
  });

  group('ISNA', () {
    test('returns true for #N/A', () {
      final result =
          eval(registry.get('ISNA')!, [const ErrorNode(FormulaError.na)]);
      expect(result, const BooleanValue(true));
    });

    test('returns false for other errors', () {
      final result = eval(
          registry.get('ISNA')!, [const ErrorNode(FormulaError.value)]);
      expect(result, const BooleanValue(false));
    });

    test('returns false for number', () {
      final result = eval(registry.get('ISNA')!, [const NumberNode(5)]);
      expect(result, const BooleanValue(false));
    });
  });

  group('TYPE', () {
    test('returns 1 for number', () {
      final result = eval(registry.get('TYPE')!, [const NumberNode(42)]);
      expect(result, const NumberValue(1));
    });

    test('returns 2 for text', () {
      final result = eval(registry.get('TYPE')!, [const TextNode('hi')]);
      expect(result, const NumberValue(2));
    });

    test('returns 4 for boolean', () {
      final result = eval(registry.get('TYPE')!, [const BooleanNode(true)]);
      expect(result, const NumberValue(4));
    });

    test('returns 16 for error', () {
      final result =
          eval(registry.get('TYPE')!, [const ErrorNode(FormulaError.ref)]);
      expect(result, const NumberValue(16));
    });

    test('returns 1 for empty (treated as number)', () {
      final result = eval(registry.get('TYPE')!, [
        CellRefNode(A1Reference.parse('A1')),
      ]);
      expect(result, const NumberValue(1));
    });
  });
}
