import 'package:a1/a1.dart';
import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/functions/function.dart';
import 'package:worksheet_formula/src/functions/logical.dart';
import 'package:worksheet_formula/src/functions/registry.dart';

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

  setUp(() {
    registry = FunctionRegistry(registerBuiltIns: false);
    registerLogicalFunctions(registry);
    context = _TestContext(registry);
  });

  FormulaValue eval(FormulaFunction fn, List<FormulaNode> args) =>
      fn.call(args, context);

  group('IF', () {
    test('returns true branch when truthy', () {
      final result = eval(registry.get('IF')!, [
        const BooleanNode(true),
        const TextNode('yes'),
        const TextNode('no'),
      ]);
      expect(result, const TextValue('yes'));
    });

    test('returns false branch when falsy', () {
      final result = eval(registry.get('IF')!, [
        const BooleanNode(false),
        const TextNode('yes'),
        const TextNode('no'),
      ]);
      expect(result, const TextValue('no'));
    });

    test('returns FALSE when no else branch and condition is false', () {
      final result = eval(registry.get('IF')!, [
        const BooleanNode(false),
        const TextNode('yes'),
      ]);
      expect(result, const BooleanValue(false));
    });

    test('propagates condition error', () {
      final result = eval(registry.get('IF')!, [
        const ErrorNode(FormulaError.ref),
        const TextNode('yes'),
        const TextNode('no'),
      ]);
      expect(result, const ErrorValue(FormulaError.ref));
    });
  });

  group('AND', () {
    test('returns true when all true', () {
      final result = eval(registry.get('AND')!, [
        const BooleanNode(true),
        const NumberNode(1),
      ]);
      expect(result, const BooleanValue(true));
    });

    test('returns false when any false', () {
      final result = eval(registry.get('AND')!, [
        const BooleanNode(true),
        const BooleanNode(false),
      ]);
      expect(result, const BooleanValue(false));
    });

    test('propagates error', () {
      final result = eval(registry.get('AND')!, [
        const ErrorNode(FormulaError.na),
      ]);
      expect(result, const ErrorValue(FormulaError.na));
    });
  });

  group('OR', () {
    test('returns true when any true', () {
      final result = eval(registry.get('OR')!, [
        const BooleanNode(false),
        const BooleanNode(true),
      ]);
      expect(result, const BooleanValue(true));
    });

    test('returns false when all false', () {
      final result = eval(registry.get('OR')!, [
        const BooleanNode(false),
        const NumberNode(0),
      ]);
      expect(result, const BooleanValue(false));
    });

    test('propagates error', () {
      final result = eval(registry.get('OR')!, [
        const ErrorNode(FormulaError.na),
      ]);
      expect(result, const ErrorValue(FormulaError.na));
    });
  });

  group('NOT', () {
    test('inverts true to false', () {
      final result = eval(registry.get('NOT')!, [const BooleanNode(true)]);
      expect(result, const BooleanValue(false));
    });

    test('inverts false to true', () {
      final result = eval(registry.get('NOT')!, [const BooleanNode(false)]);
      expect(result, const BooleanValue(true));
    });

    test('propagates error', () {
      final result =
          eval(registry.get('NOT')!, [const ErrorNode(FormulaError.ref)]);
      expect(result, const ErrorValue(FormulaError.ref));
    });
  });

  group('IFERROR', () {
    test('returns value when no error', () {
      final result = eval(registry.get('IFERROR')!, [
        const NumberNode(42),
        const TextNode('error'),
      ]);
      expect(result, const NumberValue(42));
    });

    test('returns fallback when error', () {
      final result = eval(registry.get('IFERROR')!, [
        const ErrorNode(FormulaError.divZero),
        const NumberNode(0),
      ]);
      expect(result, const NumberValue(0));
    });
  });

  group('IFNA', () {
    test('returns value when not #N/A', () {
      final result = eval(registry.get('IFNA')!, [
        const NumberNode(42),
        const TextNode('fallback'),
      ]);
      expect(result, const NumberValue(42));
    });

    test('returns fallback when #N/A', () {
      final result = eval(registry.get('IFNA')!, [
        const ErrorNode(FormulaError.na),
        const TextNode('fallback'),
      ]);
      expect(result, const TextValue('fallback'));
    });

    test('does not catch other errors', () {
      final result = eval(registry.get('IFNA')!, [
        const ErrorNode(FormulaError.ref),
        const TextNode('fallback'),
      ]);
      expect(result, const ErrorValue(FormulaError.ref));
    });
  });

  group('TRUE/FALSE functions', () {
    test('TRUE() returns true', () {
      final result = eval(registry.get('TRUE')!, []);
      expect(result, const BooleanValue(true));
    });

    test('FALSE() returns false', () {
      final result = eval(registry.get('FALSE')!, []);
      expect(result, const BooleanValue(false));
    });
  });
}
