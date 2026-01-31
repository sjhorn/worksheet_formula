import 'package:a1/a1.dart';
import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/functions/function.dart';
import 'package:worksheet_formula/src/functions/registry.dart';
import 'package:worksheet_formula/src/functions/text.dart';

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
    registerTextFunctions(registry);
    context = _TestContext(registry);
  });

  FormulaValue eval(FormulaFunction fn, List<FormulaNode> args) =>
      fn.call(args, context);

  group('CONCAT', () {
    test('joins text strings', () {
      final result = eval(registry.get('CONCAT')!, [
        const TextNode('Hello'),
        const TextNode(' '),
        const TextNode('World'),
      ]);
      expect(result, const TextValue('Hello World'));
    });

    test('converts numbers to text', () {
      final result = eval(registry.get('CONCAT')!, [
        const TextNode('Value: '),
        const NumberNode(42),
      ]);
      expect(result, const TextValue('Value: 42'));
    });

    test('propagates error', () {
      final result = eval(registry.get('CONCAT')!, [
        const TextNode('a'),
        const ErrorNode(FormulaError.ref),
      ]);
      expect(result, const ErrorValue(FormulaError.ref));
    });
  });

  group('CONCATENATE', () {
    test('works the same as CONCAT', () {
      final result = eval(registry.get('CONCATENATE')!, [
        const TextNode('Hello'),
        const TextNode(' World'),
      ]);
      expect(result, const TextValue('Hello World'));
    });
  });

  group('LEFT', () {
    test('returns leftmost characters', () {
      final result = eval(registry.get('LEFT')!, [
        const TextNode('Hello'),
        const NumberNode(3),
      ]);
      expect(result, const TextValue('Hel'));
    });

    test('defaults to 1 character when no count', () {
      final result = eval(registry.get('LEFT')!, [
        const TextNode('Hello'),
      ]);
      expect(result, const TextValue('H'));
    });

    test('returns full string when count exceeds length', () {
      final result = eval(registry.get('LEFT')!, [
        const TextNode('Hi'),
        const NumberNode(10),
      ]);
      expect(result, const TextValue('Hi'));
    });

    test('negative count returns #VALUE!', () {
      final result = eval(registry.get('LEFT')!, [
        const TextNode('Hello'),
        const NumberNode(-1),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('RIGHT', () {
    test('returns rightmost characters', () {
      final result = eval(registry.get('RIGHT')!, [
        const TextNode('Hello'),
        const NumberNode(3),
      ]);
      expect(result, const TextValue('llo'));
    });

    test('defaults to 1 character when no count', () {
      final result = eval(registry.get('RIGHT')!, [
        const TextNode('Hello'),
      ]);
      expect(result, const TextValue('o'));
    });

    test('returns full string when count exceeds length', () {
      final result = eval(registry.get('RIGHT')!, [
        const TextNode('Hi'),
        const NumberNode(10),
      ]);
      expect(result, const TextValue('Hi'));
    });

    test('negative count returns #VALUE!', () {
      final result = eval(registry.get('RIGHT')!, [
        const TextNode('Hello'),
        const NumberNode(-1),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('MID', () {
    test('returns characters from middle', () {
      final result = eval(registry.get('MID')!, [
        const TextNode('Hello World'),
        const NumberNode(7),
        const NumberNode(5),
      ]);
      expect(result, const TextValue('World'));
    });

    test('start is 1-indexed', () {
      final result = eval(registry.get('MID')!, [
        const TextNode('ABCDE'),
        const NumberNode(1),
        const NumberNode(2),
      ]);
      expect(result, const TextValue('AB'));
    });

    test('returns empty when start exceeds length', () {
      final result = eval(registry.get('MID')!, [
        const TextNode('Hi'),
        const NumberNode(10),
        const NumberNode(3),
      ]);
      expect(result, const TextValue(''));
    });

    test('truncates when count exceeds remaining', () {
      final result = eval(registry.get('MID')!, [
        const TextNode('Hello'),
        const NumberNode(4),
        const NumberNode(10),
      ]);
      expect(result, const TextValue('lo'));
    });

    test('start < 1 returns #VALUE!', () {
      final result = eval(registry.get('MID')!, [
        const TextNode('Hello'),
        const NumberNode(0),
        const NumberNode(3),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('negative count returns #VALUE!', () {
      final result = eval(registry.get('MID')!, [
        const TextNode('Hello'),
        const NumberNode(1),
        const NumberNode(-1),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('non-numeric args return #VALUE!', () {
      final result = eval(registry.get('MID')!, [
        const TextNode('Hello'),
        const TextNode('a'),
        const NumberNode(3),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('LEN', () {
    test('returns length of text', () {
      final result = eval(registry.get('LEN')!, [const TextNode('Hello')]);
      expect(result, const NumberValue(5));
    });

    test('empty string returns 0', () {
      final result = eval(registry.get('LEN')!, [const TextNode('')]);
      expect(result, const NumberValue(0));
    });

    test('converts number to text then measures', () {
      final result = eval(registry.get('LEN')!, [const NumberNode(12345)]);
      expect(result, const NumberValue(5));
    });
  });

  group('LOWER', () {
    test('converts to lowercase', () {
      final result =
          eval(registry.get('LOWER')!, [const TextNode('Hello World')]);
      expect(result, const TextValue('hello world'));
    });

    test('already lowercase unchanged', () {
      final result = eval(registry.get('LOWER')!, [const TextNode('hello')]);
      expect(result, const TextValue('hello'));
    });
  });

  group('UPPER', () {
    test('converts to uppercase', () {
      final result =
          eval(registry.get('UPPER')!, [const TextNode('Hello World')]);
      expect(result, const TextValue('HELLO WORLD'));
    });

    test('already uppercase unchanged', () {
      final result = eval(registry.get('UPPER')!, [const TextNode('HELLO')]);
      expect(result, const TextValue('HELLO'));
    });
  });

  group('TRIM', () {
    test('removes leading and trailing spaces', () {
      final result =
          eval(registry.get('TRIM')!, [const TextNode('  Hello  ')]);
      expect(result, const TextValue('Hello'));
    });

    test('collapses multiple internal spaces', () {
      final result =
          eval(registry.get('TRIM')!, [const TextNode('Hello   World')]);
      expect(result, const TextValue('Hello World'));
    });

    test('handles both leading, trailing, and internal spaces', () {
      final result = eval(
          registry.get('TRIM')!, [const TextNode('  Hello   World  ')]);
      expect(result, const TextValue('Hello World'));
    });
  });

  group('TEXT', () {
    test('formats number as text', () {
      final result = eval(registry.get('TEXT')!, [
        const NumberNode(42),
        const TextNode('0'),
      ]);
      expect(result, const TextValue('42'));
    });

    test('non-numeric returns #VALUE!', () {
      final result = eval(registry.get('TEXT')!, [
        const TextNode('abc'),
        const TextNode('0'),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });
}
