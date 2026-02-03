import 'package:a1/a1.dart';
import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/functions/function.dart';
import 'package:worksheet_formula/src/functions/registry.dart';
import 'package:worksheet_formula/src/functions/web.dart';

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
    registerWebFunctions(registry);
    context = _TestContext(registry);
  });

  FormulaValue eval(FormulaFunction fn, List<FormulaNode> args) =>
      fn.call(args, context);

  group('ENCODEURL', () {
    test('basic encoding', () {
      final result = eval(registry.get('ENCODEURL')!, [
        const TextNode('hello world'),
      ]);
      expect(result, const TextValue('hello%20world'));
    });

    test('special characters', () {
      final result = eval(registry.get('ENCODEURL')!, [
        const TextNode('a=1&b=2'),
      ]);
      expect(result, const TextValue('a%3D1%26b%3D2'));
    });

    test('already-safe characters', () {
      final result = eval(registry.get('ENCODEURL')!, [
        const TextNode('hello'),
      ]);
      expect(result, const TextValue('hello'));
    });

    test('empty string', () {
      final result = eval(registry.get('ENCODEURL')!, [
        const TextNode(''),
      ]);
      expect(result, const TextValue(''));
    });

    test('URL with path', () {
      final result = eval(registry.get('ENCODEURL')!, [
        const TextNode('http://example.com/path?q=hello world'),
      ]);
      final text = (result as TextValue).value;
      expect(text.contains('%20'), isTrue);
    });

    test('unicode characters', () {
      final result = eval(registry.get('ENCODEURL')!, [
        const TextNode('café'),
      ]);
      final text = (result as TextValue).value;
      expect(text.contains('%'), isTrue); // é should be encoded
    });
  });

  group('REGEXMATCH', () {
    test('match found', () {
      final result = eval(registry.get('REGEXMATCH')!, [
        const TextNode('Hello World'),
        const TextNode('World'),
      ]);
      expect(result, const BooleanValue(true));
    });

    test('no match', () {
      final result = eval(registry.get('REGEXMATCH')!, [
        const TextNode('Hello World'),
        const TextNode('xyz'),
      ]);
      expect(result, const BooleanValue(false));
    });

    test('regex pattern', () {
      final result = eval(registry.get('REGEXMATCH')!, [
        const TextNode('abc123def'),
        const TextNode(r'\d+'),
      ]);
      expect(result, const BooleanValue(true));
    });

    test('case sensitivity', () {
      final result = eval(registry.get('REGEXMATCH')!, [
        const TextNode('Hello'),
        const TextNode('hello'),
      ]);
      expect(result, const BooleanValue(false));
    });

    test('case-insensitive using char class', () {
      final result = eval(registry.get('REGEXMATCH')!, [
        const TextNode('Hello'),
        const TextNode('[Hh]ello'),
      ]);
      expect(result, const BooleanValue(true));
    });

    test('invalid regex returns #VALUE!', () {
      final result = eval(registry.get('REGEXMATCH')!, [
        const TextNode('test'),
        const TextNode('[invalid'),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('full string match with anchors', () {
      final result = eval(registry.get('REGEXMATCH')!, [
        const TextNode('abc'),
        const TextNode(r'^[a-z]+$'),
      ]);
      expect(result, const BooleanValue(true));
    });
  });

  group('REGEXEXTRACT', () {
    test('match found', () {
      final result = eval(registry.get('REGEXEXTRACT')!, [
        const TextNode('Hello 123 World'),
        const TextNode(r'\d+'),
      ]);
      expect(result, const TextValue('123'));
    });

    test('capture group', () {
      final result = eval(registry.get('REGEXEXTRACT')!, [
        const TextNode('name: John'),
        const TextNode(r'name: (\w+)'),
      ]);
      expect(result, const TextValue('John'));
    });

    test('no match returns #N/A', () {
      final result = eval(registry.get('REGEXEXTRACT')!, [
        const TextNode('Hello World'),
        const TextNode(r'\d+'),
      ]);
      expect(result, const ErrorValue(FormulaError.na));
    });

    test('invalid regex returns #VALUE!', () {
      final result = eval(registry.get('REGEXEXTRACT')!, [
        const TextNode('test'),
        const TextNode('[invalid'),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('first match only', () {
      final result = eval(registry.get('REGEXEXTRACT')!, [
        const TextNode('abc 123 def 456'),
        const TextNode(r'\d+'),
      ]);
      expect(result, const TextValue('123'));
    });
  });

  group('REGEXREPLACE', () {
    test('basic replace', () {
      final result = eval(registry.get('REGEXREPLACE')!, [
        const TextNode('Hello World'),
        const TextNode('World'),
        const TextNode('Dart'),
      ]);
      expect(result, const TextValue('Hello Dart'));
    });

    test('regex patterns', () {
      final result = eval(registry.get('REGEXREPLACE')!, [
        const TextNode('abc 123 def 456'),
        const TextNode(r'\d+'),
        const TextNode('NUM'),
      ]);
      expect(result, const TextValue('abc NUM def NUM'));
    });

    test('empty replacement', () {
      final result = eval(registry.get('REGEXREPLACE')!, [
        const TextNode('Hello 123 World'),
        const TextNode(r'\d+'),
        const TextNode(''),
      ]);
      expect(result, const TextValue('Hello  World'));
    });

    test('no match leaves unchanged', () {
      final result = eval(registry.get('REGEXREPLACE')!, [
        const TextNode('Hello World'),
        const TextNode(r'\d+'),
        const TextNode('NUM'),
      ]);
      expect(result, const TextValue('Hello World'));
    });

    test('invalid regex returns #VALUE!', () {
      final result = eval(registry.get('REGEXREPLACE')!, [
        const TextNode('test'),
        const TextNode('[invalid'),
        const TextNode('x'),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('replaces all occurrences (not just first)', () {
      final result = eval(registry.get('REGEXREPLACE')!, [
        const TextNode('aaa'),
        const TextNode('a'),
        const TextNode('b'),
      ]);
      expect(result, const TextValue('bbb'));
    });
  });
}
