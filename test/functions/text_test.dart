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

    test('format 0.00 gives two decimal places', () {
      final result = eval(registry.get('TEXT')!, [
        const NumberNode(3.14159),
        const TextNode('0.00'),
      ]);
      expect(result, const TextValue('3.14'));
    });

    test('format #,##0 adds thousands separator', () {
      final result = eval(registry.get('TEXT')!, [
        const NumberNode(1234567),
        const TextNode('#,##0'),
      ]);
      expect(result, const TextValue('1,234,567'));
    });

    test('format 0% shows percentage', () {
      final result = eval(registry.get('TEXT')!, [
        const NumberNode(0.75),
        const TextNode('0%'),
      ]);
      expect(result, const TextValue('75%'));
    });

    test('format 0.0% shows percentage with decimal', () {
      final result = eval(registry.get('TEXT')!, [
        const NumberNode(0.756),
        const TextNode('0.0%'),
      ]);
      expect(result, const TextValue('75.6%'));
    });

    test('format 0 shows leading zeros with 000', () {
      final result = eval(registry.get('TEXT')!, [
        const NumberNode(5),
        const TextNode('000'),
      ]);
      expect(result, const TextValue('005'));
    });

    test('format 0.0E+0 scientific notation', () {
      final result = eval(registry.get('TEXT')!, [
        const NumberNode(1234),
        const TextNode('0.0E+0'),
      ]);
      expect(result, const TextValue('1.2E+3'));
    });

    test('format #,##0.00 combined', () {
      final result = eval(registry.get('TEXT')!, [
        const NumberNode(1234.5),
        const TextNode('#,##0.00'),
      ]);
      expect(result, const TextValue('1,234.50'));
    });

    test('format # no leading zeros for small fraction', () {
      final result = eval(registry.get('TEXT')!, [
        const NumberNode(0.5),
        const TextNode('#.##'),
      ]);
      expect(result, const TextValue('.5'));
    });
  });

  group('FIND', () {
    test('finds position case-sensitive', () {
      final result = eval(registry.get('FIND')!, [
        const TextNode('World'),
        const TextNode('Hello World'),
      ]);
      expect(result, const NumberValue(7));
    });

    test('case-sensitive: lowercase not found', () {
      final result = eval(registry.get('FIND')!, [
        const TextNode('world'),
        const TextNode('Hello World'),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('with start_num', () {
      final result = eval(registry.get('FIND')!, [
        const TextNode('l'),
        const TextNode('Hello'),
        const NumberNode(4),
      ]);
      expect(result, const NumberValue(4));
    });

    test('not found returns #VALUE!', () {
      final result = eval(registry.get('FIND')!, [
        const TextNode('xyz'),
        const TextNode('Hello'),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('SEARCH', () {
    test('case-insensitive search', () {
      final result = eval(registry.get('SEARCH')!, [
        const TextNode('world'),
        const TextNode('Hello World'),
      ]);
      expect(result, const NumberValue(7));
    });

    test('wildcard ? matches single char', () {
      final result = eval(registry.get('SEARCH')!, [
        const TextNode('h?llo'),
        const TextNode('Hello World'),
      ]);
      expect(result, const NumberValue(1));
    });

    test('wildcard * matches multiple chars', () {
      final result = eval(registry.get('SEARCH')!, [
        const TextNode('H*d'),
        const TextNode('Hello World'),
      ]);
      expect(result, const NumberValue(1));
    });

    test('not found returns #VALUE!', () {
      final result = eval(registry.get('SEARCH')!, [
        const TextNode('xyz'),
        const TextNode('Hello'),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('SUBSTITUTE', () {
    test('replaces all occurrences', () {
      final result = eval(registry.get('SUBSTITUTE')!, [
        const TextNode('Hello Hello'),
        const TextNode('Hello'),
        const TextNode('World'),
      ]);
      expect(result, const TextValue('World World'));
    });

    test('replaces specific instance', () {
      final result = eval(registry.get('SUBSTITUTE')!, [
        const TextNode('aaa'),
        const TextNode('a'),
        const TextNode('b'),
        const NumberNode(2),
      ]);
      expect(result, const TextValue('aba'));
    });

    test('empty old_text returns original', () {
      final result = eval(registry.get('SUBSTITUTE')!, [
        const TextNode('Hello'),
        const TextNode(''),
        const TextNode('World'),
      ]);
      expect(result, const TextValue('Hello'));
    });
  });

  group('REPLACE', () {
    test('replaces by position', () {
      final result = eval(registry.get('REPLACE')!, [
        const TextNode('Hello World'),
        const NumberNode(7),
        const NumberNode(5),
        const TextNode('Dart'),
      ]);
      expect(result, const TextValue('Hello Dart'));
    });

    test('start < 1 returns #VALUE!', () {
      final result = eval(registry.get('REPLACE')!, [
        const TextNode('Hello'),
        const NumberNode(0),
        const NumberNode(1),
        const TextNode('X'),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('VALUE', () {
    test('converts text to number', () {
      final result =
          eval(registry.get('VALUE')!, [const TextNode('42')]);
      expect(result, const NumberValue(42));
    });

    test('converts decimal text', () {
      final result =
          eval(registry.get('VALUE')!, [const TextNode('3.14')]);
      expect(result, const NumberValue(3.14));
    });

    test('non-numeric text returns #VALUE!', () {
      final result =
          eval(registry.get('VALUE')!, [const TextNode('abc')]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('number input passed through', () {
      final result =
          eval(registry.get('VALUE')!, [const NumberNode(42)]);
      expect(result, const NumberValue(42));
    });
  });

  group('TEXTJOIN', () {
    test('joins with delimiter', () {
      final result = eval(registry.get('TEXTJOIN')!, [
        const TextNode(', '),
        const BooleanNode(false),
        const TextNode('a'),
        const TextNode('b'),
        const TextNode('c'),
      ]);
      expect(result, const TextValue('a, b, c'));
    });

    test('ignores empty when true', () {
      final result = eval(registry.get('TEXTJOIN')!, [
        const TextNode('-'),
        const BooleanNode(true),
        const TextNode('a'),
        const TextNode(''),
        const TextNode('c'),
      ]);
      expect(result, const TextValue('a-c'));
    });

    test('includes empty when false', () {
      final result = eval(registry.get('TEXTJOIN')!, [
        const TextNode('-'),
        const BooleanNode(false),
        const TextNode('a'),
        const TextNode(''),
        const TextNode('c'),
      ]);
      expect(result, const TextValue('a--c'));
    });
  });

  group('PROPER', () {
    test('capitalizes first letter of each word', () {
      final result =
          eval(registry.get('PROPER')!, [const TextNode('hello world')]);
      expect(result, const TextValue('Hello World'));
    });

    test('handles mixed case', () {
      final result =
          eval(registry.get('PROPER')!, [const TextNode('hELLO wORLD')]);
      expect(result, const TextValue('Hello World'));
    });

    test('handles single word', () {
      final result =
          eval(registry.get('PROPER')!, [const TextNode('hello')]);
      expect(result, const TextValue('Hello'));
    });
  });

  group('EXACT', () {
    test('returns true for identical strings', () {
      final result = eval(registry.get('EXACT')!, [
        const TextNode('Hello'),
        const TextNode('Hello'),
      ]);
      expect(result, const BooleanValue(true));
    });

    test('returns false for different case', () {
      final result = eval(registry.get('EXACT')!, [
        const TextNode('Hello'),
        const TextNode('hello'),
      ]);
      expect(result, const BooleanValue(false));
    });

    test('returns false for different strings', () {
      final result = eval(registry.get('EXACT')!, [
        const TextNode('Hello'),
        const TextNode('World'),
      ]);
      expect(result, const BooleanValue(false));
    });
  });

  group('REPT', () {
    test('repeats text', () {
      final result = eval(registry.get('REPT')!, [
        const TextNode('abc'),
        const NumberNode(3),
      ]);
      expect(result, const TextValue('abcabcabc'));
    });

    test('count 0 returns empty', () {
      final result = eval(registry.get('REPT')!, [
        const TextNode('abc'),
        const NumberNode(0),
      ]);
      expect(result, const TextValue(''));
    });

    test('negative count returns #VALUE!', () {
      final result = eval(registry.get('REPT')!, [
        const TextNode('abc'),
        const NumberNode(-1),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('CHAR', () {
    test('returns character for code', () {
      final result = eval(registry.get('CHAR')!, [const NumberNode(65)]);
      expect(result, const TextValue('A'));
    });

    test('code 32 returns space', () {
      final result = eval(registry.get('CHAR')!, [const NumberNode(32)]);
      expect(result, const TextValue(' '));
    });

    test('code < 1 returns #VALUE!', () {
      final result = eval(registry.get('CHAR')!, [const NumberNode(0)]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('code > 255 returns #VALUE!', () {
      final result = eval(registry.get('CHAR')!, [const NumberNode(256)]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('CODE', () {
    test('returns code for character', () {
      final result = eval(registry.get('CODE')!, [const TextNode('A')]);
      expect(result, const NumberValue(65));
    });

    test('returns code for first character', () {
      final result = eval(registry.get('CODE')!, [const TextNode('ABC')]);
      expect(result, const NumberValue(65));
    });

    test('empty text returns #VALUE!', () {
      final result = eval(registry.get('CODE')!, [const TextNode('')]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('CLEAN', () {
    test('removes non-printable characters', () {
      final result = eval(
          registry.get('CLEAN')!, [TextNode('Hello\x00\x01World\x1F')]);
      expect(result, const TextValue('HelloWorld'));
    });

    test('leaves printable characters unchanged', () {
      final result =
          eval(registry.get('CLEAN')!, [const TextNode('Hello World')]);
      expect(result, const TextValue('Hello World'));
    });
  });

  group('DOLLAR', () {
    test('formats positive number', () {
      final result = eval(registry.get('DOLLAR')!, [const NumberNode(1234.56)]);
      expect(result, const TextValue('\$1,234.56'));
    });

    test('formats negative number with parentheses', () {
      final result =
          eval(registry.get('DOLLAR')!, [const NumberNode(-1234.56)]);
      expect(result, const TextValue('(\$1,234.56)'));
    });

    test('custom decimals', () {
      final result = eval(registry.get('DOLLAR')!, [
        const NumberNode(1234.567),
        const NumberNode(1),
      ]);
      expect(result, const TextValue('\$1,234.6'));
    });

    test('zero decimals', () {
      final result = eval(registry.get('DOLLAR')!, [
        const NumberNode(1234.56),
        const NumberNode(0),
      ]);
      expect(result, const TextValue('\$1,235'));
    });
  });

  group('FIXED', () {
    test('formats with decimals', () {
      final result = eval(registry.get('FIXED')!, [
        const NumberNode(1234.567),
        const NumberNode(1),
      ]);
      expect(result, const TextValue('1,234.6'));
    });

    test('no commas when requested', () {
      final result = eval(registry.get('FIXED')!, [
        const NumberNode(1234.567),
        const NumberNode(1),
        const BooleanNode(true),
      ]);
      expect(result, const TextValue('1234.6'));
    });

    test('zero decimals', () {
      final result = eval(registry.get('FIXED')!, [
        const NumberNode(1234.567),
        const NumberNode(0),
      ]);
      expect(result, const TextValue('1,235'));
    });

    test('negative decimals rounds to left', () {
      final result = eval(registry.get('FIXED')!, [
        const NumberNode(1234),
        const NumberNode(-2),
      ]);
      expect(result, const TextValue('1,200'));
    });
  });

  group('T', () {
    test('returns text when text', () {
      final result = eval(registry.get('T')!, [const TextNode('hello')]);
      expect(result, const TextValue('hello'));
    });

    test('returns empty for number', () {
      final result = eval(registry.get('T')!, [const NumberNode(42)]);
      expect(result, const TextValue(''));
    });

    test('returns empty for boolean', () {
      final result = eval(registry.get('T')!, [const BooleanNode(true)]);
      expect(result, const TextValue(''));
    });

    test('propagates error', () {
      final result =
          eval(registry.get('T')!, [const ErrorNode(FormulaError.ref)]);
      expect(result, const ErrorValue(FormulaError.ref));
    });
  });

  group('NUMBERVALUE', () {
    test('parses simple number', () {
      final result =
          eval(registry.get('NUMBERVALUE')!, [const TextNode('1234.56')]);
      expect(result, const NumberValue(1234.56));
    });

    test('custom separators', () {
      final result = eval(registry.get('NUMBERVALUE')!, [
        const TextNode('1.234,56'),
        const TextNode(','),
        const TextNode('.'),
      ]);
      expect(result, const NumberValue(1234.56));
    });

    test('percentage', () {
      final result =
          eval(registry.get('NUMBERVALUE')!, [const TextNode('75%')]);
      expect(result, const NumberValue(0.75));
    });

    test('invalid text returns #VALUE!', () {
      final result =
          eval(registry.get('NUMBERVALUE')!, [const TextNode('abc')]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('UNICHAR', () {
    test('returns unicode character', () {
      final result = eval(registry.get('UNICHAR')!, [const NumberNode(65)]);
      expect(result, const TextValue('A'));
    });

    test('returns emoji', () {
      final result = eval(registry.get('UNICHAR')!, [const NumberNode(9829)]);
      expect(result, const TextValue('\u2665'));
    });

    test('code < 1 returns #VALUE!', () {
      final result = eval(registry.get('UNICHAR')!, [const NumberNode(0)]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('UNICODE', () {
    test('returns code point', () {
      final result = eval(registry.get('UNICODE')!, [const TextNode('A')]);
      expect(result, const NumberValue(65));
    });

    test('empty text returns #VALUE!', () {
      final result = eval(registry.get('UNICODE')!, [const TextNode('')]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('TEXTBEFORE', () {
    test('returns text before delimiter', () {
      final result = eval(registry.get('TEXTBEFORE')!, [
        const TextNode('Hello-World-Test'),
        const TextNode('-'),
      ]);
      expect(result, const TextValue('Hello'));
    });

    test('second instance', () {
      final result = eval(registry.get('TEXTBEFORE')!, [
        const TextNode('Hello-World-Test'),
        const TextNode('-'),
        const NumberNode(2),
      ]);
      expect(result, const TextValue('Hello-World'));
    });

    test('negative instance counts from end', () {
      final result = eval(registry.get('TEXTBEFORE')!, [
        const TextNode('Hello-World-Test'),
        const TextNode('-'),
        const NumberNode(-1),
      ]);
      expect(result, const TextValue('Hello-World'));
    });

    test('case-insensitive match', () {
      final result = eval(registry.get('TEXTBEFORE')!, [
        const TextNode('Hello WORLD test'),
        const TextNode('world'),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      expect(result, const TextValue('Hello '));
    });

    test('not found returns #N/A', () {
      final result = eval(registry.get('TEXTBEFORE')!, [
        const TextNode('Hello'),
        const TextNode('-'),
      ]);
      expect(result, const ErrorValue(FormulaError.na));
    });

    test('if_not_found returns custom value', () {
      final result = eval(registry.get('TEXTBEFORE')!, [
        const TextNode('Hello'),
        const TextNode('-'),
        const NumberNode(1),
        const NumberNode(0),
        const NumberNode(0),
        const TextNode('NOT FOUND'),
      ]);
      expect(result, const TextValue('NOT FOUND'));
    });
  });

  group('TEXTAFTER', () {
    test('returns text after delimiter', () {
      final result = eval(registry.get('TEXTAFTER')!, [
        const TextNode('Hello-World-Test'),
        const TextNode('-'),
      ]);
      expect(result, const TextValue('World-Test'));
    });

    test('second instance', () {
      final result = eval(registry.get('TEXTAFTER')!, [
        const TextNode('Hello-World-Test'),
        const TextNode('-'),
        const NumberNode(2),
      ]);
      expect(result, const TextValue('Test'));
    });

    test('negative instance counts from end', () {
      final result = eval(registry.get('TEXTAFTER')!, [
        const TextNode('Hello-World-Test'),
        const TextNode('-'),
        const NumberNode(-1),
      ]);
      expect(result, const TextValue('Test'));
    });

    test('not found returns #N/A', () {
      final result = eval(registry.get('TEXTAFTER')!, [
        const TextNode('Hello'),
        const TextNode('-'),
      ]);
      expect(result, const ErrorValue(FormulaError.na));
    });
  });

  group('TEXTSPLIT', () {
    test('splits by column delimiter', () {
      final result = eval(registry.get('TEXTSPLIT')!, [
        const TextNode('a,b,c'),
        const TextNode(','),
      ]);
      expect(result, isA<RangeValue>());
      final range = result as RangeValue;
      expect(range.rowCount, 1);
      expect(range.columnCount, 3);
      expect(range.values[0][0], const TextValue('a'));
      expect(range.values[0][1], const TextValue('b'));
      expect(range.values[0][2], const TextValue('c'));
    });

    test('splits by row and column delimiters', () {
      final result = eval(registry.get('TEXTSPLIT')!, [
        const TextNode('a,b;c,d'),
        const TextNode(','),
        const TextNode(';'),
      ]);
      expect(result, isA<RangeValue>());
      final range = result as RangeValue;
      expect(range.rowCount, 2);
      expect(range.columnCount, 2);
      expect(range.values[0][0], const TextValue('a'));
      expect(range.values[0][1], const TextValue('b'));
      expect(range.values[1][0], const TextValue('c'));
      expect(range.values[1][1], const TextValue('d'));
    });

    test('pads uneven rows with #N/A', () {
      final result = eval(registry.get('TEXTSPLIT')!, [
        const TextNode('a,b;c'),
        const TextNode(','),
        const TextNode(';'),
      ]);
      expect(result, isA<RangeValue>());
      final range = result as RangeValue;
      expect(range.rowCount, 2);
      expect(range.columnCount, 2);
      expect(range.values[0][0], const TextValue('a'));
      expect(range.values[0][1], const TextValue('b'));
      expect(range.values[1][0], const TextValue('c'));
      expect(range.values[1][1], const ErrorValue(FormulaError.na));
    });

    test('single value returns text', () {
      final result = eval(registry.get('TEXTSPLIT')!, [
        const TextNode('hello'),
        const TextNode(','),
      ]);
      expect(result, const TextValue('hello'));
    });
  });
}
