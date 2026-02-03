import 'package:a1/a1.dart';
import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/functions/function.dart';
import 'package:worksheet_formula/src/functions/registry.dart';
import 'package:worksheet_formula/src/functions/engineering.dart';

class _TestContext implements EvaluationContext {
  @override
  A1 get currentCell => 'A1'.a1;
  @override
  String? get currentSheet => null;
  @override
  bool get isCancelled => false;

  final FunctionRegistry _registry;
  final Map<String, FormulaValue> rangeOverrides = {};
  _TestContext(this._registry);

  @override
  FormulaValue getCellValue(A1 cell) => const EmptyValue();
  @override
  FormulaValue getRangeValues(A1Range range) {
    final key = range.toString();
    return rangeOverrides[key] ?? const FormulaValue.error(FormulaError.ref);
  }

  @override
  FormulaFunction? getFunction(String name) => _registry.get(name);

  void setRange(String ref, FormulaValue value) {
    final a1Ref = A1Reference.parse(ref);
    rangeOverrides[a1Ref.range.toString()] = value;
  }
}

void main() {
  late FunctionRegistry registry;
  late _TestContext context;

  setUp(() {
    registry = FunctionRegistry(registerBuiltIns: false);
    registerEngineeringFunctions(registry);
    context = _TestContext(registry);
  });

  FormulaValue eval(FormulaFunction fn, List<FormulaNode> args) =>
      fn.call(args, context);

  // =========================================================================
  // Wave 1 — Comparison + Bitwise
  // =========================================================================

  group('DELTA', () {
    test('equal values return 1', () {
      final result =
          eval(registry.get('DELTA')!, [const NumberNode(5), const NumberNode(5)]);
      expect((result as NumberValue).value, 1);
    });

    test('unequal values return 0', () {
      final result =
          eval(registry.get('DELTA')!, [const NumberNode(5), const NumberNode(4)]);
      expect((result as NumberValue).value, 0);
    });

    test('single arg defaults second to 0', () {
      final result = eval(registry.get('DELTA')!, [const NumberNode(0)]);
      expect((result as NumberValue).value, 1);
    });

    test('single arg non-zero returns 0', () {
      final result = eval(registry.get('DELTA')!, [const NumberNode(5)]);
      expect((result as NumberValue).value, 0);
    });

    test('non-numeric returns #VALUE!', () {
      final result = eval(
          registry.get('DELTA')!, [const TextNode('abc'), const NumberNode(0)]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('GESTEP', () {
    test('number >= step returns 1', () {
      final result = eval(
          registry.get('GESTEP')!, [const NumberNode(5), const NumberNode(4)]);
      expect((result as NumberValue).value, 1);
    });

    test('number < step returns 0', () {
      final result = eval(
          registry.get('GESTEP')!, [const NumberNode(3), const NumberNode(4)]);
      expect((result as NumberValue).value, 0);
    });

    test('equal returns 1', () {
      final result = eval(
          registry.get('GESTEP')!, [const NumberNode(4), const NumberNode(4)]);
      expect((result as NumberValue).value, 1);
    });

    test('single arg defaults step to 0', () {
      final result = eval(registry.get('GESTEP')!, [const NumberNode(1)]);
      expect((result as NumberValue).value, 1);
    });

    test('negative number single arg returns 0', () {
      final result = eval(registry.get('GESTEP')!, [const NumberNode(-1)]);
      expect((result as NumberValue).value, 0);
    });

    test('non-numeric returns #VALUE!', () {
      final result = eval(registry.get('GESTEP')!, [const TextNode('abc')]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('BITAND', () {
    test('basic AND', () {
      final result = eval(
          registry.get('BITAND')!, [const NumberNode(13), const NumberNode(25)]);
      expect((result as NumberValue).value, 9);
    });

    test('zero AND returns 0', () {
      final result = eval(
          registry.get('BITAND')!, [const NumberNode(0), const NumberNode(255)]);
      expect((result as NumberValue).value, 0);
    });

    test('negative returns #NUM!', () {
      final result = eval(
          registry.get('BITAND')!, [const NumberNode(-1), const NumberNode(5)]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('exceeds 2^48-1 returns #NUM!', () {
      final result = eval(registry.get('BITAND')!,
          [const NumberNode(281474976710656), const NumberNode(5)]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('non-integer returns #NUM!', () {
      final result = eval(registry.get('BITAND')!,
          [const NumberNode(1.5), const NumberNode(5)]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('BITOR', () {
    test('basic OR', () {
      final result = eval(
          registry.get('BITOR')!, [const NumberNode(23), const NumberNode(10)]);
      expect((result as NumberValue).value, 31);
    });

    test('negative returns #NUM!', () {
      final result = eval(
          registry.get('BITOR')!, [const NumberNode(-1), const NumberNode(5)]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('BITXOR', () {
    test('basic XOR', () {
      final result = eval(
          registry.get('BITXOR')!, [const NumberNode(5), const NumberNode(3)]);
      expect((result as NumberValue).value, 6);
    });

    test('negative returns #NUM!', () {
      final result = eval(
          registry.get('BITXOR')!, [const NumberNode(-1), const NumberNode(5)]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('BITLSHIFT', () {
    test('left shift', () {
      final result = eval(registry.get('BITLSHIFT')!,
          [const NumberNode(4), const NumberNode(2)]);
      expect((result as NumberValue).value, 16);
    });

    test('negative shift = right shift', () {
      final result = eval(registry.get('BITLSHIFT')!,
          [const NumberNode(16), const NumberNode(-2)]);
      expect((result as NumberValue).value, 4);
    });

    test('result exceeds 48 bits returns #NUM!', () {
      final result = eval(registry.get('BITLSHIFT')!,
          [const NumberNode(1), const NumberNode(48)]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('negative number returns #NUM!', () {
      final result = eval(registry.get('BITLSHIFT')!,
          [const NumberNode(-1), const NumberNode(2)]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('BITRSHIFT', () {
    test('right shift', () {
      final result = eval(registry.get('BITRSHIFT')!,
          [const NumberNode(13), const NumberNode(2)]);
      expect((result as NumberValue).value, 3);
    });

    test('negative shift = left shift', () {
      final result = eval(registry.get('BITRSHIFT')!,
          [const NumberNode(4), const NumberNode(-2)]);
      expect((result as NumberValue).value, 16);
    });

    test('negative number returns #NUM!', () {
      final result = eval(registry.get('BITRSHIFT')!,
          [const NumberNode(-1), const NumberNode(2)]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  // =========================================================================
  // Wave 2 — Base Conversion
  // =========================================================================

  group('BIN2DEC', () {
    test('positive binary', () {
      final result =
          eval(registry.get('BIN2DEC')!, [const TextNode('1100100')]);
      expect((result as NumberValue).value, 100);
    });

    test('negative binary (two\'s complement)', () {
      final result =
          eval(registry.get('BIN2DEC')!, [const TextNode('1111111111')]);
      expect((result as NumberValue).value, -1);
    });

    test('zero', () {
      final result = eval(registry.get('BIN2DEC')!, [const TextNode('0')]);
      expect((result as NumberValue).value, 0);
    });

    test('invalid binary returns #NUM!', () {
      final result = eval(registry.get('BIN2DEC')!, [const TextNode('2')]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('too many digits returns #NUM!', () {
      final result =
          eval(registry.get('BIN2DEC')!, [const TextNode('11111111111')]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('BIN2HEX', () {
    test('basic conversion', () {
      final result =
          eval(registry.get('BIN2HEX')!, [const TextNode('11111011')]);
      expect((result as TextValue).value, 'FB');
    });

    test('with places', () {
      final result = eval(registry.get('BIN2HEX')!,
          [const TextNode('1110'), const NumberNode(4)]);
      expect((result as TextValue).value, '000E');
    });

    test('negative binary', () {
      final result =
          eval(registry.get('BIN2HEX')!, [const TextNode('1110000000')]);
      expect((result as TextValue).value, 'FFFFFFFF80');
    });

    test('places too small returns #NUM!', () {
      final result = eval(registry.get('BIN2HEX')!,
          [const TextNode('11111011'), const NumberNode(1)]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('BIN2OCT', () {
    test('basic conversion', () {
      final result =
          eval(registry.get('BIN2OCT')!, [const TextNode('1001')]);
      expect((result as TextValue).value, '11');
    });

    test('with places', () {
      final result = eval(registry.get('BIN2OCT')!,
          [const TextNode('1100'), const NumberNode(4)]);
      expect((result as TextValue).value, '0014');
    });

    test('negative binary', () {
      final result =
          eval(registry.get('BIN2OCT')!, [const TextNode('1111111111')]);
      expect((result as TextValue).value, '7777777777');
    });
  });

  group('DEC2BIN', () {
    test('positive number', () {
      final result = eval(registry.get('DEC2BIN')!, [const NumberNode(100)]);
      expect((result as TextValue).value, '1100100');
    });

    test('negative number', () {
      final result = eval(registry.get('DEC2BIN')!, [const NumberNode(-100)]);
      expect((result as TextValue).value, '1110011100');
    });

    test('zero', () {
      final result = eval(registry.get('DEC2BIN')!, [const NumberNode(0)]);
      expect((result as TextValue).value, '0');
    });

    test('with places', () {
      final result = eval(registry.get('DEC2BIN')!,
          [const NumberNode(9), const NumberNode(6)]);
      expect((result as TextValue).value, '001001');
    });

    test('out of range returns #NUM!', () {
      final result = eval(registry.get('DEC2BIN')!, [const NumberNode(512)]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('-513 out of range returns #NUM!', () {
      final result = eval(registry.get('DEC2BIN')!, [const NumberNode(-513)]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('DEC2HEX', () {
    test('positive number', () {
      final result = eval(registry.get('DEC2HEX')!, [const NumberNode(100)]);
      expect((result as TextValue).value, '64');
    });

    test('negative number', () {
      final result = eval(registry.get('DEC2HEX')!, [const NumberNode(-54)]);
      expect((result as TextValue).value, 'FFFFFFFFCA');
    });

    test('with places', () {
      final result = eval(registry.get('DEC2HEX')!,
          [const NumberNode(100), const NumberNode(4)]);
      expect((result as TextValue).value, '0064');
    });
  });

  group('DEC2OCT', () {
    test('positive number', () {
      final result = eval(registry.get('DEC2OCT')!, [const NumberNode(58)]);
      expect((result as TextValue).value, '72');
    });

    test('negative number', () {
      final result = eval(registry.get('DEC2OCT')!, [const NumberNode(-100)]);
      expect((result as TextValue).value, '7777777634');
    });

    test('with places', () {
      final result = eval(registry.get('DEC2OCT')!,
          [const NumberNode(58), const NumberNode(4)]);
      expect((result as TextValue).value, '0072');
    });
  });

  group('HEX2BIN', () {
    test('basic conversion', () {
      final result = eval(registry.get('HEX2BIN')!, [const TextNode('F')]);
      expect((result as TextValue).value, '1111');
    });

    test('with places', () {
      final result = eval(registry.get('HEX2BIN')!,
          [const TextNode('F'), const NumberNode(8)]);
      expect((result as TextValue).value, '00001111');
    });

    test('negative hex', () {
      final result =
          eval(registry.get('HEX2BIN')!, [const TextNode('FFFFFFFE00')]);
      expect((result as TextValue).value, '1000000000');
    });

    test('value at binary limit (511) succeeds', () {
      final result = eval(registry.get('HEX2BIN')!, [const TextNode('1FF')]);
      expect((result as TextValue).value, '111111111');
    });

    test('too large for binary returns #NUM!', () {
      final result = eval(registry.get('HEX2BIN')!, [const TextNode('200')]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('HEX2DEC', () {
    test('positive hex', () {
      final result = eval(registry.get('HEX2DEC')!, [const TextNode('A5')]);
      expect((result as NumberValue).value, 165);
    });

    test('negative hex (two\'s complement)', () {
      final result =
          eval(registry.get('HEX2DEC')!, [const TextNode('FFFFFFFFFF')]);
      expect((result as NumberValue).value, -1);
    });

    test('zero', () {
      final result = eval(registry.get('HEX2DEC')!, [const TextNode('0')]);
      expect((result as NumberValue).value, 0);
    });

    test('too many digits returns #NUM!', () {
      final result =
          eval(registry.get('HEX2DEC')!, [const TextNode('1FFFFFFFFFF')]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('HEX2OCT', () {
    test('basic conversion', () {
      final result = eval(registry.get('HEX2OCT')!, [const TextNode('F')]);
      expect((result as TextValue).value, '17');
    });

    test('with places', () {
      final result = eval(registry.get('HEX2OCT')!,
          [const TextNode('F'), const NumberNode(4)]);
      expect((result as TextValue).value, '0017');
    });

    test('negative hex', () {
      final result =
          eval(registry.get('HEX2OCT')!, [const TextNode('FFFFFFFF00')]);
      expect((result as TextValue).value, '7777777400');
    });
  });

  group('OCT2BIN', () {
    test('basic conversion', () {
      final result = eval(registry.get('OCT2BIN')!, [const TextNode('3')]);
      expect((result as TextValue).value, '11');
    });

    test('with places', () {
      final result = eval(registry.get('OCT2BIN')!,
          [const TextNode('3'), const NumberNode(4)]);
      expect((result as TextValue).value, '0011');
    });

    test('negative octal', () {
      final result =
          eval(registry.get('OCT2BIN')!, [const TextNode('7777777000')]);
      expect((result as TextValue).value, '1000000000');
    });

    test('value at binary limit (777 octal = 511) succeeds', () {
      final result = eval(registry.get('OCT2BIN')!, [const TextNode('777')]);
      expect((result as TextValue).value, '111111111');
    });

    test('too large for binary returns #NUM!', () {
      final result = eval(registry.get('OCT2BIN')!, [const TextNode('1000')]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('OCT2DEC', () {
    test('positive octal', () {
      final result = eval(registry.get('OCT2DEC')!, [const TextNode('54')]);
      expect((result as NumberValue).value, 44);
    });

    test('negative octal (two\'s complement)', () {
      final result =
          eval(registry.get('OCT2DEC')!, [const TextNode('7777777777')]);
      expect((result as NumberValue).value, -1);
    });
  });

  group('OCT2HEX', () {
    test('basic conversion', () {
      final result = eval(registry.get('OCT2HEX')!, [const TextNode('100')]);
      expect((result as TextValue).value, '40');
    });

    test('with places', () {
      final result = eval(registry.get('OCT2HEX')!,
          [const TextNode('100'), const NumberNode(4)]);
      expect((result as TextValue).value, '0040');
    });

    test('negative octal', () {
      final result =
          eval(registry.get('OCT2HEX')!, [const TextNode('7777777000')]);
      expect((result as TextValue).value, 'FFFFFFFE00');
    });
  });

  // =========================================================================
  // Wave 3 — Number Format + Error Functions
  // =========================================================================

  group('BASE', () {
    test('decimal to binary', () {
      final result = eval(
          registry.get('BASE')!, [const NumberNode(7), const NumberNode(2)]);
      expect((result as TextValue).value, '111');
    });

    test('decimal to hex', () {
      final result = eval(
          registry.get('BASE')!, [const NumberNode(255), const NumberNode(16)]);
      expect((result as TextValue).value, 'FF');
    });

    test('with min_length', () {
      final result = eval(registry.get('BASE')!,
          [const NumberNode(7), const NumberNode(2), const NumberNode(8)]);
      expect((result as TextValue).value, '00000111');
    });

    test('base < 2 returns #NUM!', () {
      final result = eval(
          registry.get('BASE')!, [const NumberNode(7), const NumberNode(1)]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('base > 36 returns #NUM!', () {
      final result = eval(
          registry.get('BASE')!, [const NumberNode(7), const NumberNode(37)]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('negative number returns #NUM!', () {
      final result = eval(
          registry.get('BASE')!, [const NumberNode(-1), const NumberNode(2)]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('zero returns "0"', () {
      final result = eval(
          registry.get('BASE')!, [const NumberNode(0), const NumberNode(16)]);
      expect((result as TextValue).value, '0');
    });
  });

  group('DECIMAL', () {
    test('binary to decimal', () {
      final result = eval(
          registry.get('DECIMAL')!, [const TextNode('111'), const NumberNode(2)]);
      expect((result as NumberValue).value, 7);
    });

    test('hex to decimal', () {
      final result = eval(registry.get('DECIMAL')!,
          [const TextNode('FF'), const NumberNode(16)]);
      expect((result as NumberValue).value, 255);
    });

    test('base < 2 returns #NUM!', () {
      final result = eval(
          registry.get('DECIMAL')!, [const TextNode('1'), const NumberNode(1)]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('base > 36 returns #NUM!', () {
      final result = eval(registry.get('DECIMAL')!,
          [const TextNode('1'), const NumberNode(37)]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('invalid digit returns #NUM!', () {
      final result = eval(
          registry.get('DECIMAL')!, [const TextNode('G'), const NumberNode(16)]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('ARABIC', () {
    test('simple numeral', () {
      final result = eval(registry.get('ARABIC')!, [const TextNode('MCMXCIX')]);
      expect((result as NumberValue).value, 1999);
    });

    test('simple numeral II', () {
      final result = eval(registry.get('ARABIC')!, [const TextNode('LVII')]);
      expect((result as NumberValue).value, 57);
    });

    test('empty string returns 0', () {
      final result = eval(registry.get('ARABIC')!, [const TextNode('')]);
      expect((result as NumberValue).value, 0);
    });

    test('negative Roman', () {
      final result = eval(registry.get('ARABIC')!, [const TextNode('-XIV')]);
      expect((result as NumberValue).value, -14);
    });

    test('invalid returns #VALUE!', () {
      final result = eval(registry.get('ARABIC')!, [const TextNode('ABC')]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('ROMAN', () {
    test('basic conversion', () {
      final result = eval(registry.get('ROMAN')!, [const NumberNode(499)]);
      expect((result as TextValue).value, 'CDXCIX');
    });

    test('zero returns empty string', () {
      final result = eval(registry.get('ROMAN')!, [const NumberNode(0)]);
      expect((result as TextValue).value, '');
    });

    test('3999', () {
      final result = eval(registry.get('ROMAN')!, [const NumberNode(3999)]);
      expect((result as TextValue).value, 'MMMCMXCIX');
    });

    test('negative returns #VALUE!', () {
      final result = eval(registry.get('ROMAN')!, [const NumberNode(-1)]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('above 3999 returns #VALUE!', () {
      final result = eval(registry.get('ROMAN')!, [const NumberNode(4000)]);
      expect(result, const ErrorValue(FormulaError.value));
    });

    test('classic form (form 0)', () {
      final result = eval(registry.get('ROMAN')!,
          [const NumberNode(499), const NumberNode(0)]);
      expect((result as TextValue).value, 'CDXCIX');
    });
  });

  group('ERF', () {
    test('single arg', () {
      final result = eval(registry.get('ERF')!, [const NumberNode(1)]);
      expect((result as NumberValue).value, closeTo(0.8427, 0.001));
    });

    test('two args: erf(upper) - erf(lower)', () {
      final result = eval(
          registry.get('ERF')!, [const NumberNode(0), const NumberNode(1)]);
      expect((result as NumberValue).value, closeTo(0.8427, 0.001));
    });

    test('erf(0) = 0', () {
      final result = eval(registry.get('ERF')!, [const NumberNode(0)]);
      expect((result as NumberValue).value, closeTo(0.0, 0.0001));
    });

    test('negative arg', () {
      final result = eval(registry.get('ERF')!, [const NumberNode(-1)]);
      expect((result as NumberValue).value, closeTo(-0.8427, 0.001));
    });
  });

  group('ERF.PRECISE', () {
    test('basic', () {
      final result =
          eval(registry.get('ERF.PRECISE')!, [const NumberNode(1)]);
      expect((result as NumberValue).value, closeTo(0.8427, 0.001));
    });
  });

  group('ERFC', () {
    test('basic', () {
      final result = eval(registry.get('ERFC')!, [const NumberNode(1)]);
      expect((result as NumberValue).value, closeTo(0.1573, 0.001));
    });

    test('erfc(0) = 1', () {
      final result = eval(registry.get('ERFC')!, [const NumberNode(0)]);
      expect((result as NumberValue).value, closeTo(1.0, 0.0001));
    });
  });

  group('ERFC.PRECISE', () {
    test('basic', () {
      final result =
          eval(registry.get('ERFC.PRECISE')!, [const NumberNode(1)]);
      expect((result as NumberValue).value, closeTo(0.1573, 0.001));
    });
  });

  // =========================================================================
  // Wave 4 — Complex Number Functions
  // =========================================================================

  group('COMPLEX', () {
    test('positive imaginary', () {
      final result = eval(registry.get('COMPLEX')!,
          [const NumberNode(3), const NumberNode(4)]);
      expect((result as TextValue).value, '3+4i');
    });

    test('negative imaginary', () {
      final result = eval(registry.get('COMPLEX')!,
          [const NumberNode(3), const NumberNode(-4)]);
      expect((result as TextValue).value, '3-4i');
    });

    test('zero imaginary', () {
      final result = eval(registry.get('COMPLEX')!,
          [const NumberNode(3), const NumberNode(0)]);
      expect((result as TextValue).value, '3');
    });

    test('zero real', () {
      final result = eval(registry.get('COMPLEX')!,
          [const NumberNode(0), const NumberNode(4)]);
      expect((result as TextValue).value, '4i');
    });

    test('both zero', () {
      final result = eval(registry.get('COMPLEX')!,
          [const NumberNode(0), const NumberNode(0)]);
      expect((result as TextValue).value, '0');
    });

    test('custom suffix j', () {
      final result = eval(registry.get('COMPLEX')!,
          [const NumberNode(3), const NumberNode(4), const TextNode('j')]);
      expect((result as TextValue).value, '3+4j');
    });

    test('imaginary = 1', () {
      final result = eval(registry.get('COMPLEX')!,
          [const NumberNode(0), const NumberNode(1)]);
      expect((result as TextValue).value, 'i');
    });

    test('imaginary = -1', () {
      final result = eval(registry.get('COMPLEX')!,
          [const NumberNode(0), const NumberNode(-1)]);
      expect((result as TextValue).value, '-i');
    });

    test('invalid suffix returns #VALUE!', () {
      final result = eval(registry.get('COMPLEX')!,
          [const NumberNode(1), const NumberNode(2), const TextNode('k')]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('IMREAL', () {
    test('positive complex', () {
      final result = eval(registry.get('IMREAL')!, [const TextNode('3+4i')]);
      expect((result as NumberValue).value, closeTo(3, 0.0001));
    });

    test('pure imaginary', () {
      final result = eval(registry.get('IMREAL')!, [const TextNode('4i')]);
      expect((result as NumberValue).value, closeTo(0, 0.0001));
    });

    test('real only', () {
      final result = eval(registry.get('IMREAL')!, [const TextNode('5')]);
      expect((result as NumberValue).value, closeTo(5, 0.0001));
    });

    test('invalid returns #NUM!', () {
      final result = eval(registry.get('IMREAL')!, [const TextNode('abc')]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('IMAGINARY', () {
    test('positive complex', () {
      final result =
          eval(registry.get('IMAGINARY')!, [const TextNode('3+4i')]);
      expect((result as NumberValue).value, closeTo(4, 0.0001));
    });

    test('real only', () {
      final result = eval(registry.get('IMAGINARY')!, [const TextNode('5')]);
      expect((result as NumberValue).value, closeTo(0, 0.0001));
    });

    test('negative imaginary', () {
      final result =
          eval(registry.get('IMAGINARY')!, [const TextNode('3-4i')]);
      expect((result as NumberValue).value, closeTo(-4, 0.0001));
    });
  });

  group('IMABS', () {
    test('3+4i = 5', () {
      final result = eval(registry.get('IMABS')!, [const TextNode('3+4i')]);
      expect((result as NumberValue).value, closeTo(5, 0.0001));
    });

    test('pure real', () {
      final result = eval(registry.get('IMABS')!, [const TextNode('-3')]);
      expect((result as NumberValue).value, closeTo(3, 0.0001));
    });
  });

  group('IMARGUMENT', () {
    test('first quadrant', () {
      final result =
          eval(registry.get('IMARGUMENT')!, [const TextNode('3+4i')]);
      expect((result as NumberValue).value, closeTo(0.9273, 0.001));
    });

    test('zero returns #DIV/0!', () {
      final result = eval(registry.get('IMARGUMENT')!, [const TextNode('0')]);
      expect(result, const ErrorValue(FormulaError.divZero));
    });

    test('pure imaginary', () {
      final result = eval(registry.get('IMARGUMENT')!, [const TextNode('i')]);
      expect(
          (result as NumberValue).value, closeTo(1.5708, 0.001)); // pi/2
    });
  });

  group('IMCONJUGATE', () {
    test('basic', () {
      final result =
          eval(registry.get('IMCONJUGATE')!, [const TextNode('3+4i')]);
      expect((result as TextValue).value, '3-4i');
    });

    test('negative imaginary', () {
      final result =
          eval(registry.get('IMCONJUGATE')!, [const TextNode('3-4i')]);
      expect((result as TextValue).value, '3+4i');
    });
  });

  group('IMSUM', () {
    test('two complex numbers', () {
      final result = eval(registry.get('IMSUM')!,
          [const TextNode('3+4i'), const TextNode('1+2i')]);
      expect((result as TextValue).value, '4+6i');
    });

    test('three complex numbers', () {
      final result = eval(registry.get('IMSUM')!,
          [const TextNode('1+i'), const TextNode('2+2i'), const TextNode('3+3i')]);
      expect((result as TextValue).value, '6+6i');
    });

    test('single number', () {
      final result =
          eval(registry.get('IMSUM')!, [const TextNode('3+4i')]);
      expect((result as TextValue).value, '3+4i');
    });
  });

  group('IMSUB', () {
    test('basic subtraction', () {
      final result = eval(registry.get('IMSUB')!,
          [const TextNode('13+4i'), const TextNode('5+3i')]);
      expect((result as TextValue).value, '8+i');
    });
  });

  group('IMPRODUCT', () {
    test('two complex numbers', () {
      final result = eval(registry.get('IMPRODUCT')!,
          [const TextNode('3+4i'), const TextNode('5-3i')]);
      // (3+4i)(5-3i) = 15-9i+20i-12i² = 15+11i+12 = 27+11i
      expect((result as TextValue).value, '27+11i');
    });

    test('single number', () {
      final result =
          eval(registry.get('IMPRODUCT')!, [const TextNode('3+4i')]);
      expect((result as TextValue).value, '3+4i');
    });
  });

  group('IMDIV', () {
    test('basic division', () {
      // (2+4i)/(1+2i) = (2+4i)(1-2i)/((1+2i)(1-2i)) = (2-4i+4i-8i²)/(1+4) = (10)/5 = 2
      final result = eval(registry.get('IMDIV')!,
          [const TextNode('-238+240i'), const TextNode('10+24i')]);
      // (-238+240i)/(10+24i): multiply by conj(10-24i)
      // num: (-238)(10)+(-238)(-24i)+(240i)(10)+(240i)(-24i)
      //    = -2380+5712i+2400i-5760i² = -2380+5760+8112i = 3380+8112i
      // den: 100+576 = 676
      // result: 5+12i
      expect((result as TextValue).value, '5+12i');
    });

    test('division by zero returns #NUM!', () {
      final result = eval(
          registry.get('IMDIV')!, [const TextNode('1+i'), const TextNode('0')]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('IMPOWER', () {
    test('square', () {
      // (2+3i)^2 = 4+12i-9 = -5+12i
      final result = eval(registry.get('IMPOWER')!,
          [const TextNode('2+3i'), const NumberNode(2)]);
      final parsed = result as TextValue;
      // Parse and check approximately
      expect(parsed.value, contains('-5'));
      expect(parsed.value, contains('12'));
    });

    test('zero power', () {
      final result = eval(registry.get('IMPOWER')!,
          [const TextNode('2+3i'), const NumberNode(0)]);
      expect((result as TextValue).value, '1');
    });
  });

  group('IMSQRT', () {
    test('basic sqrt', () {
      // sqrt(4) = 2
      final result = eval(registry.get('IMSQRT')!, [const TextNode('4')]);
      final parsed = result as TextValue;
      // Should be "2" (real only)
      expect(parsed.value, '2');
    });

    test('negative number', () {
      // sqrt(-1) = i
      final result = eval(registry.get('IMSQRT')!, [const TextNode('-1')]);
      final parsed = result as TextValue;
      expect(parsed.value, 'i');
    });
  });

  group('IMEXP', () {
    test('e^0 = 1', () {
      final result = eval(registry.get('IMEXP')!, [const TextNode('0')]);
      expect((result as TextValue).value, '1');
    });

    test('e^(pi*i) ≈ -1', () {
      final result = eval(
          registry.get('IMEXP')!, [const TextNode('0+3.14159265358979i')]);
      // e^(πi) = -1 (approximately)
      final text = (result as TextValue).value;
      // parse back
      expect(text, contains('-1'));
    });
  });

  group('IMLN', () {
    test('ln(1) = 0', () {
      final result = eval(registry.get('IMLN')!, [const TextNode('1')]);
      expect((result as TextValue).value, '0');
    });

    test('ln(0) returns #NUM!', () {
      final result = eval(registry.get('IMLN')!, [const TextNode('0')]);
      expect(result, const ErrorValue(FormulaError.num));
    });

    test('ln(e) ≈ 1', () {
      final result =
          eval(registry.get('IMLN')!, [const TextNode('2.71828182845905')]);
      final text = (result as TextValue).value;
      expect(text, contains('1'));
    });
  });

  group('IMLOG10', () {
    test('log10(10) = 1', () {
      final result = eval(registry.get('IMLOG10')!, [const TextNode('10')]);
      final text = (result as TextValue).value;
      expect(text, contains('1'));
    });

    test('log10(0) returns #NUM!', () {
      final result = eval(registry.get('IMLOG10')!, [const TextNode('0')]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('IMLOG2', () {
    test('log2(8) = 3', () {
      final result = eval(registry.get('IMLOG2')!, [const TextNode('8')]);
      final text = (result as TextValue).value;
      expect(text, contains('3'));
    });

    test('log2(0) returns #NUM!', () {
      final result = eval(registry.get('IMLOG2')!, [const TextNode('0')]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('IMSIN', () {
    test('sin(0) = 0', () {
      final result = eval(registry.get('IMSIN')!, [const TextNode('0')]);
      expect((result as TextValue).value, '0');
    });

    test('sin(pi/2) ≈ 1', () {
      final result = eval(
          registry.get('IMSIN')!, [const TextNode('1.5707963267949')]);
      final text = (result as TextValue).value;
      expect(text, contains('1'));
    });
  });

  group('IMCOS', () {
    test('cos(0) = 1', () {
      final result = eval(registry.get('IMCOS')!, [const TextNode('0')]);
      expect((result as TextValue).value, '1');
    });
  });

  group('IMTAN', () {
    test('tan(0) = 0', () {
      final result = eval(registry.get('IMTAN')!, [const TextNode('0')]);
      expect((result as TextValue).value, '0');
    });
  });

  group('IMSINH', () {
    test('sinh(0) = 0', () {
      final result = eval(registry.get('IMSINH')!, [const TextNode('0')]);
      expect((result as TextValue).value, '0');
    });
  });

  group('IMCOSH', () {
    test('cosh(0) = 1', () {
      final result = eval(registry.get('IMCOSH')!, [const TextNode('0')]);
      expect((result as TextValue).value, '1');
    });
  });

  group('IMSEC', () {
    test('sec(0) = 1', () {
      final result = eval(registry.get('IMSEC')!, [const TextNode('0')]);
      expect((result as TextValue).value, '1');
    });
  });

  group('IMSECH', () {
    test('sech(0) = 1', () {
      final result = eval(registry.get('IMSECH')!, [const TextNode('0')]);
      expect((result as TextValue).value, '1');
    });
  });

  group('IMCSC', () {
    test('csc(pi/2) = 1', () {
      final result = eval(
          registry.get('IMCSC')!, [const TextNode('1.5707963267949')]);
      final text = (result as TextValue).value;
      expect(text, contains('1'));
    });

    test('csc(0) returns #NUM!', () {
      final result = eval(registry.get('IMCSC')!, [const TextNode('0')]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('IMCSCH', () {
    test('csch(1) ≈ 0.8509', () {
      final result = eval(registry.get('IMCSCH')!, [const TextNode('1')]);
      final text = (result as TextValue).value;
      // Parse real part
      final val = double.tryParse(text);
      expect(val, isNotNull);
      expect(val!, closeTo(0.8509, 0.001));
    });

    test('csch(0) returns #NUM!', () {
      final result = eval(registry.get('IMCSCH')!, [const TextNode('0')]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('IMCOT', () {
    test('cot(pi/4) ≈ 1', () {
      final result = eval(
          registry.get('IMCOT')!, [const TextNode('0.785398163397448')]);
      final text = (result as TextValue).value;
      final val = double.tryParse(text);
      expect(val, isNotNull);
      expect(val!, closeTo(1.0, 0.001));
    });

    test('cot(0) returns #NUM!', () {
      final result = eval(registry.get('IMCOT')!, [const TextNode('0')]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  // =========================================================================
  // Wave 5 — CONVERT
  // =========================================================================

  group('CONVERT', () {
    // Weight
    test('kg to lbm', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('kg'), const TextNode('lbm')]);
      expect((result as NumberValue).value, closeTo(2.20462, 0.001));
    });

    test('lbm to kg', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('lbm'), const TextNode('kg')]);
      expect((result as NumberValue).value, closeTo(0.453592, 0.001));
    });

    // Distance
    test('miles to km', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('mi'), const TextNode('km')]);
      expect((result as NumberValue).value, closeTo(1.60934, 0.001));
    });

    test('meters to feet', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('m'), const TextNode('ft')]);
      expect((result as NumberValue).value, closeTo(3.28084, 0.001));
    });

    test('inches to cm', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('in'), const TextNode('cm')]);
      expect((result as NumberValue).value, closeTo(2.54, 0.001));
    });

    // Temperature
    test('Celsius to Fahrenheit', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(100), const TextNode('C'), const TextNode('F')]);
      expect((result as NumberValue).value, closeTo(212, 0.001));
    });

    test('Fahrenheit to Celsius', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(212), const TextNode('F'), const TextNode('C')]);
      expect((result as NumberValue).value, closeTo(100, 0.001));
    });

    test('Celsius to Kelvin', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(0), const TextNode('C'), const TextNode('K')]);
      expect((result as NumberValue).value, closeTo(273.15, 0.001));
    });

    test('Kelvin to Celsius', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(273.15), const TextNode('K'), const TextNode('C')]);
      expect((result as NumberValue).value, closeTo(0, 0.001));
    });

    test('Fahrenheit to Kelvin', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(32), const TextNode('F'), const TextNode('K')]);
      expect((result as NumberValue).value, closeTo(273.15, 0.001));
    });

    // Time
    test('hours to minutes', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('hr'), const TextNode('mn')]);
      expect((result as NumberValue).value, closeTo(60, 0.001));
    });

    test('days to hours', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('day'), const TextNode('hr')]);
      expect((result as NumberValue).value, closeTo(24, 0.001));
    });

    // Pressure
    test('atm to Pa', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('atm'), const TextNode('Pa')]);
      expect((result as NumberValue).value, closeTo(101325, 1));
    });

    // Energy
    test('calorie to joule', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('cal'), const TextNode('J')]);
      expect((result as NumberValue).value, closeTo(4.1868, 0.001));
    });

    // Power
    test('hp to watt', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('HP'), const TextNode('W')]);
      expect((result as NumberValue).value, closeTo(745.7, 0.1));
    });

    // Volume
    test('liters to gallons', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('l'), const TextNode('gal')]);
      expect((result as NumberValue).value, closeTo(0.264172, 0.001));
    });

    // Area
    test('hectare to acres', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('ha'), const TextNode('uk_acre')]);
      expect((result as NumberValue).value, closeTo(2.47105, 0.001));
    });

    // Information
    test('bytes to bits', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('byte'), const TextNode('bit')]);
      expect((result as NumberValue).value, closeTo(8, 0.001));
    });

    // Speed
    test('mph to km/h', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(60), const TextNode('mph'), const TextNode('km/h')]);
      expect((result as NumberValue).value, closeTo(96.5606, 0.01));
    });

    // Metric prefix
    test('km to m', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('km'), const TextNode('m')]);
      expect((result as NumberValue).value, closeTo(1000, 0.001));
    });

    test('cm to m', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(100), const TextNode('cm'), const TextNode('m')]);
      expect((result as NumberValue).value, closeTo(1, 0.001));
    });

    test('kg to g', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('kg'), const TextNode('g')]);
      expect((result as NumberValue).value, closeTo(1000, 0.001));
    });

    // Binary prefix
    test('kilobyte to byte', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('kibyte'), const TextNode('byte')]);
      expect((result as NumberValue).value, closeTo(1024, 0.001));
    });

    // Incompatible units
    test('incompatible units returns #N/A', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('m'), const TextNode('kg')]);
      expect(result, const ErrorValue(FormulaError.na));
    });

    // Unknown unit
    test('unknown unit returns #N/A', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('xyz'), const TextNode('m')]);
      expect(result, const ErrorValue(FormulaError.na));
    });

    // Force
    test('newton to dyn', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('N'), const TextNode('dyn')]);
      expect((result as NumberValue).value, closeTo(100000, 1));
    });

    // Magnetism
    test('tesla to gauss', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(1), const TextNode('T'), const TextNode('ga')]);
      expect((result as NumberValue).value, closeTo(10000, 1));
    });

    // Same unit
    test('same unit returns same value', () {
      final result = eval(registry.get('CONVERT')!,
          [const NumberNode(42), const TextNode('m'), const TextNode('m')]);
      expect((result as NumberValue).value, closeTo(42, 0.001));
    });
  });

  // =========================================================================
  // Registration
  // =========================================================================

  group('Registration', () {
    test('all 54 functions are registered', () {
      final reg = FunctionRegistry(registerBuiltIns: false);
      registerEngineeringFunctions(reg);
      expect(reg.names.length, 54);
    });

    test('total with built-ins is 363', () {
      final reg = FunctionRegistry();
      expect(reg.names.length, greaterThanOrEqualTo(363));
    });
  });
}
