import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/ast/operators.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/parser/formula_parser.dart';

void main() {
  late FormulaParser parser;

  setUp(() {
    parser = FormulaParser();
  });

  group('literals', () {
    test('integer', () {
      final node = parser.parse('42');
      expect(node, isA<NumberNode>());
      expect((node as NumberNode).value, 42);
    });

    test('decimal', () {
      final node = parser.parse('3.14');
      expect(node, isA<NumberNode>());
      expect((node as NumberNode).value, 3.14);
    });

    test('leading decimal', () {
      final node = parser.parse('.5');
      expect(node, isA<NumberNode>());
      expect((node as NumberNode).value, 0.5);
    });

    test('string', () {
      final node = parser.parse('"hello"');
      expect(node, isA<TextNode>());
      expect((node as TextNode).value, 'hello');
    });

    test('string with escaped quotes', () {
      final node = parser.parse('"say ""hi"""');
      expect(node, isA<TextNode>());
      expect((node as TextNode).value, 'say "hi"');
    });

    test('empty string', () {
      final node = parser.parse('""');
      expect(node, isA<TextNode>());
      expect((node as TextNode).value, '');
    });

    test('TRUE', () {
      final node = parser.parse('TRUE');
      expect(node, isA<BooleanNode>());
      expect((node as BooleanNode).value, true);
    });

    test('FALSE', () {
      final node = parser.parse('FALSE');
      expect(node, isA<BooleanNode>());
      expect((node as BooleanNode).value, false);
    });

    test('boolean is case-insensitive', () {
      expect(parser.parse('true'), isA<BooleanNode>());
      expect(parser.parse('True'), isA<BooleanNode>());
    });
  });

  group('error literals', () {
    test('#DIV/0!', () {
      final node = parser.parse('#DIV/0!');
      expect(node, isA<ErrorNode>());
      expect((node as ErrorNode).error, FormulaError.divZero);
    });

    test('#VALUE!', () {
      final node = parser.parse('#VALUE!');
      expect(node, isA<ErrorNode>());
      expect((node as ErrorNode).error, FormulaError.value);
    });

    test('#REF!', () {
      final node = parser.parse('#REF!');
      expect(node, isA<ErrorNode>());
      expect((node as ErrorNode).error, FormulaError.ref);
    });

    test('#NAME?', () {
      final node = parser.parse('#NAME?');
      expect(node, isA<ErrorNode>());
      expect((node as ErrorNode).error, FormulaError.name);
    });

    test('#NUM!', () {
      final node = parser.parse('#NUM!');
      expect(node, isA<ErrorNode>());
      expect((node as ErrorNode).error, FormulaError.num);
    });

    test('#N/A', () {
      final node = parser.parse('#N/A');
      expect(node, isA<ErrorNode>());
      expect((node as ErrorNode).error, FormulaError.na);
    });

    test('#NULL!', () {
      final node = parser.parse('#NULL!');
      expect(node, isA<ErrorNode>());
      expect((node as ErrorNode).error, FormulaError.null_);
    });
  });

  group('cell references', () {
    test('simple cell', () {
      final node = parser.parse('A1');
      expect(node, isA<CellRefNode>());
    });

    test('multi-letter column', () {
      final node = parser.parse('AA100');
      expect(node, isA<CellRefNode>());
    });

    test('absolute reference', () {
      final node = parser.parse(r'$A$1');
      expect(node, isA<CellRefNode>());
    });

    test('mixed absolute reference', () {
      final node = parser.parse(r'$A1');
      expect(node, isA<CellRefNode>());
      final node2 = parser.parse(r'A$1');
      expect(node2, isA<CellRefNode>());
    });
  });

  group('range references', () {
    test('simple range', () {
      final node = parser.parse('A1:B2');
      expect(node, isA<RangeRefNode>());
    });

    test('absolute range', () {
      final node = parser.parse(r'$A$1:$B$2');
      expect(node, isA<RangeRefNode>());
    });
  });

  group('formula prefix', () {
    test('with = prefix', () {
      final node = parser.parse('=42');
      expect(node, isA<NumberNode>());
      expect((node as NumberNode).value, 42);
    });

    test('with = and expression', () {
      final node = parser.parse('=1+2');
      expect(node, isA<BinaryOpNode>());
    });
  });

  group('arithmetic operators', () {
    test('addition', () {
      final node = parser.parse('1+2');
      expect(node, isA<BinaryOpNode>());
      final op = node as BinaryOpNode;
      expect(op.operator, BinaryOperator.add);
    });

    test('subtraction', () {
      final node = parser.parse('5-3');
      expect(node, isA<BinaryOpNode>());
      final op = node as BinaryOpNode;
      expect(op.operator, BinaryOperator.subtract);
    });

    test('multiplication', () {
      final node = parser.parse('2*3');
      expect(node, isA<BinaryOpNode>());
      final op = node as BinaryOpNode;
      expect(op.operator, BinaryOperator.multiply);
    });

    test('division', () {
      final node = parser.parse('10/2');
      expect(node, isA<BinaryOpNode>());
      final op = node as BinaryOpNode;
      expect(op.operator, BinaryOperator.divide);
    });

    test('power', () {
      final node = parser.parse('2^3');
      expect(node, isA<BinaryOpNode>());
      final op = node as BinaryOpNode;
      expect(op.operator, BinaryOperator.power);
    });
  });

  group('operator precedence', () {
    test('multiply before add: 1+2*3 = 1+(2*3)', () {
      final node = parser.parse('1+2*3') as BinaryOpNode;
      expect(node.operator, BinaryOperator.add);
      expect(node.left, isA<NumberNode>());
      expect(node.right, isA<BinaryOpNode>());
      expect((node.right as BinaryOpNode).operator, BinaryOperator.multiply);
    });

    test('power before multiply: 2*3^2 = 2*(3^2)', () {
      final node = parser.parse('2*3^2') as BinaryOpNode;
      expect(node.operator, BinaryOperator.multiply);
      expect(node.right, isA<BinaryOpNode>());
      expect((node.right as BinaryOpNode).operator, BinaryOperator.power);
    });

    test('parentheses override precedence: (1+2)*3', () {
      final node = parser.parse('(1+2)*3') as BinaryOpNode;
      expect(node.operator, BinaryOperator.multiply);
      expect(node.left, isA<ParenthesizedNode>());
    });
  });

  group('comparison operators', () {
    test('equal', () {
      final node = parser.parse('A1=5') as BinaryOpNode;
      expect(node.operator, BinaryOperator.equal);
    });

    test('not equal', () {
      final node = parser.parse('A1<>5') as BinaryOpNode;
      expect(node.operator, BinaryOperator.notEqual);
    });

    test('less than', () {
      final node = parser.parse('A1<5') as BinaryOpNode;
      expect(node.operator, BinaryOperator.lessThan);
    });

    test('greater than', () {
      final node = parser.parse('A1>5') as BinaryOpNode;
      expect(node.operator, BinaryOperator.greaterThan);
    });

    test('less equal', () {
      final node = parser.parse('A1<=5') as BinaryOpNode;
      expect(node.operator, BinaryOperator.lessEqual);
    });

    test('greater equal', () {
      final node = parser.parse('A1>=5') as BinaryOpNode;
      expect(node.operator, BinaryOperator.greaterEqual);
    });
  });

  group('concatenation', () {
    test('& operator', () {
      final node = parser.parse('"hello"&" world"') as BinaryOpNode;
      expect(node.operator, BinaryOperator.concat);
    });
  });

  group('unary operators', () {
    test('negation', () {
      final node = parser.parse('-5');
      expect(node, isA<UnaryOpNode>());
      final op = node as UnaryOpNode;
      expect(op.operator, UnaryOperator.negate);
    });

    test('percentage', () {
      final node = parser.parse('50%');
      expect(node, isA<UnaryOpNode>());
      final op = node as UnaryOpNode;
      expect(op.operator, UnaryOperator.percent);
    });
  });

  group('function calls', () {
    test('no arguments', () {
      final node = parser.parse('TRUE()');
      expect(node, isA<FunctionCallNode>());
      final fn = node as FunctionCallNode;
      expect(fn.name, 'TRUE');
      expect(fn.arguments, isEmpty);
    });

    test('single argument', () {
      final node = parser.parse('ABS(5)');
      expect(node, isA<FunctionCallNode>());
      final fn = node as FunctionCallNode;
      expect(fn.name, 'ABS');
      expect(fn.arguments.length, 1);
    });

    test('multiple arguments with comma', () {
      final node = parser.parse('IF(A1>0,1,2)');
      expect(node, isA<FunctionCallNode>());
      final fn = node as FunctionCallNode;
      expect(fn.name, 'IF');
      expect(fn.arguments.length, 3);
    });

    test('nested function calls', () {
      final node = parser.parse('SUM(ABS(A1),5)');
      expect(node, isA<FunctionCallNode>());
      final fn = node as FunctionCallNode;
      expect(fn.name, 'SUM');
      expect(fn.arguments[0], isA<FunctionCallNode>());
    });

    test('range argument', () {
      final node = parser.parse('SUM(A1:A10)');
      expect(node, isA<FunctionCallNode>());
      final fn = node as FunctionCallNode;
      expect(fn.arguments[0], isA<RangeRefNode>());
    });

    test('dotted function name', () {
      final node = parser.parse('MODE.SNGL(1,2,3)');
      expect(node, isA<FunctionCallNode>());
      final fn = node as FunctionCallNode;
      expect(fn.name, 'MODE.SNGL');
      expect(fn.arguments.length, 3);
    });

    test('dotted function name with two dots', () {
      final node = parser.parse('CEILING.MATH(5)');
      expect(node, isA<FunctionCallNode>());
      final fn = node as FunctionCallNode;
      expect(fn.name, 'CEILING.MATH');
      expect(fn.arguments.length, 1);
    });
  });

  group('bare identifiers', () {
    test('single letter x parses as NameNode', () {
      final node = parser.parse('x');
      expect(node, isA<NameNode>());
      expect((node as NameNode).name, 'x');
    });

    test('multi-letter count parses as NameNode', () {
      final node = parser.parse('count');
      expect(node, isA<NameNode>());
      expect((node as NameNode).name, 'count');
    });

    test('underscore my_var parses as NameNode', () {
      final node = parser.parse('my_var');
      expect(node, isA<NameNode>());
      expect((node as NameNode).name, 'my_var');
    });

    test('cell ref A1 still parses as CellRefNode', () {
      final node = parser.parse('A1');
      expect(node, isA<CellRefNode>());
    });

    test('cell ref AA100 still parses as CellRefNode', () {
      final node = parser.parse('AA100');
      expect(node, isA<CellRefNode>());
    });

    test('TRUE still parses as BooleanNode', () {
      final node = parser.parse('TRUE');
      expect(node, isA<BooleanNode>());
    });

    test('FALSE still parses as BooleanNode', () {
      final node = parser.parse('FALSE');
      expect(node, isA<BooleanNode>());
    });

    test('true (lowercase) still parses as BooleanNode', () {
      final node = parser.parse('true');
      expect(node, isA<BooleanNode>());
    });

    test('SUM(1) still parses as FunctionCallNode', () {
      final node = parser.parse('SUM(1)');
      expect(node, isA<FunctionCallNode>());
    });

    test('identifier in expression x+1', () {
      final node = parser.parse('x+1');
      expect(node, isA<BinaryOpNode>());
      final op = node as BinaryOpNode;
      expect(op.left, isA<NameNode>());
      expect(op.operator, BinaryOperator.add);
    });

    test('LAMBDA(x, x+1) parses with NameNode args', () {
      final node = parser.parse('LAMBDA(x, x+1)');
      expect(node, isA<FunctionCallNode>());
      final fn = node as FunctionCallNode;
      expect(fn.name, 'LAMBDA');
      expect(fn.arguments[0], isA<NameNode>());
      expect(fn.arguments[1], isA<BinaryOpNode>());
    });
  });

  group('call expressions', () {
    test('LAMBDA(x,x+1)(5) parses as CallExpressionNode', () {
      final node = parser.parse('LAMBDA(x,x+1)(5)');
      expect(node, isA<CallExpressionNode>());
      final call = node as CallExpressionNode;
      expect(call.function, isA<FunctionCallNode>());
      expect((call.function as FunctionCallNode).name, 'LAMBDA');
      expect(call.arguments.length, 1);
      expect(call.arguments[0], isA<NumberNode>());
    });

    test('chained call parses as nested CallExpressionNode', () {
      final node = parser.parse('LAMBDA(x,LAMBDA(y,x+y))(1)(2)');
      expect(node, isA<CallExpressionNode>());
      final outer = node as CallExpressionNode;
      // The inner part should also be a CallExpressionNode
      expect(outer.function, isA<CallExpressionNode>());
      expect(outer.arguments.length, 1);
    });

    test('parenthesized expression followed by (args) parses as call', () {
      final node = parser.parse('(LAMBDA(x,x+1))(5)');
      expect(node, isA<CallExpressionNode>());
    });
  });

  group('complex expressions', () {
    test('=SUM(A1:A10)*2+1', () {
      final node = parser.parse('=SUM(A1:A10)*2+1');
      expect(node, isA<BinaryOpNode>());
      final add = node as BinaryOpNode;
      expect(add.operator, BinaryOperator.add);
      expect(add.left, isA<BinaryOpNode>());
      final mul = add.left as BinaryOpNode;
      expect(mul.operator, BinaryOperator.multiply);
      expect(mul.left, isA<FunctionCallNode>());
    });

    test('=IF(A1>0,"positive","negative")', () {
      final node = parser.parse('=IF(A1>0,"positive","negative")');
      expect(node, isA<FunctionCallNode>());
      final fn = node as FunctionCallNode;
      expect(fn.name, 'IF');
      expect(fn.arguments.length, 3);
      expect(fn.arguments[0], isA<BinaryOpNode>());
      expect(fn.arguments[1], isA<TextNode>());
      expect(fn.arguments[2], isA<TextNode>());
    });
  });

  group('tryParse and isValid', () {
    test('tryParse returns node on success', () {
      expect(parser.tryParse('42'), isA<NumberNode>());
    });

    test('tryParse returns null on failure', () {
      expect(parser.tryParse('='), isNull);
    });

    test('isValid', () {
      expect(parser.isValid('42'), isTrue);
      expect(parser.isValid('=SUM(A1:A10)'), isTrue);
      expect(parser.isValid('='), isFalse);
    });
  });

  group('FormulaParseException', () {
    test('thrown on invalid formula', () {
      expect(
        () => parser.parse('='),
        throwsA(isA<FormulaParseException>()),
      );
    });

    test('contains position and formula', () {
      try {
        parser.parse('=');
        fail('Should have thrown');
      } on FormulaParseException catch (e) {
        expect(e.formula, '=');
        expect(e.position, greaterThanOrEqualTo(0));
        expect(e.message, isNotEmpty);
        expect(e.toString(), contains('FormulaParseException'));
      }
    });
  });

  group('descriptive error messages', () {
    test('truncated formula shows unexpected end', () {
      try {
        parser.parse('=SUM(');
        fail('Should have thrown');
      } on FormulaParseException catch (e) {
        expect(
          e.message.toLowerCase(),
          anyOf(
            contains('unexpected end'),
            contains('end of'),
            contains('unexpected character'),
          ),
        );
      }
    });

    test('unmatched opening parenthesis', () {
      try {
        parser.parse('=(1+2');
        fail('Should have thrown');
      } on FormulaParseException catch (e) {
        expect(
          e.message.toLowerCase(),
          anyOf(
            contains('parenthesis'),
            contains('expected'),
            contains('end of'),
          ),
        );
      }
    });

    test('extra closing parenthesis', () {
      try {
        parser.parse('=1+2)');
        fail('Should have thrown');
      } on FormulaParseException catch (e) {
        expect(e.position, greaterThan(0));
        expect(e.message.toLowerCase(), contains('unexpected'));
      }
    });

    test('incomplete expression', () {
      try {
        parser.parse('=1+');
        fail('Should have thrown');
      } on FormulaParseException catch (e) {
        expect(e.position, greaterThan(0));
      }
    });

    test('toString includes visual pointer', () {
      try {
        parser.parse('=1+2)');
        fail('Should have thrown');
      } on FormulaParseException catch (e) {
        final str = e.toString();
        expect(str, contains('^'));
        expect(str, contains('=1+2)'));
      }
    });
  });
}
