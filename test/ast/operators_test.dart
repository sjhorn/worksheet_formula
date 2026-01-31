import 'package:test/test.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/ast/operators.dart';

void main() {
  group('BinaryOperator', () {
    group('arithmetic', () {
      test('add', () {
        expect(
          BinaryOperator.add.apply(
            const NumberValue(2),
            const NumberValue(3),
          ),
          const NumberValue(5),
        );
      });

      test('subtract', () {
        expect(
          BinaryOperator.subtract.apply(
            const NumberValue(10),
            const NumberValue(3),
          ),
          const NumberValue(7),
        );
      });

      test('multiply', () {
        expect(
          BinaryOperator.multiply.apply(
            const NumberValue(4),
            const NumberValue(5),
          ),
          const NumberValue(20),
        );
      });

      test('divide', () {
        expect(
          BinaryOperator.divide.apply(
            const NumberValue(10),
            const NumberValue(4),
          ),
          const NumberValue(2.5),
        );
      });

      test('divide by zero returns #DIV/0!', () {
        expect(
          BinaryOperator.divide.apply(
            const NumberValue(10),
            const NumberValue(0),
          ),
          const ErrorValue(FormulaError.divZero),
        );
      });

      test('power', () {
        expect(
          BinaryOperator.power.apply(
            const NumberValue(2),
            const NumberValue(3),
          ),
          const NumberValue(8),
        );
      });

      test('arithmetic with non-numeric left returns #VALUE!', () {
        expect(
          BinaryOperator.add.apply(
            const TextValue('hello'),
            const NumberValue(1),
          ),
          const ErrorValue(FormulaError.value),
        );
      });

      test('arithmetic with non-numeric right returns #VALUE!', () {
        expect(
          BinaryOperator.multiply.apply(
            const NumberValue(1),
            const TextValue('hello'),
          ),
          const ErrorValue(FormulaError.value),
        );
      });

      test('divide with non-numeric returns #VALUE!', () {
        expect(
          BinaryOperator.divide.apply(
            const TextValue('a'),
            const NumberValue(1),
          ),
          const ErrorValue(FormulaError.value),
        );
        expect(
          BinaryOperator.divide.apply(
            const NumberValue(1),
            const TextValue('a'),
          ),
          const ErrorValue(FormulaError.value),
        );
      });

      test('arithmetic with numeric text coerces', () {
        expect(
          BinaryOperator.add.apply(
            const TextValue('5'),
            const NumberValue(3),
          ),
          const NumberValue(8),
        );
      });

      test('arithmetic with boolean coerces', () {
        expect(
          BinaryOperator.add.apply(
            const BooleanValue(true),
            const NumberValue(3),
          ),
          const NumberValue(4),
        );
      });
    });

    group('comparison', () {
      test('equal numbers', () {
        expect(
          BinaryOperator.equal.apply(
            const NumberValue(5),
            const NumberValue(5),
          ),
          const BooleanValue(true),
        );
        expect(
          BinaryOperator.equal.apply(
            const NumberValue(5),
            const NumberValue(6),
          ),
          const BooleanValue(false),
        );
      });

      test('notEqual', () {
        expect(
          BinaryOperator.notEqual.apply(
            const NumberValue(5),
            const NumberValue(6),
          ),
          const BooleanValue(true),
        );
      });

      test('lessThan', () {
        expect(
          BinaryOperator.lessThan.apply(
            const NumberValue(3),
            const NumberValue(5),
          ),
          const BooleanValue(true),
        );
        expect(
          BinaryOperator.lessThan.apply(
            const NumberValue(5),
            const NumberValue(3),
          ),
          const BooleanValue(false),
        );
      });

      test('greaterThan', () {
        expect(
          BinaryOperator.greaterThan.apply(
            const NumberValue(5),
            const NumberValue(3),
          ),
          const BooleanValue(true),
        );
      });

      test('lessEqual', () {
        expect(
          BinaryOperator.lessEqual.apply(
            const NumberValue(3),
            const NumberValue(3),
          ),
          const BooleanValue(true),
        );
        expect(
          BinaryOperator.lessEqual.apply(
            const NumberValue(3),
            const NumberValue(5),
          ),
          const BooleanValue(true),
        );
      });

      test('greaterEqual', () {
        expect(
          BinaryOperator.greaterEqual.apply(
            const NumberValue(5),
            const NumberValue(5),
          ),
          const BooleanValue(true),
        );
        expect(
          BinaryOperator.greaterEqual.apply(
            const NumberValue(3),
            const NumberValue(5),
          ),
          const BooleanValue(false),
        );
      });

      test('text comparison is case-insensitive', () {
        expect(
          BinaryOperator.equal.apply(
            const TextValue('Hello'),
            const TextValue('hello'),
          ),
          const BooleanValue(true),
        );
      });

      test('text ordering', () {
        expect(
          BinaryOperator.lessThan.apply(
            const TextValue('apple'),
            const TextValue('banana'),
          ),
          const BooleanValue(true),
        );
      });

      test('boolean comparison', () {
        expect(
          BinaryOperator.equal.apply(
            const BooleanValue(true),
            const BooleanValue(true),
          ),
          const BooleanValue(true),
        );
        expect(
          BinaryOperator.greaterThan.apply(
            const BooleanValue(true),
            const BooleanValue(false),
          ),
          const BooleanValue(true),
        );
      });

      test('mixed type comparison falls back to numeric then text', () {
        // Number vs numeric text
        expect(
          BinaryOperator.equal.apply(
            const NumberValue(5),
            const TextValue('5'),
          ),
          const BooleanValue(true),
        );
      });
    });

    group('concatenation', () {
      test('concat joins text', () {
        expect(
          BinaryOperator.concat.apply(
            const TextValue('hello'),
            const TextValue(' world'),
          ),
          const TextValue('hello world'),
        );
      });

      test('concat coerces numbers to text', () {
        expect(
          BinaryOperator.concat.apply(
            const TextValue('value: '),
            const NumberValue(42),
          ),
          const TextValue('value: 42'),
        );
      });
    });

    group('error propagation', () {
      test('left error propagates', () {
        expect(
          BinaryOperator.add.apply(
            const ErrorValue(FormulaError.ref),
            const NumberValue(1),
          ),
          const ErrorValue(FormulaError.ref),
        );
      });

      test('right error propagates', () {
        expect(
          BinaryOperator.add.apply(
            const NumberValue(1),
            const ErrorValue(FormulaError.na),
          ),
          const ErrorValue(FormulaError.na),
        );
      });
    });

    group('metadata', () {
      test('symbols are correct', () {
        expect(BinaryOperator.add.symbol, '+');
        expect(BinaryOperator.subtract.symbol, '-');
        expect(BinaryOperator.multiply.symbol, '*');
        expect(BinaryOperator.divide.symbol, '/');
        expect(BinaryOperator.power.symbol, '^');
        expect(BinaryOperator.equal.symbol, '=');
        expect(BinaryOperator.notEqual.symbol, '<>');
        expect(BinaryOperator.lessThan.symbol, '<');
        expect(BinaryOperator.greaterThan.symbol, '>');
        expect(BinaryOperator.lessEqual.symbol, '<=');
        expect(BinaryOperator.greaterEqual.symbol, '>=');
        expect(BinaryOperator.concat.symbol, '&');
      });

      test('precedences follow arithmetic rules', () {
        // comparison < concat < add/sub < mul/div < power
        expect(
          BinaryOperator.equal.precedence,
          lessThan(BinaryOperator.concat.precedence),
        );
        expect(
          BinaryOperator.concat.precedence,
          lessThan(BinaryOperator.add.precedence),
        );
        expect(
          BinaryOperator.add.precedence,
          equals(BinaryOperator.subtract.precedence),
        );
        expect(
          BinaryOperator.add.precedence,
          lessThan(BinaryOperator.multiply.precedence),
        );
        expect(
          BinaryOperator.multiply.precedence,
          equals(BinaryOperator.divide.precedence),
        );
        expect(
          BinaryOperator.multiply.precedence,
          lessThan(BinaryOperator.power.precedence),
        );
      });
    });
  });

  group('UnaryOperator', () {
    test('negate', () {
      expect(
        UnaryOperator.negate.apply(const NumberValue(5)),
        const NumberValue(-5),
      );
      expect(
        UnaryOperator.negate.apply(const NumberValue(-3)),
        const NumberValue(3),
      );
    });

    test('negate non-numeric returns #VALUE!', () {
      expect(
        UnaryOperator.negate.apply(const TextValue('hello')),
        const ErrorValue(FormulaError.value),
      );
    });

    test('positive returns operand unchanged', () {
      const value = NumberValue(5);
      expect(UnaryOperator.positive.apply(value), value);
    });

    test('percent divides by 100', () {
      expect(
        UnaryOperator.percent.apply(const NumberValue(50)),
        const NumberValue(0.5),
      );
    });

    test('percent non-numeric returns #VALUE!', () {
      expect(
        UnaryOperator.percent.apply(const TextValue('hello')),
        const ErrorValue(FormulaError.value),
      );
    });

    test('error propagates through unary', () {
      expect(
        UnaryOperator.negate.apply(const ErrorValue(FormulaError.ref)),
        const ErrorValue(FormulaError.ref),
      );
    });

    test('symbols are correct', () {
      expect(UnaryOperator.negate.symbol, '-');
      expect(UnaryOperator.positive.symbol, '+');
      expect(UnaryOperator.percent.symbol, '%');
    });
  });
}
