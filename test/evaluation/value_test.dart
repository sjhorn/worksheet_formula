import 'package:test/test.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';

void main() {
  group('NumberValue', () {
    test('toNumber returns the value', () {
      expect(const NumberValue(42).toNumber(), 42);
      expect(const NumberValue(3.14).toNumber(), 3.14);
      expect(const NumberValue(-1).toNumber(), -1);
      expect(const NumberValue(0).toNumber(), 0);
    });

    test('toText returns string representation', () {
      expect(const NumberValue(42).toText(), '42');
      expect(const NumberValue(3.14).toText(), '3.14');
    });

    test('toBool returns false for zero, true otherwise', () {
      expect(const NumberValue(0).toBool(), false);
      expect(const NumberValue(1).toBool(), true);
      expect(const NumberValue(-1).toBool(), true);
    });

    test('isTruthy matches toBool', () {
      expect(const NumberValue(0).isTruthy, false);
      expect(const NumberValue(1).isTruthy, true);
    });

    test('isError is false', () {
      expect(const NumberValue(42).isError, false);
    });

    test('equality', () {
      expect(const NumberValue(42), const NumberValue(42));
      expect(const NumberValue(42), isNot(const NumberValue(43)));
      expect(const NumberValue(42).hashCode, const NumberValue(42).hashCode);
    });

    test('toString', () {
      expect(const NumberValue(42).toString(), 'NumberValue(42)');
    });

    test('factory constructor', () {
      const value = FormulaValue.number(42);
      expect(value, isA<NumberValue>());
      expect((value as NumberValue).value, 42);
    });
  });

  group('TextValue', () {
    test('toNumber parses numeric strings', () {
      expect(const TextValue('42').toNumber(), 42);
      expect(const TextValue('3.14').toNumber(), 3.14);
    });

    test('toNumber returns null for non-numeric strings', () {
      expect(const TextValue('hello').toNumber(), null);
      expect(const TextValue('').toNumber(), null);
    });

    test('toText returns the value', () {
      expect(const TextValue('hello').toText(), 'hello');
      expect(const TextValue('').toText(), '');
    });

    test('toBool returns false for empty, true otherwise', () {
      expect(const TextValue('').toBool(), false);
      expect(const TextValue('hello').toBool(), true);
    });

    test('isTruthy matches toBool', () {
      expect(const TextValue('').isTruthy, false);
      expect(const TextValue('hello').isTruthy, true);
    });

    test('isError is false', () {
      expect(const TextValue('hello').isError, false);
    });

    test('equality', () {
      expect(const TextValue('a'), const TextValue('a'));
      expect(const TextValue('a'), isNot(const TextValue('b')));
    });

    test('toString', () {
      expect(const TextValue('hello').toString(), 'TextValue("hello")');
    });

    test('factory constructor', () {
      const value = FormulaValue.text('hello');
      expect(value, isA<TextValue>());
      expect((value as TextValue).value, 'hello');
    });
  });

  group('BooleanValue', () {
    test('toNumber returns 1 for true, 0 for false', () {
      expect(const BooleanValue(true).toNumber(), 1);
      expect(const BooleanValue(false).toNumber(), 0);
    });

    test('toText returns TRUE/FALSE', () {
      expect(const BooleanValue(true).toText(), 'TRUE');
      expect(const BooleanValue(false).toText(), 'FALSE');
    });

    test('toBool returns the value', () {
      expect(const BooleanValue(true).toBool(), true);
      expect(const BooleanValue(false).toBool(), false);
    });

    test('isTruthy matches the value', () {
      expect(const BooleanValue(true).isTruthy, true);
      expect(const BooleanValue(false).isTruthy, false);
    });

    test('isError is false', () {
      expect(const BooleanValue(true).isError, false);
    });

    test('equality', () {
      expect(const BooleanValue(true), const BooleanValue(true));
      expect(const BooleanValue(true), isNot(const BooleanValue(false)));
    });

    test('toString', () {
      expect(const BooleanValue(true).toString(), 'BooleanValue(true)');
      expect(const BooleanValue(false).toString(), 'BooleanValue(false)');
    });

    test('factory constructor', () {
      const value = FormulaValue.boolean(true);
      expect(value, isA<BooleanValue>());
      expect((value as BooleanValue).value, true);
    });
  });

  group('ErrorValue', () {
    test('toNumber returns null', () {
      expect(const ErrorValue(FormulaError.divZero).toNumber(), null);
    });

    test('toText returns the error code', () {
      expect(const ErrorValue(FormulaError.divZero).toText(), '#DIV/0!');
      expect(const ErrorValue(FormulaError.ref).toText(), '#REF!');
    });

    test('toBool returns false', () {
      expect(const ErrorValue(FormulaError.value).toBool(), false);
    });

    test('isTruthy is false', () {
      expect(const ErrorValue(FormulaError.value).isTruthy, false);
    });

    test('isError is true', () {
      expect(const ErrorValue(FormulaError.divZero).isError, true);
    });

    test('equality', () {
      expect(
        const ErrorValue(FormulaError.divZero),
        const ErrorValue(FormulaError.divZero),
      );
      expect(
        const ErrorValue(FormulaError.divZero),
        isNot(const ErrorValue(FormulaError.ref)),
      );
    });

    test('toString', () {
      expect(
        const ErrorValue(FormulaError.divZero).toString(),
        'ErrorValue(#DIV/0!)',
      );
    });

    test('factory constructor', () {
      const value = FormulaValue.error(FormulaError.na);
      expect(value, isA<ErrorValue>());
      expect((value as ErrorValue).error, FormulaError.na);
    });
  });

  group('EmptyValue', () {
    test('toNumber returns 0', () {
      expect(const EmptyValue().toNumber(), 0);
    });

    test('toText returns empty string', () {
      expect(const EmptyValue().toText(), '');
    });

    test('toBool returns false', () {
      expect(const EmptyValue().toBool(), false);
    });

    test('isTruthy is false', () {
      expect(const EmptyValue().isTruthy, false);
    });

    test('isError is false', () {
      expect(const EmptyValue().isError, false);
    });

    test('toString', () {
      expect(const EmptyValue().toString(), 'EmptyValue()');
    });

    test('factory constructor', () {
      const value = FormulaValue.empty();
      expect(value, isA<EmptyValue>());
    });
  });

  group('RangeValue', () {
    test('rowCount and columnCount', () {
      const range = RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
      ]);
      expect(range.rowCount, 2);
      expect(range.columnCount, 2);
    });

    test('empty range dimensions', () {
      const range = RangeValue([]);
      expect(range.rowCount, 0);
      expect(range.columnCount, 0);
    });

    test('flat returns all values', () {
      const range = RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
      ]);
      expect(
        range.flat.toList(),
        [
          const NumberValue(1),
          const NumberValue(2),
          const NumberValue(3),
          const NumberValue(4),
        ],
      );
    });

    test('numbers returns only numeric values', () {
      const range = RangeValue([
        [NumberValue(1), TextValue('hello')],
        [NumberValue(3), BooleanValue(true)],
      ]);
      expect(range.numbers.toList(), [1, 3]);
    });

    test('toNumber for single cell range', () {
      const range = RangeValue([
        [NumberValue(42)],
      ]);
      expect(range.toNumber(), 42);
    });

    test('toNumber returns null for multi-cell range', () {
      const range = RangeValue([
        [NumberValue(1), NumberValue(2)],
      ]);
      expect(range.toNumber(), null);
    });

    test('toText joins with commas and semicolons', () {
      const range = RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
      ]);
      expect(range.toText(), '1,2;3,4');
    });

    test('isTruthy is true when non-empty', () {
      const empty = RangeValue([]);
      const nonEmpty = RangeValue([
        [NumberValue(1)],
      ]);
      expect(empty.isTruthy, false);
      expect(nonEmpty.isTruthy, true);
    });

    test('isError is false', () {
      const range = RangeValue([]);
      expect(range.isError, false);
    });

    test('toString', () {
      const range = RangeValue([
        [NumberValue(1), NumberValue(2)],
        [NumberValue(3), NumberValue(4)],
      ]);
      expect(range.toString(), 'RangeValue(2x2)');
    });

    test('factory constructor', () {
      const value = FormulaValue.range([
        [NumberValue(1)],
      ]);
      expect(value, isA<RangeValue>());
    });
  });

  group('FunctionValue', () {
    test('toNumber returns null', () {
      final fv = FunctionValue(['x'], (args) => const NumberValue(0));
      expect(fv.toNumber(), null);
    });

    test('toText returns #LAMBDA', () {
      final fv = FunctionValue(['x'], (args) => const NumberValue(0));
      expect(fv.toText(), '#LAMBDA');
    });

    test('toBool returns false', () {
      final fv = FunctionValue(['x'], (args) => const NumberValue(0));
      expect(fv.toBool(), false);
    });

    test('isTruthy is false', () {
      final fv = FunctionValue(['x'], (args) => const NumberValue(0));
      expect(fv.isTruthy, false);
    });

    test('isError is false', () {
      final fv = FunctionValue(['x'], (args) => const NumberValue(0));
      expect(fv.isError, false);
    });

    test('toString', () {
      final fv = FunctionValue(['x', 'y'], (args) => const NumberValue(0));
      expect(fv.toString(), 'FunctionValue(x, y)');
    });

    test('invoke calls the function', () {
      final fv = FunctionValue(['x'], (args) => FormulaValue.number(args[0].toNumber()! * 2));
      expect(fv.invoke([const NumberValue(5)]), const NumberValue(10));
    });
  });

  group('OmittedValue', () {
    test('toNumber returns 0', () {
      expect(const OmittedValue().toNumber(), 0);
    });

    test('toText returns empty string', () {
      expect(const OmittedValue().toText(), '');
    });

    test('toBool returns false', () {
      expect(const OmittedValue().toBool(), false);
    });

    test('isTruthy is false', () {
      expect(const OmittedValue().isTruthy, false);
    });

    test('isError is false', () {
      expect(const OmittedValue().isError, false);
    });

    test('toString', () {
      expect(const OmittedValue().toString(), 'OmittedValue()');
    });
  });

  group('FormulaValue sealed class', () {
    test('pattern matching works on all subtypes', () {
      final values = <FormulaValue>[
        const NumberValue(1),
        const TextValue('hi'),
        const BooleanValue(true),
        const ErrorValue(FormulaError.na),
        const EmptyValue(),
        const RangeValue([]),
        FunctionValue([], (args) => const NumberValue(0)),
        const OmittedValue(),
      ];

      final types = values.map((v) => switch (v) {
        NumberValue() => 'number',
        TextValue() => 'text',
        BooleanValue() => 'boolean',
        ErrorValue() => 'error',
        EmptyValue() => 'empty',
        RangeValue() => 'range',
        FunctionValue() => 'function',
        OmittedValue() => 'omitted',
      }).toList();

      expect(types, [
        'number', 'text', 'boolean', 'error', 'empty', 'range',
        'function', 'omitted',
      ]);
    });
  });
}
