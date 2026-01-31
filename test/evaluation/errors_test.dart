import 'package:test/test.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';

void main() {
  group('FormulaError', () {
    test('has correct error codes', () {
      expect(FormulaError.divZero.code, '#DIV/0!');
      expect(FormulaError.value.code, '#VALUE!');
      expect(FormulaError.ref.code, '#REF!');
      expect(FormulaError.name.code, '#NAME?');
      expect(FormulaError.num.code, '#NUM!');
      expect(FormulaError.na.code, '#N/A');
      expect(FormulaError.null_.code, '#NULL!');
      expect(FormulaError.calc.code, '#CALC!');
      expect(FormulaError.circular.code, '#CIRCULAR!');
    });

    test('toString returns the error code', () {
      for (final error in FormulaError.values) {
        expect(error.toString(), error.code);
      }
    });

    test('contains all expected error types', () {
      expect(FormulaError.values.length, 9);
    });
  });
}
