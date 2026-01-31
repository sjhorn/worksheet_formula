/// Excel-compatible formula errors.
enum FormulaError {
  /// #DIV/0! - Division by zero
  divZero('#DIV/0!'),

  /// #VALUE! - Wrong type of argument
  value('#VALUE!'),

  /// #REF! - Invalid cell reference
  ref('#REF!'),

  /// #NAME? - Unrecognized formula name
  name('#NAME?'),

  /// #NUM! - Invalid numeric value
  num('#NUM!'),

  /// #N/A - Value not available
  na('#N/A'),

  /// #NULL! - Incorrect range operator
  null_('#NULL!'),

  /// #CALC! - Calculation error
  calc('#CALC!'),

  /// #CIRCULAR! - Circular reference detected
  circular('#CIRCULAR!');

  final String code;
  const FormulaError(this.code);

  @override
  String toString() => code;
}
