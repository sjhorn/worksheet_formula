import 'errors.dart';

/// Represents any value that can result from formula evaluation.
sealed class FormulaValue {
  const FormulaValue();

  const factory FormulaValue.number(num value) = NumberValue;
  const factory FormulaValue.text(String value) = TextValue;
  const factory FormulaValue.boolean(bool value) = BooleanValue;
  const factory FormulaValue.error(FormulaError error) = ErrorValue;
  const factory FormulaValue.range(List<List<FormulaValue>> values) =
      RangeValue;
  const factory FormulaValue.empty() = EmptyValue;

  /// Convert to number (for arithmetic operations).
  num? toNumber();

  /// Convert to string (for text operations).
  String toText();

  /// Convert to boolean (for logical operations).
  bool toBool();

  /// Is this value "truthy" for IF conditions?
  bool get isTruthy;

  /// Is this an error value?
  bool get isError => this is ErrorValue;
}

class NumberValue extends FormulaValue {
  final num value;
  const NumberValue(this.value);

  @override
  num? toNumber() => value;

  @override
  String toText() => value.toString();

  @override
  bool toBool() => value != 0;

  @override
  bool get isTruthy => value != 0;

  @override
  bool operator ==(Object other) =>
      other is NumberValue && other.value == value;

  @override
  int get hashCode => value.hashCode;

  @override
  String toString() => 'NumberValue($value)';
}

class TextValue extends FormulaValue {
  final String value;
  const TextValue(this.value);

  @override
  num? toNumber() => num.tryParse(value);

  @override
  String toText() => value;

  @override
  bool toBool() => value.isNotEmpty;

  @override
  bool get isTruthy => value.isNotEmpty;

  @override
  bool operator ==(Object other) =>
      other is TextValue && other.value == value;

  @override
  int get hashCode => value.hashCode;

  @override
  String toString() => 'TextValue("$value")';
}

class BooleanValue extends FormulaValue {
  final bool value;
  const BooleanValue(this.value);

  @override
  num? toNumber() => value ? 1 : 0;

  @override
  String toText() => value ? 'TRUE' : 'FALSE';

  @override
  bool toBool() => value;

  @override
  bool get isTruthy => value;

  @override
  bool operator ==(Object other) =>
      other is BooleanValue && other.value == value;

  @override
  int get hashCode => value.hashCode;

  @override
  String toString() => 'BooleanValue($value)';
}

class ErrorValue extends FormulaValue {
  final FormulaError error;
  const ErrorValue(this.error);

  @override
  num? toNumber() => null;

  @override
  String toText() => error.code;

  @override
  bool toBool() => false;

  @override
  bool get isTruthy => false;

  @override
  bool operator ==(Object other) =>
      other is ErrorValue && other.error == error;

  @override
  int get hashCode => error.hashCode;

  @override
  String toString() => 'ErrorValue(${error.code})';
}

class RangeValue extends FormulaValue {
  final List<List<FormulaValue>> values;
  const RangeValue(this.values);

  int get rowCount => values.length;
  int get columnCount => values.isEmpty ? 0 : values.first.length;

  /// Flatten to a single list of values.
  Iterable<FormulaValue> get flat => values.expand((row) => row);

  /// Get numeric values only.
  Iterable<num> get numbers =>
      flat.whereType<NumberValue>().map((v) => v.value);

  @override
  num? toNumber() => values.length == 1 && values.first.length == 1
      ? values.first.first.toNumber()
      : null;

  @override
  String toText() => values
      .map((row) => row.map((v) => v.toText()).join(','))
      .join(';');

  @override
  bool toBool() => toNumber() != 0;

  @override
  bool get isTruthy => values.isNotEmpty;

  @override
  String toString() => 'RangeValue(${rowCount}x$columnCount)';
}

class EmptyValue extends FormulaValue {
  const EmptyValue();

  @override
  num? toNumber() => 0;

  @override
  String toText() => '';

  @override
  bool toBool() => false;

  @override
  bool get isTruthy => false;

  @override
  String toString() => 'EmptyValue()';
}
