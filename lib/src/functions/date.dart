import '../ast/nodes.dart';
import '../evaluation/context.dart';
import '../evaluation/errors.dart';
import '../evaluation/value.dart';
import 'function.dart';
import 'registry.dart';

/// Register all date functions.
void registerDateFunctions(FunctionRegistry registry) {
  registry.registerAll([
    DateFunction(),
    TodayFunction(),
    NowFunction(),
    YearFunction(),
    MonthFunction(),
    DayFunction(),
  ]);
}

/// Excel epoch: December 30, 1899 (UTC to avoid DST issues).
final _excelEpoch = DateTime.utc(1899, 12, 30);

/// Convert a DateTime to an Excel serial number.
num _dateToSerial(DateTime date) {
  final utc = DateTime.utc(date.year, date.month, date.day);
  return utc.difference(_excelEpoch).inDays;
}

/// Convert an Excel serial number to a DateTime.
DateTime _serialToDate(num serial) {
  return _excelEpoch.add(Duration(days: serial.toInt()));
}

/// DATE(year, month, day) - Creates a date serial number.
class DateFunction extends FormulaFunction {
  @override
  String get name => 'DATE';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final year = values[0].toNumber()?.toInt();
    final month = values[1].toNumber()?.toInt();
    final day = values[2].toNumber()?.toInt();

    if (year == null || month == null || day == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    final date = DateTime.utc(year, month, day);
    return FormulaValue.number(date.difference(_excelEpoch).inDays);
  }
}

/// TODAY() - Returns the current date as a serial number.
class TodayFunction extends FormulaFunction {
  @override
  String get name => 'TODAY';
  @override
  int get minArgs => 0;
  @override
  int get maxArgs => 0;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final now = DateTime.now();
    final today = DateTime(now.year, now.month, now.day);
    return FormulaValue.number(_dateToSerial(today));
  }
}

/// NOW() - Returns the current date and time as a serial number.
class NowFunction extends FormulaFunction {
  @override
  String get name => 'NOW';
  @override
  int get minArgs => 0;
  @override
  int get maxArgs => 0;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final now = DateTime.now();
    final datePart = _dateToSerial(DateTime(now.year, now.month, now.day));
    final timeFraction =
        (now.hour * 3600 + now.minute * 60 + now.second) / 86400;
    return FormulaValue.number(datePart + timeFraction);
  }
}

/// YEAR(serial_number) - Extracts the year from a date serial number.
class YearFunction extends FormulaFunction {
  @override
  String get name => 'YEAR';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(_serialToDate(n).year);
  }
}

/// MONTH(serial_number) - Extracts the month from a date serial number.
class MonthFunction extends FormulaFunction {
  @override
  String get name => 'MONTH';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(_serialToDate(n).month);
  }
}

/// DAY(serial_number) - Extracts the day from a date serial number.
class DayFunction extends FormulaFunction {
  @override
  String get name => 'DAY';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(_serialToDate(n).day);
  }
}
