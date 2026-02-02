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
    DaysFunction(),
    DatedifFunction(),
    DateValueFunction(),
    WeekdayFunction(),
    HourFunction(),
    MinuteFunction(),
    SecondFunction(),
    TimeFunction(),
    EdateFunction(),
    EomonthFunction(),
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

/// DAYS(end_date, start_date) - Days between two dates.
class DaysFunction extends FormulaFunction {
  @override
  String get name => 'DAYS';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final endDate = values[0].toNumber();
    final startDate = values[1].toNumber();

    if (endDate == null || startDate == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    return FormulaValue.number(endDate.toInt() - startDate.toInt());
  }
}

/// DATEDIF(start_date, end_date, unit) - Difference in various units.
class DatedifFunction extends FormulaFunction {
  @override
  String get name => 'DATEDIF';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final startSerial = values[0].toNumber();
    final endSerial = values[1].toNumber();
    final unit = values[2].toText().toUpperCase();

    if (startSerial == null || endSerial == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (startSerial > endSerial) {
      return const FormulaValue.error(FormulaError.num);
    }

    final start = _serialToDate(startSerial);
    final end = _serialToDate(endSerial);

    return switch (unit) {
      'Y' => FormulaValue.number(_fullYears(start, end)),
      'M' => FormulaValue.number(_fullMonths(start, end)),
      'D' => FormulaValue.number(endSerial.toInt() - startSerial.toInt()),
      'YM' => FormulaValue.number(_fullMonths(start, end) % 12),
      'YD' => FormulaValue.number(
          DateTime.utc(start.year, end.month, end.day)
                  .difference(start)
                  .inDays %
              (DateTime.utc(start.year + 1, start.month, start.day)
                      .difference(start)
                      .inDays)),
      'MD' => _mdDiff(start, end),
      _ => const FormulaValue.error(FormulaError.num),
    };
  }

  int _fullYears(DateTime start, DateTime end) {
    var years = end.year - start.year;
    if (end.month < start.month ||
        (end.month == start.month && end.day < start.day)) {
      years--;
    }
    return years;
  }

  int _fullMonths(DateTime start, DateTime end) {
    var months = (end.year - start.year) * 12 + end.month - start.month;
    if (end.day < start.day) months--;
    return months;
  }

  FormulaValue _mdDiff(DateTime start, DateTime end) {
    var days = end.day - start.day;
    if (days < 0) {
      // Get the last day of the month before end
      final prevMonth = DateTime.utc(end.year, end.month, 0);
      days = prevMonth.day - start.day + end.day;
    }
    return FormulaValue.number(days);
  }
}

/// DATEVALUE(date_text) - Converts date text to serial number.
class DateValueFunction extends FormulaFunction {
  @override
  String get name => 'DATEVALUE';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    if (value is NumberValue) return value;
    final text = value.toText().trim();

    // Try ISO 8601 format (YYYY-MM-DD)
    var date = DateTime.tryParse(text);
    if (date != null) {
      return FormulaValue.number(_dateToSerial(date));
    }

    // Try M/D/YYYY format
    final slashMatch = RegExp(r'^(\d{1,2})/(\d{1,2})/(\d{4})$').firstMatch(text);
    if (slashMatch != null) {
      final month = int.parse(slashMatch.group(1)!);
      final day = int.parse(slashMatch.group(2)!);
      final year = int.parse(slashMatch.group(3)!);
      date = DateTime.utc(year, month, day);
      return FormulaValue.number(_dateToSerial(date));
    }

    return const FormulaValue.error(FormulaError.value);
  }
}

/// WEEKDAY(serial_number, [return_type]) - Day of the week.
class WeekdayFunction extends FormulaFunction {
  @override
  String get name => 'WEEKDAY';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final serial = values[0].toNumber();
    final returnType = args.length > 1 ? values[1].toNumber()?.toInt() ?? 1 : 1;

    if (serial == null) return const FormulaValue.error(FormulaError.value);

    final date = _serialToDate(serial);
    // Dart: Monday=1 .. Sunday=7
    final dartDay = date.weekday;

    return switch (returnType) {
      1 => FormulaValue.number(dartDay == 7 ? 1 : dartDay + 1), // Sunday=1..Saturday=7
      2 => FormulaValue.number(dartDay), // Monday=1..Sunday=7
      3 => FormulaValue.number(dartDay - 1), // Monday=0..Sunday=6
      _ => const FormulaValue.error(FormulaError.num),
    };
  }
}

/// HOUR(serial_number) - Extract hour from time.
class HourFunction extends FormulaFunction {
  @override
  String get name => 'HOUR';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber()?.toDouble();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    final fraction = n - n.truncateToDouble();
    final totalSeconds = (fraction * 86400).round();
    return FormulaValue.number(totalSeconds ~/ 3600);
  }
}

/// MINUTE(serial_number) - Extract minute from time.
class MinuteFunction extends FormulaFunction {
  @override
  String get name => 'MINUTE';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber()?.toDouble();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    final fraction = n - n.truncateToDouble();
    final totalSeconds = (fraction * 86400).round();
    return FormulaValue.number((totalSeconds % 3600) ~/ 60);
  }
}

/// SECOND(serial_number) - Extract second from time.
class SecondFunction extends FormulaFunction {
  @override
  String get name => 'SECOND';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final n = value.toNumber()?.toDouble();
    if (n == null) return const FormulaValue.error(FormulaError.value);
    final fraction = n - n.truncateToDouble();
    final totalSeconds = (fraction * 86400).round();
    return FormulaValue.number(totalSeconds % 60);
  }
}

/// TIME(hour, minute, second) - Construct time as fraction of day.
class TimeFunction extends FormulaFunction {
  @override
  String get name => 'TIME';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final hour = values[0].toNumber()?.toInt();
    final minute = values[1].toNumber()?.toInt();
    final second = values[2].toNumber()?.toInt();

    if (hour == null || minute == null || second == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    return FormulaValue.number(
        (hour * 3600 + minute * 60 + second) / 86400);
  }
}

/// EDATE(start_date, months) - Add months to a date.
class EdateFunction extends FormulaFunction {
  @override
  String get name => 'EDATE';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final serial = values[0].toNumber();
    final months = values[1].toNumber()?.toInt();

    if (serial == null || months == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    final date = _serialToDate(serial);
    final result = DateTime.utc(date.year, date.month + months, date.day);
    return FormulaValue.number(_dateToSerial(result));
  }
}

/// EOMONTH(start_date, months) - End of month after adding months.
class EomonthFunction extends FormulaFunction {
  @override
  String get name => 'EOMONTH';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final serial = values[0].toNumber();
    final months = values[1].toNumber()?.toInt();

    if (serial == null || months == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    final date = _serialToDate(serial);
    // Day 0 of the next month = last day of target month
    final result = DateTime.utc(date.year, date.month + months + 1, 0);
    return FormulaValue.number(_dateToSerial(result));
  }
}
