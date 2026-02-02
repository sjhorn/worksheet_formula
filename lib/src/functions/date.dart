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
    TimeValueFunction(),
    WeekNumFunction(),
    IsoWeekNumFunction(),
    NetworkDaysFunction(),
    NetworkDaysIntlFunction(),
    WorkdayFunction(),
    WorkdayIntlFunction(),
    Days360Function(),
    YearFracFunction(),
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

// --- Shared helpers for weekend/holiday functions ---

/// Parse a weekend code into a set of Dart weekday numbers (1=Mon..7=Sun).
Set<int> _parseWeekendCode(int code) {
  return switch (code) {
    1 => {6, 7},    // Sat, Sun
    2 => {7, 1},    // Sun, Mon
    3 => {1, 2},    // Mon, Tue
    4 => {2, 3},    // Tue, Wed
    5 => {3, 4},    // Wed, Thu
    6 => {4, 5},    // Thu, Fri
    7 => {5, 6},    // Fri, Sat
    11 => {7},      // Sun only
    12 => {1},      // Mon only
    13 => {2},      // Tue only
    14 => {3},      // Wed only
    15 => {4},      // Thu only
    16 => {5},      // Fri only
    17 => {6},      // Sat only
    _ => {6, 7},    // default Sat, Sun
  };
}

/// Parse a 7-char weekend string (Mon-Sun, 1=weekend) into weekday set.
Set<int>? _parseWeekendString(String s) {
  if (s.length != 7) return null;
  final result = <int>{};
  for (var i = 0; i < 7; i++) {
    if (s[i] == '1') {
      result.add(i + 1); // 1=Mon..7=Sun
    } else if (s[i] != '0') {
      return null;
    }
  }
  return result;
}

/// Collect holiday serial numbers from a range argument.
Set<int> _collectHolidays(
    List<FormulaNode> args, int holidayArgIndex, EvaluationContext context) {
  if (args.length <= holidayArgIndex) return {};
  final value = args[holidayArgIndex].evaluate(context);
  final holidays = <int>{};
  if (value is RangeValue) {
    for (final cell in value.flat) {
      final n = cell.toNumber();
      if (n != null) holidays.add(n.toInt());
    }
  } else {
    final n = value.toNumber();
    if (n != null) holidays.add(n.toInt());
  }
  return holidays;
}

bool _isWorkday(DateTime date, Set<int> weekendDays, Set<int> holidays) {
  return !weekendDays.contains(date.weekday) &&
      !holidays.contains(_dateToSerial(date).toInt());
}

/// TIMEVALUE(time_text) - Converts a time text string to a fractional day number.
class TimeValueFunction extends FormulaFunction {
  @override
  String get name => 'TIMEVALUE';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final text = value.toText().trim();
    final match =
        RegExp(r'^(\d{1,2}):(\d{2})(?::(\d{2}))?$').firstMatch(text);
    if (match == null) return const FormulaValue.error(FormulaError.value);
    final hour = int.parse(match.group(1)!);
    final minute = int.parse(match.group(2)!);
    final second = match.group(3) != null ? int.parse(match.group(3)!) : 0;
    if (hour > 23 || minute > 59 || second > 59) {
      return const FormulaValue.error(FormulaError.value);
    }
    return FormulaValue.number(
        (hour * 3600 + minute * 60 + second) / 86400);
  }
}

/// WEEKNUM(serial_number, [return_type]) - Returns the week number of a date.
class WeekNumFunction extends FormulaFunction {
  @override
  String get name => 'WEEKNUM';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final serial = values[0].toNumber();
    final returnType =
        args.length > 1 ? values[1].toNumber()?.toInt() ?? 1 : 1;

    if (serial == null) return const FormulaValue.error(FormulaError.value);

    final date = _serialToDate(serial);
    final jan1 = DateTime.utc(date.year, 1, 1);
    final dayOfYear = date.difference(jan1).inDays;

    int dayOffset;
    switch (returnType) {
      case 1: // Sunday start
        dayOffset = jan1.weekday % 7; // Sun=0, Mon=1, ..., Sat=6
      case 2: // Monday start
        dayOffset = jan1.weekday - 1; // Mon=0, Tue=1, ..., Sun=6
      default:
        return const FormulaValue.error(FormulaError.num);
    }

    return FormulaValue.number((dayOfYear + dayOffset) ~/ 7 + 1);
  }
}

/// ISOWEEKNUM(date) - Returns the ISO 8601 week number of a date.
class IsoWeekNumFunction extends FormulaFunction {
  @override
  String get name => 'ISOWEEKNUM';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final value = args[0].evaluate(context);
    final serial = value.toNumber();
    if (serial == null) return const FormulaValue.error(FormulaError.value);

    final date = _serialToDate(serial);
    // Find Thursday of the current week
    final thursday =
        date.add(Duration(days: DateTime.thursday - date.weekday));
    // Find Jan 1 of the Thursday's year
    final jan1 = DateTime.utc(thursday.year, 1, 1);
    // Find Thursday of the week containing Jan 1
    final jan1Thursday =
        jan1.add(Duration(days: DateTime.thursday - jan1.weekday));
    final weekNum =
        (thursday.difference(jan1Thursday).inDays / 7).round() + 1;
    return FormulaValue.number(weekNum);
  }
}

/// NETWORKDAYS(start_date, end_date, [holidays]) - Returns the number of working days.
class NetworkDaysFunction extends FormulaFunction {
  @override
  String get name => 'NETWORKDAYS';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final startSerial = values[0].toNumber();
    final endSerial = values[1].toNumber();

    if (startSerial == null || endSerial == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    final holidays = _collectHolidays(args, 2, context);
    const weekendDays = {6, 7}; // Sat, Sun

    var start = _serialToDate(startSerial);
    var end = _serialToDate(endSerial);
    final sign = start.isAfter(end) ? -1 : 1;
    if (sign == -1) {
      final tmp = start;
      start = end;
      end = tmp;
    }

    var count = 0;
    var current = start;
    while (!current.isAfter(end)) {
      if (_isWorkday(current, weekendDays, holidays)) count++;
      current = current.add(const Duration(days: 1));
    }
    return FormulaValue.number(count * sign);
  }
}

/// NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])
class NetworkDaysIntlFunction extends FormulaFunction {
  @override
  String get name => 'NETWORKDAYS.INTL';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final startSerial = values[0].toNumber();
    final endSerial = values[1].toNumber();

    if (startSerial == null || endSerial == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    Set<int> weekendDays;
    if (args.length > 2) {
      final weekendArg = values[2];
      if (weekendArg is TextValue) {
        final parsed = _parseWeekendString(weekendArg.value);
        if (parsed == null) {
          return const FormulaValue.error(FormulaError.value);
        }
        weekendDays = parsed;
      } else {
        final code = weekendArg.toNumber()?.toInt() ?? 1;
        weekendDays = _parseWeekendCode(code);
      }
    } else {
      weekendDays = {6, 7};
    }

    final holidays = _collectHolidays(args, 3, context);

    var start = _serialToDate(startSerial);
    var end = _serialToDate(endSerial);
    final sign = start.isAfter(end) ? -1 : 1;
    if (sign == -1) {
      final tmp = start;
      start = end;
      end = tmp;
    }

    var count = 0;
    var current = start;
    while (!current.isAfter(end)) {
      if (_isWorkday(current, weekendDays, holidays)) count++;
      current = current.add(const Duration(days: 1));
    }
    return FormulaValue.number(count * sign);
  }
}

/// WORKDAY(start_date, days, [holidays]) - Returns the date after N working days.
class WorkdayFunction extends FormulaFunction {
  @override
  String get name => 'WORKDAY';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final startSerial = values[0].toNumber();
    final days = values[1].toNumber()?.toInt();

    if (startSerial == null || days == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    final holidays = _collectHolidays(args, 2, context);
    const weekendDays = {6, 7};

    var current = _serialToDate(startSerial);
    final step = days >= 0 ? 1 : -1;
    var remaining = days.abs();

    while (remaining > 0) {
      current = current.add(Duration(days: step));
      if (_isWorkday(current, weekendDays, holidays)) remaining--;
    }

    return FormulaValue.number(_dateToSerial(current));
  }
}

/// WORKDAY.INTL(start_date, days, [weekend], [holidays])
class WorkdayIntlFunction extends FormulaFunction {
  @override
  String get name => 'WORKDAY.INTL';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final startSerial = values[0].toNumber();
    final days = values[1].toNumber()?.toInt();

    if (startSerial == null || days == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    Set<int> weekendDays;
    if (args.length > 2) {
      final weekendArg = values[2];
      if (weekendArg is TextValue) {
        final parsed = _parseWeekendString(weekendArg.value);
        if (parsed == null) {
          return const FormulaValue.error(FormulaError.value);
        }
        weekendDays = parsed;
      } else {
        final code = weekendArg.toNumber()?.toInt() ?? 1;
        weekendDays = _parseWeekendCode(code);
      }
    } else {
      weekendDays = {6, 7};
    }

    final holidays = _collectHolidays(args, 3, context);

    var current = _serialToDate(startSerial);
    final step = days >= 0 ? 1 : -1;
    var remaining = days.abs();

    while (remaining > 0) {
      current = current.add(Duration(days: step));
      if (_isWorkday(current, weekendDays, holidays)) remaining--;
    }

    return FormulaValue.number(_dateToSerial(current));
  }
}

/// DAYS360(start_date, end_date, [method]) - Days between dates on 30/360 basis.
class Days360Function extends FormulaFunction {
  @override
  String get name => 'DAYS360';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final startSerial = values[0].toNumber();
    final endSerial = values[1].toNumber();
    final european = args.length > 2 ? values[2].isTruthy : false;

    if (startSerial == null || endSerial == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    final startDate = _serialToDate(startSerial);
    final endDate = _serialToDate(endSerial);

    var d1 = startDate.day;
    var d2 = endDate.day;
    final m1 = startDate.month;
    final m2 = endDate.month;
    final y1 = startDate.year;
    final y2 = endDate.year;

    if (european) {
      if (d1 == 31) d1 = 30;
      if (d2 == 31) d2 = 30;
    } else {
      // US method
      if (d1 == 31) d1 = 30;
      if (d2 == 31 && d1 >= 30) d2 = 30;
    }

    return FormulaValue.number(
        (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1));
  }
}

/// YEARFRAC(start_date, end_date, [basis]) - Fraction of a year between two dates.
class YearFracFunction extends FormulaFunction {
  @override
  String get name => 'YEARFRAC';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final startSerial = values[0].toNumber();
    final endSerial = values[1].toNumber();
    final basis = args.length > 2 ? values[2].toNumber()?.toInt() ?? 0 : 0;

    if (startSerial == null || endSerial == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (basis < 0 || basis > 4) {
      return const FormulaValue.error(FormulaError.num);
    }

    var s = startSerial.toInt();
    var e = endSerial.toInt();
    if (s > e) {
      final tmp = s;
      s = e;
      e = tmp;
    }

    final startDate = _serialToDate(s);
    final endDate = _serialToDate(e);
    final actualDays = e - s;

    return switch (basis) {
      0 => _yearFracUS30360(startDate, endDate),
      1 => _yearFracActualActual(startDate, endDate, actualDays),
      2 => FormulaValue.number(actualDays / 360),
      3 => FormulaValue.number(actualDays / 365),
      4 => _yearFracEuro30360(startDate, endDate),
      _ => const FormulaValue.error(FormulaError.num),
    };
  }

  FormulaValue _yearFracUS30360(DateTime start, DateTime end) {
    var d1 = start.day;
    var d2 = end.day;
    final m1 = start.month;
    final m2 = end.month;
    final y1 = start.year;
    final y2 = end.year;

    if (d1 == 31) d1 = 30;
    if (d2 == 31 && d1 >= 30) d2 = 30;

    final days = (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
    return FormulaValue.number(days / 360);
  }

  FormulaValue _yearFracActualActual(
      DateTime start, DateTime end, int actualDays) {
    if (start.year == end.year) {
      final daysInYear = _isLeapYear(start.year) ? 366 : 365;
      return FormulaValue.number(actualDays / daysInYear);
    }
    // Average the days in the years spanned
    var totalDaysInYears = 0;
    for (var y = start.year; y <= end.year; y++) {
      totalDaysInYears += _isLeapYear(y) ? 366 : 365;
    }
    final avgDaysPerYear = totalDaysInYears / (end.year - start.year + 1);
    return FormulaValue.number(actualDays / avgDaysPerYear);
  }

  FormulaValue _yearFracEuro30360(DateTime start, DateTime end) {
    var d1 = start.day;
    var d2 = end.day;
    if (d1 == 31) d1 = 30;
    if (d2 == 31) d2 = 30;

    final days = (end.year - start.year) * 360 +
        (end.month - start.month) * 30 +
        (d2 - d1);
    return FormulaValue.number(days / 360);
  }

  bool _isLeapYear(int year) {
    return (year % 4 == 0 && year % 100 != 0) || year % 400 == 0;
  }
}
