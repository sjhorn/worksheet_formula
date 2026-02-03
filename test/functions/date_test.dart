import 'package:a1/a1.dart';
import 'package:test/test.dart';
import 'package:worksheet_formula/src/ast/nodes.dart';
import 'package:worksheet_formula/src/evaluation/context.dart';
import 'package:worksheet_formula/src/evaluation/errors.dart';
import 'package:worksheet_formula/src/evaluation/value.dart';
import 'package:worksheet_formula/src/functions/function.dart';
import 'package:worksheet_formula/src/functions/registry.dart';
import 'package:worksheet_formula/src/functions/date.dart';

class _TestContext implements EvaluationContext {
  @override
  A1 get currentCell => 'A1'.a1;
  @override
  String? get currentSheet => null;
  @override
  FormulaValue? getVariable(String name) => null;
  @override
  bool get isCancelled => false;

  final FunctionRegistry _registry;
  _TestContext(this._registry);

  @override
  FormulaValue getCellValue(A1 cell) => const EmptyValue();
  @override
  FormulaValue getRangeValues(A1Range range) =>
      const FormulaValue.error(FormulaError.ref);
  @override
  FormulaFunction? getFunction(String name) => _registry.get(name);
}

void main() {
  late FunctionRegistry registry;
  late _TestContext context;

  setUp(() {
    registry = FunctionRegistry(registerBuiltIns: false);
    registerDateFunctions(registry);
    context = _TestContext(registry);
  });

  FormulaValue eval(FormulaFunction fn, List<FormulaNode> args) =>
      fn.call(args, context);

  group('DATE', () {
    test('creates serial number for known date', () {
      // January 1, 2024 = serial 45292 in Excel
      final result = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      expect(result, const NumberValue(45292));
    });

    test('January 1, 1900 returns 2', () {
      // Excel treats Jan 1, 1900 as serial 1, but due to the Lotus 1-2-3 bug
      // (which treats 1900 as a leap year), the actual difference from epoch
      // (Dec 30, 1899) is 2 days.
      final result = eval(registry.get('DATE')!, [
        const NumberNode(1900),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      expect(result, const NumberValue(2));
    });

    test('month overflow rolls to next year', () {
      // DATE(2024, 13, 1) = DATE(2025, 1, 1)
      final result = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(13),
        const NumberNode(1),
      ]);
      final expected = eval(registry.get('DATE')!, [
        const NumberNode(2025),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      expect(result, expected);
    });

    test('non-numeric args return #VALUE!', () {
      final result = eval(registry.get('DATE')!, [
        const TextNode('abc'),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('TODAY', () {
    test('returns a number', () {
      final result = eval(registry.get('TODAY')!, []);
      expect(result, isA<NumberValue>());
    });

    test('returns reasonable serial number', () {
      // We are past 2023, so serial should be > 45000
      final result = eval(registry.get('TODAY')!, []);
      expect((result as NumberValue).value, greaterThan(45000));
    });

    test('returns integer (no fractional part)', () {
      final result = eval(registry.get('TODAY')!, []);
      final value = (result as NumberValue).value;
      expect(value, equals(value.toInt()));
    });
  });

  group('NOW', () {
    test('returns a number', () {
      final result = eval(registry.get('NOW')!, []);
      expect(result, isA<NumberValue>());
    });

    test('is >= TODAY', () {
      final today = eval(registry.get('TODAY')!, []);
      final now = eval(registry.get('NOW')!, []);
      expect(
        (now as NumberValue).value,
        greaterThanOrEqualTo((today as NumberValue).value),
      );
    });
  });

  group('YEAR', () {
    test('extracts year from date serial', () {
      // DATE(2024, 3, 15) -> YEAR -> 2024
      final dateSerial = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(3),
        const NumberNode(15),
      ]);
      final result = eval(
        registry.get('YEAR')!,
        [NumberNode((dateSerial as NumberValue).value)],
      );
      expect(result, const NumberValue(2024));
    });

    test('non-numeric returns #VALUE!', () {
      final result =
          eval(registry.get('YEAR')!, [const TextNode('abc')]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('MONTH', () {
    test('extracts month from date serial', () {
      final dateSerial = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(3),
        const NumberNode(15),
      ]);
      final result = eval(
        registry.get('MONTH')!,
        [NumberNode((dateSerial as NumberValue).value)],
      );
      expect(result, const NumberValue(3));
    });
  });

  group('DAY', () {
    test('extracts day from date serial', () {
      final dateSerial = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(3),
        const NumberNode(15),
      ]);
      final result = eval(
        registry.get('DAY')!,
        [NumberNode((dateSerial as NumberValue).value)],
      );
      expect(result, const NumberValue(15));
    });
  });

  group('round-trip', () {
    test('DATE(YEAR(d), MONTH(d), DAY(d)) equals d', () {
      final original = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(7),
        const NumberNode(4),
      ]);
      final serial = (original as NumberValue).value;

      final year = eval(registry.get('YEAR')!, [NumberNode(serial)]);
      final month = eval(registry.get('MONTH')!, [NumberNode(serial)]);
      final day = eval(registry.get('DAY')!, [NumberNode(serial)]);

      final roundTrip = eval(registry.get('DATE')!, [
        NumberNode((year as NumberValue).value),
        NumberNode((month as NumberValue).value),
        NumberNode((day as NumberValue).value),
      ]);
      expect(roundTrip, original);
    });
  });

  group('DAYS', () {
    test('returns difference between dates', () {
      final jan1 = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      final jan31 = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(31),
      ]);
      final result = eval(registry.get('DAYS')!, [
        NumberNode((jan31 as NumberValue).value),
        NumberNode((jan1 as NumberValue).value),
      ]);
      expect(result, const NumberValue(30));
    });

    test('negative when end < start', () {
      final jan1 = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      final jan31 = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(31),
      ]);
      final result = eval(registry.get('DAYS')!, [
        NumberNode((jan1 as NumberValue).value),
        NumberNode((jan31 as NumberValue).value),
      ]);
      expect(result, const NumberValue(-30));
    });
  });

  group('DATEDIF', () {
    test('unit Y returns full years', () {
      final start = eval(registry.get('DATE')!, [
        const NumberNode(2020),
        const NumberNode(3),
        const NumberNode(15),
      ]);
      final end = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(7),
        const NumberNode(4),
      ]);
      final result = eval(registry.get('DATEDIF')!, [
        NumberNode((start as NumberValue).value),
        NumberNode((end as NumberValue).value),
        const TextNode('Y'),
      ]);
      expect(result, const NumberValue(4));
    });

    test('unit M returns full months', () {
      final start = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(15),
      ]);
      final end = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(4),
        const NumberNode(20),
      ]);
      final result = eval(registry.get('DATEDIF')!, [
        NumberNode((start as NumberValue).value),
        NumberNode((end as NumberValue).value),
        const TextNode('M'),
      ]);
      expect(result, const NumberValue(3));
    });

    test('unit D returns days', () {
      final start = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      final end = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(31),
      ]);
      final result = eval(registry.get('DATEDIF')!, [
        NumberNode((start as NumberValue).value),
        NumberNode((end as NumberValue).value),
        const TextNode('D'),
      ]);
      expect(result, const NumberValue(30));
    });

    test('start > end returns #NUM!', () {
      final result = eval(registry.get('DATEDIF')!, [
        const NumberNode(45300),
        const NumberNode(45200),
        const TextNode('D'),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });

  group('DATEVALUE', () {
    test('parses ISO format', () {
      final result = eval(
          registry.get('DATEVALUE')!, [const TextNode('2024-01-01')]);
      expect(result, const NumberValue(45292));
    });

    test('parses M/D/YYYY format', () {
      final result = eval(
          registry.get('DATEVALUE')!, [const TextNode('1/1/2024')]);
      expect(result, const NumberValue(45292));
    });

    test('invalid text returns #VALUE!', () {
      final result = eval(
          registry.get('DATEVALUE')!, [const TextNode('not a date')]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('WEEKDAY', () {
    test('returns day of week (default type 1: Sun=1..Sat=7)', () {
      // 2024-01-01 is a Monday
      final serial = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      final result = eval(registry.get('WEEKDAY')!, [
        NumberNode((serial as NumberValue).value),
      ]);
      expect(result, const NumberValue(2)); // Monday = 2 in type 1
    });

    test('type 2: Mon=1..Sun=7', () {
      final serial = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      final result = eval(registry.get('WEEKDAY')!, [
        NumberNode((serial as NumberValue).value),
        const NumberNode(2),
      ]);
      expect(result, const NumberValue(1)); // Monday = 1 in type 2
    });
  });

  group('HOUR', () {
    test('extracts hour from time', () {
      // 0.5 = noon (12:00:00)
      final result = eval(registry.get('HOUR')!, [const NumberNode(0.5)]);
      expect(result, const NumberValue(12));
    });

    test('extracts hour from full datetime', () {
      // 45292.75 = 6pm on Jan 1 2024
      final result =
          eval(registry.get('HOUR')!, [const NumberNode(45292.75)]);
      expect(result, const NumberValue(18));
    });
  });

  group('MINUTE', () {
    test('extracts minute from time', () {
      // 0.5 + 30min/1440 = 12:30
      final result = eval(
          registry.get('MINUTE')!, [NumberNode(0.5 + 30 / 1440)]);
      expect(result, const NumberValue(30));
    });
  });

  group('SECOND', () {
    test('extracts second from time', () {
      // 45 seconds = 45/86400
      final result = eval(
          registry.get('SECOND')!, [NumberNode(45 / 86400)]);
      expect(result, const NumberValue(45));
    });
  });

  group('TIME', () {
    test('constructs time fraction', () {
      final result = eval(registry.get('TIME')!, [
        const NumberNode(12),
        const NumberNode(0),
        const NumberNode(0),
      ]);
      expect(result, const NumberValue(0.5)); // Noon = 0.5
    });

    test('6am = 0.25', () {
      final result = eval(registry.get('TIME')!, [
        const NumberNode(6),
        const NumberNode(0),
        const NumberNode(0),
      ]);
      expect(result, const NumberValue(0.25));
    });
  });

  group('EDATE', () {
    test('adds months to date', () {
      final jan = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(15),
      ]);
      final result = eval(registry.get('EDATE')!, [
        NumberNode((jan as NumberValue).value),
        const NumberNode(3),
      ]);
      final expected = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(4),
        const NumberNode(15),
      ]);
      expect(result, expected);
    });

    test('negative months subtracts', () {
      final mar = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(3),
        const NumberNode(15),
      ]);
      final result = eval(registry.get('EDATE')!, [
        NumberNode((mar as NumberValue).value),
        const NumberNode(-2),
      ]);
      final expected = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(15),
      ]);
      expect(result, expected);
    });
  });

  group('EOMONTH', () {
    test('returns end of month', () {
      final jan15 = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(15),
      ]);
      final result = eval(registry.get('EOMONTH')!, [
        NumberNode((jan15 as NumberValue).value),
        const NumberNode(0),
      ]);
      final expected = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(31),
      ]);
      expect(result, expected);
    });

    test('returns end of next month', () {
      final jan15 = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(15),
      ]);
      final result = eval(registry.get('EOMONTH')!, [
        NumberNode((jan15 as NumberValue).value),
        const NumberNode(1),
      ]);
      // Feb 2024 has 29 days (leap year)
      final expected = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(2),
        const NumberNode(29),
      ]);
      expect(result, expected);
    });

    test('negative months goes backward', () {
      final mar15 = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(3),
        const NumberNode(15),
      ]);
      final result = eval(registry.get('EOMONTH')!, [
        NumberNode((mar15 as NumberValue).value),
        const NumberNode(-1),
      ]);
      final expected = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(2),
        const NumberNode(29),
      ]);
      expect(result, expected);
    });
  });

  group('TIMEVALUE', () {
    test('parses HH:MM:SS', () {
      final result =
          eval(registry.get('TIMEVALUE')!, [const TextNode('12:00:00')]);
      expect(result, const NumberValue(0.5));
    });

    test('parses HH:MM', () {
      final result =
          eval(registry.get('TIMEVALUE')!, [const TextNode('6:00')]);
      expect(result, const NumberValue(0.25));
    });

    test('invalid text returns #VALUE!', () {
      final result =
          eval(registry.get('TIMEVALUE')!, [const TextNode('not a time')]);
      expect(result, const ErrorValue(FormulaError.value));
    });
  });

  group('WEEKNUM', () {
    test('type 1 Sunday start', () {
      // Jan 1, 2024 is a Monday → week 1 (type 1: Sun start)
      final serial = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      final result = eval(registry.get('WEEKNUM')!, [
        NumberNode((serial as NumberValue).value),
      ]);
      expect(result, const NumberValue(1));
    });

    test('type 2 Monday start', () {
      // Jan 7, 2024 is a Sunday → still week 1 for Monday start
      final serial = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(7),
      ]);
      final result = eval(registry.get('WEEKNUM')!, [
        NumberNode((serial as NumberValue).value),
        const NumberNode(2),
      ]);
      expect(result, const NumberValue(1));
    });

    test('Jan 8 Monday start gives week 2', () {
      final serial = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(8),
      ]);
      final result = eval(registry.get('WEEKNUM')!, [
        NumberNode((serial as NumberValue).value),
        const NumberNode(2),
      ]);
      expect(result, const NumberValue(2));
    });
  });

  group('ISOWEEKNUM', () {
    test('Jan 1 2024 is ISO week 1', () {
      final serial = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      final result = eval(registry.get('ISOWEEKNUM')!, [
        NumberNode((serial as NumberValue).value),
      ]);
      expect(result, const NumberValue(1));
    });

    test('Dec 31 2024 is ISO week 1 of 2025', () {
      // Dec 31, 2024 is a Tuesday. The Thursday of that week is Jan 2, 2025.
      final serial = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(12),
        const NumberNode(31),
      ]);
      final result = eval(registry.get('ISOWEEKNUM')!, [
        NumberNode((serial as NumberValue).value),
      ]);
      expect(result, const NumberValue(1));
    });
  });

  group('NETWORKDAYS', () {
    test('counts working days', () {
      // Jan 1-5, 2024 (Mon-Fri) = 5 working days
      final start = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      final end = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(5),
      ]);
      final result = eval(registry.get('NETWORKDAYS')!, [
        NumberNode((start as NumberValue).value),
        NumberNode((end as NumberValue).value),
      ]);
      expect(result, const NumberValue(5));
    });

    test('skips weekends', () {
      // Jan 1-7, 2024 (Mon-Sun) = 5 working days
      final start = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      final end = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(7),
      ]);
      final result = eval(registry.get('NETWORKDAYS')!, [
        NumberNode((start as NumberValue).value),
        NumberNode((end as NumberValue).value),
      ]);
      expect(result, const NumberValue(5));
    });
  });

  group('NETWORKDAYS.INTL', () {
    test('custom weekend code 11 (Sunday only)', () {
      // Jan 1-7, 2024 (Mon-Sun) with Sunday-only weekend = 6 working days
      final start = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      final end = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(7),
      ]);
      final result = eval(registry.get('NETWORKDAYS.INTL')!, [
        NumberNode((start as NumberValue).value),
        NumberNode((end as NumberValue).value),
        const NumberNode(11),
      ]);
      expect(result, const NumberValue(6));
    });
  });

  group('WORKDAY', () {
    test('adds working days', () {
      // Start Jan 1, 2024 (Monday), add 5 working days = Jan 8 (Monday)
      final start = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      final result = eval(registry.get('WORKDAY')!, [
        NumberNode((start as NumberValue).value),
        const NumberNode(5),
      ]);
      final expected = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(8),
      ]);
      expect(result, expected);
    });

    test('negative days goes backward', () {
      final start = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(8),
      ]);
      final result = eval(registry.get('WORKDAY')!, [
        NumberNode((start as NumberValue).value),
        const NumberNode(-5),
      ]);
      final expected = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      expect(result, expected);
    });
  });

  group('WORKDAY.INTL', () {
    test('custom weekend code', () {
      final start = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      // With Sunday-only weekend (code 11), 6 days lands on Jan 8
      final result = eval(registry.get('WORKDAY.INTL')!, [
        NumberNode((start as NumberValue).value),
        const NumberNode(6),
        const NumberNode(11),
      ]);
      final expected = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(8),
      ]);
      expect(result, expected);
    });
  });

  group('DAYS360', () {
    test('US method (default)', () {
      final start = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      final end = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(7),
        const NumberNode(1),
      ]);
      final result = eval(registry.get('DAYS360')!, [
        NumberNode((start as NumberValue).value),
        NumberNode((end as NumberValue).value),
      ]);
      expect(result, const NumberValue(180));
    });

    test('European method', () {
      final start = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(31),
      ]);
      final end = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(3),
        const NumberNode(31),
      ]);
      final result = eval(registry.get('DAYS360')!, [
        NumberNode((start as NumberValue).value),
        NumberNode((end as NumberValue).value),
        const BooleanNode(true),
      ]);
      expect(result, const NumberValue(60));
    });
  });

  group('YEARFRAC', () {
    test('basis 0 (US 30/360)', () {
      final start = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      final end = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(7),
        const NumberNode(1),
      ]);
      final result = eval(registry.get('YEARFRAC')!, [
        NumberNode((start as NumberValue).value),
        NumberNode((end as NumberValue).value),
        const NumberNode(0),
      ]);
      expect(result, const NumberValue(0.5));
    });

    test('basis 2 (Actual/360)', () {
      final start = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      final end = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(31),
      ]);
      final result = eval(registry.get('YEARFRAC')!, [
        NumberNode((start as NumberValue).value),
        NumberNode((end as NumberValue).value),
        const NumberNode(2),
      ]);
      expect(
          (result as NumberValue).value, closeTo(30 / 360, 0.0001));
    });

    test('basis 3 (Actual/365)', () {
      final start = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(1),
      ]);
      final end = eval(registry.get('DATE')!, [
        const NumberNode(2024),
        const NumberNode(1),
        const NumberNode(31),
      ]);
      final result = eval(registry.get('YEARFRAC')!, [
        NumberNode((start as NumberValue).value),
        NumberNode((end as NumberValue).value),
        const NumberNode(3),
      ]);
      expect(
          (result as NumberValue).value, closeTo(30 / 365, 0.0001));
    });

    test('invalid basis returns #NUM!', () {
      final result = eval(registry.get('YEARFRAC')!, [
        const NumberNode(45292),
        const NumberNode(45322),
        const NumberNode(5),
      ]);
      expect(result, const ErrorValue(FormulaError.num));
    });
  });
}
