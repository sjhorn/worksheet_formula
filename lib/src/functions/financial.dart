import 'dart:math' as math;

import '../ast/nodes.dart';
import '../evaluation/context.dart';
import '../evaluation/errors.dart';
import '../evaluation/value.dart';
import 'function.dart';
import 'registry.dart';

/// Register all financial functions.
void registerFinancialFunctions(FunctionRegistry registry) {
  registry.registerAll([
    // Wave 1 — Simple Closed-Form
    EffectFunction(),
    NominalFunction(),
    PdurationFunction(),
    RriFunction(),
    IspmtFunction(),
    SlnFunction(),
    SydFunction(),
    DollarDeFunction(),
    DollarFrFunction(),
    FvScheduleFunction(),
    NpvFunction(),
    TbillEqFunction(),
    TbillPriceFunction(),
    TbillYieldFunction(),
    // Wave 2 — TVM Annuity
    PmtFunction(),
    FvFunction(),
    PvFunction(),
    NperFunction(),
    IpmtFunction(),
    PpmtFunction(),
    CumipmtFunction(),
    CumprincFunction(),
    // Wave 3 — Iterative Solvers
    RateFunction(),
    IrrFunction(),
    XnpvFunction(),
    XirrFunction(),
    // Wave 4 — MIRR + Complex Depreciation
    MirrFunction(),
    DbFunction(),
    DdbFunction(),
    VdbFunction(),
    // Wave 5 — Simple Bond/Security Functions
    DiscFunction(),
    IntrateFunction(),
    ReceivedFunction(),
    PriceDiscFunction(),
    PriceMatFunction(),
    AccrintFunction(),
    // Wave 6 — Complex Bond Functions
    PriceFunction(),
    YieldFunction(),
    DurationFunction(),
    MdurationFunction(),
  ]);
}

// ─── Shared Helpers ────────────────────────────────────────────

/// Excel epoch: December 30, 1899 (UTC to avoid DST issues).
final _excelEpoch = DateTime.utc(1899, 12, 30);

/// Convert an Excel serial number to a DateTime.
DateTime _serialToDate(num serial) {
  return _excelEpoch.add(Duration(days: serial.toInt()));
}

/// Check if a year is a leap year.
bool _isLeapYear(int year) {
  return (year % 4 == 0 && year % 100 != 0) || year % 400 == 0;
}

/// Compute year fraction between two dates according to day count basis.
double _yearFrac(DateTime start, DateTime end, int basis) {
  var s = start;
  var e = end;
  if (s.isAfter(e)) {
    final tmp = s;
    s = e;
    e = tmp;
  }
  final actualDays = e.difference(s).inDays;

  switch (basis) {
    case 0: // US 30/360
      var d1 = s.day;
      var d2 = e.day;
      if (d1 == 31) d1 = 30;
      if (d2 == 31 && d1 >= 30) d2 = 30;
      final days = (e.year - s.year) * 360 +
          (e.month - s.month) * 30 +
          (d2 - d1);
      return days / 360;
    case 1: // Actual/Actual
      if (s.year == e.year) {
        final daysInYear = _isLeapYear(s.year) ? 366 : 365;
        return actualDays / daysInYear;
      }
      var totalDaysInYears = 0;
      for (var y = s.year; y <= e.year; y++) {
        totalDaysInYears += _isLeapYear(y) ? 366 : 365;
      }
      final avgDaysPerYear = totalDaysInYears / (e.year - s.year + 1);
      return actualDays / avgDaysPerYear;
    case 2: // Actual/360
      return actualDays / 360;
    case 3: // Actual/365
      return actualDays / 365;
    case 4: // European 30/360
      var d1 = s.day;
      var d2 = e.day;
      if (d1 == 31) d1 = 30;
      if (d2 == 31) d2 = 30;
      final days = (e.year - s.year) * 360 +
          (e.month - s.month) * 30 +
          (d2 - d1);
      return days / 360;
    default:
      return actualDays / 360;
  }
}

// ─── TVM Core Helpers ──────────────────────────────────────────

/// Calculate payment amount.
double _pmt(double rate, double nper, double pv, double fv, int type) {
  if (rate == 0) {
    return -(pv + fv) / nper;
  }
  final pvif = math.pow(1 + rate, nper);
  var pmt = rate * (pv * pvif + fv) / (pvif - 1);
  if (type == 1) {
    pmt /= (1 + rate);
  }
  return -pmt;
}

/// Calculate future value.
double _fv(double rate, double nper, double pmt, double pv, int type) {
  if (rate == 0) {
    return -(pv + pmt * nper);
  }
  final pvif = math.pow(1 + rate, nper);
  var fvResult = -pv * pvif - pmt * (1 + rate * type) * (pvif - 1) / rate;
  return fvResult;
}

/// Calculate interest payment for a specific period.
double _ipmt(
    double rate, double per, double nper, double pv, double fv, int type) {
  final payment = _pmt(rate, nper, pv, fv, type);
  double interest;
  if (per == 1) {
    if (type == 1) {
      interest = 0;
    } else {
      interest = -pv * rate;
    }
  } else {
    if (type == 1) {
      final fvBefore =
          _fv(rate, per - 2, payment, pv, 1);
      interest = -(fvBefore + payment) * rate;
    } else {
      interest = _fv(rate, per - 1, payment, pv, 0) * rate;
    }
  }
  return interest;
}

// ─── Newton-Raphson Solver ─────────────────────────────────────

/// Solve f(x) = 0 using Newton-Raphson method.
double? _newtonRaphson(
  double Function(double) f,
  double Function(double) df,
  double guess, {
  int maxIter = 100,
  double tol = 1e-10,
}) {
  var x = guess;
  for (var i = 0; i < maxIter; i++) {
    final fx = f(x);
    if (fx.abs() < tol) return x;
    final dfx = df(x);
    if (dfx.abs() < 1e-14) return null;
    final xNew = x - fx / dfx;
    if ((xNew - x).abs() < tol) return xNew;
    x = xNew;
  }
  return null;
}

// ─── Coupon Date Helpers ───────────────────────────────────────

/// Number of coupon periods between settlement and maturity.
int _coupNum(DateTime settlement, DateTime maturity, int frequency) {
  var count = 0;
  var date = maturity;
  while (date.isAfter(settlement)) {
    date = DateTime.utc(date.year, date.month - (12 ~/ frequency), date.day);
    count++;
  }
  return count;
}

/// Next coupon date after settlement.
DateTime _coupNcd(DateTime settlement, DateTime maturity, int frequency) {
  var date = maturity;
  final months = 12 ~/ frequency;
  while (date.isAfter(settlement)) {
    date = DateTime.utc(date.year, date.month - months, date.day);
  }
  return DateTime.utc(date.year, date.month + months, date.day);
}

/// Previous coupon date before settlement.
DateTime _coupPcd(DateTime settlement, DateTime maturity, int frequency) {
  var date = maturity;
  final months = 12 ~/ frequency;
  while (date.isAfter(settlement)) {
    date = DateTime.utc(date.year, date.month - months, date.day);
  }
  return date;
}

/// Days in the coupon period containing settlement.
int _coupDays(
    DateTime settlement, DateTime maturity, int frequency, int basis) {
  if (basis == 0 || basis == 4) {
    return 360 ~/ frequency;
  }
  final pcd = _coupPcd(settlement, maturity, frequency);
  final ncd = _coupNcd(settlement, maturity, frequency);
  return ncd.difference(pcd).inDays;
}

/// Days from beginning of coupon period to settlement.
int _coupDaysBs(
    DateTime settlement, DateTime maturity, int frequency, int basis) {
  if (basis == 0 || basis == 4) {
    final pcd = _coupPcd(settlement, maturity, frequency);
    var d1 = pcd.day;
    var d2 = settlement.day;
    if (basis == 0) {
      if (d1 == 31) d1 = 30;
      if (d2 == 31 && d1 >= 30) d2 = 30;
    } else {
      if (d1 == 31) d1 = 30;
      if (d2 == 31) d2 = 30;
    }
    return (settlement.year - pcd.year) * 360 +
        (settlement.month - pcd.month) * 30 +
        (d2 - d1);
  }
  final pcd = _coupPcd(settlement, maturity, frequency);
  return settlement.difference(pcd).inDays;
}

/// Collect numbers from a FormulaValue (handles ranges).
List<double> _collectValuesFromArg(FormulaNode arg, EvaluationContext context) {
  final value = arg.evaluate(context);
  final result = <double>[];
  if (value is RangeValue) {
    for (final cell in value.flat) {
      final n = cell.toNumber();
      if (n != null) result.add(n.toDouble());
    }
  } else {
    final n = value.toNumber();
    if (n != null) result.add(n.toDouble());
  }
  return result;
}

// ═══════════════════════════════════════════════════════════════
// Wave 1 — Simple Closed-Form
// ═══════════════════════════════════════════════════════════════

/// EFFECT(nominal_rate, npery) - Effective annual interest rate.
class EffectFunction extends FormulaFunction {
  @override
  String get name => 'EFFECT';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final rate = values[0].toNumber()?.toDouble();
    final npery = values[1].toNumber()?.toInt();

    if (rate == null || npery == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (rate <= 0) return const FormulaValue.error(FormulaError.num);
    if (npery < 1) return const FormulaValue.error(FormulaError.num);

    return FormulaValue.number(math.pow(1 + rate / npery, npery) - 1);
  }
}

/// NOMINAL(effect_rate, npery) - Nominal annual interest rate.
class NominalFunction extends FormulaFunction {
  @override
  String get name => 'NOMINAL';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final rate = values[0].toNumber()?.toDouble();
    final npery = values[1].toNumber()?.toInt();

    if (rate == null || npery == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (rate <= 0) return const FormulaValue.error(FormulaError.num);
    if (npery < 1) return const FormulaValue.error(FormulaError.num);

    return FormulaValue.number(
        npery * (math.pow(1 + rate, 1.0 / npery) - 1));
  }
}

/// PDURATION(rate, pv, fv) - Periods required for investment to reach a value.
class PdurationFunction extends FormulaFunction {
  @override
  String get name => 'PDURATION';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final rate = values[0].toNumber()?.toDouble();
    final pv = values[1].toNumber()?.toDouble();
    final fv = values[2].toNumber()?.toDouble();

    if (rate == null || pv == null || fv == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (rate <= 0) return const FormulaValue.error(FormulaError.num);
    if (pv <= 0 || fv <= 0) return const FormulaValue.error(FormulaError.num);

    return FormulaValue.number(
        (math.log(fv) - math.log(pv)) / math.log(1 + rate));
  }
}

/// RRI(nper, pv, fv) - Equivalent interest rate for growth of investment.
class RriFunction extends FormulaFunction {
  @override
  String get name => 'RRI';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final nper = values[0].toNumber()?.toDouble();
    final pv = values[1].toNumber()?.toDouble();
    final fv = values[2].toNumber()?.toDouble();

    if (nper == null || pv == null || fv == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (nper <= 0) return const FormulaValue.error(FormulaError.num);
    if (pv == 0) return const FormulaValue.error(FormulaError.num);

    return FormulaValue.number(math.pow(fv / pv, 1.0 / nper) - 1);
  }
}

/// ISPMT(rate, per, nper, pv) - Interest paid for a given period.
class IspmtFunction extends FormulaFunction {
  @override
  String get name => 'ISPMT';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final rate = values[0].toNumber()?.toDouble();
    final per = values[1].toNumber()?.toDouble();
    final nper = values[2].toNumber()?.toDouble();
    final pv = values[3].toNumber()?.toDouble();

    if (rate == null || per == null || nper == null || pv == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    return FormulaValue.number(pv * rate * (per / nper - 1));
  }
}

/// SLN(cost, salvage, life) - Straight-line depreciation.
class SlnFunction extends FormulaFunction {
  @override
  String get name => 'SLN';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final cost = values[0].toNumber()?.toDouble();
    final salvage = values[1].toNumber()?.toDouble();
    final life = values[2].toNumber()?.toDouble();

    if (cost == null || salvage == null || life == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (life == 0) return const FormulaValue.error(FormulaError.divZero);

    return FormulaValue.number((cost - salvage) / life);
  }
}

/// SYD(cost, salvage, life, per) - Sum-of-years-digits depreciation.
class SydFunction extends FormulaFunction {
  @override
  String get name => 'SYD';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final cost = values[0].toNumber()?.toDouble();
    final salvage = values[1].toNumber()?.toDouble();
    final life = values[2].toNumber()?.toDouble();
    final per = values[3].toNumber()?.toDouble();

    if (cost == null || salvage == null || life == null || per == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (life <= 0) return const FormulaValue.error(FormulaError.num);
    if (per < 1 || per > life) {
      return const FormulaValue.error(FormulaError.num);
    }

    return FormulaValue.number(
        (cost - salvage) * (life - per + 1) * 2 / (life * (life + 1)));
  }
}

/// DOLLARDE(fractional_dollar, fraction) - Convert fractional to decimal.
class DollarDeFunction extends FormulaFunction {
  @override
  String get name => 'DOLLARDE';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final fracDollar = values[0].toNumber()?.toDouble();
    final fraction = values[1].toNumber()?.toInt();

    if (fracDollar == null || fraction == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (fraction < 1) return const FormulaValue.error(FormulaError.num);

    final sign = fracDollar < 0 ? -1.0 : 1.0;
    final absDollar = fracDollar.abs();
    final intPart = absDollar.truncateToDouble();
    final fracPart = absDollar - intPart;

    // The fractional part is in units of 1/fraction
    // e.g., 1.02 with fraction 16 means 1 + 02/16
    final digits = fraction.toString().length;
    final divisor = math.pow(10, digits);
    final numerator = (fracPart * divisor).round();
    return FormulaValue.number(sign * (intPart + numerator / fraction));
  }
}

/// DOLLARFR(decimal_dollar, fraction) - Convert decimal to fractional.
class DollarFrFunction extends FormulaFunction {
  @override
  String get name => 'DOLLARFR';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final decDollar = values[0].toNumber()?.toDouble();
    final fraction = values[1].toNumber()?.toInt();

    if (decDollar == null || fraction == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (fraction < 1) return const FormulaValue.error(FormulaError.num);

    final sign = decDollar < 0 ? -1.0 : 1.0;
    final absDollar = decDollar.abs();
    final intPart = absDollar.truncateToDouble();
    final fracPart = absDollar - intPart;

    final digits = fraction.toString().length;
    final divisor = math.pow(10, digits);
    final numerator = fracPart * fraction;
    return FormulaValue.number(sign * (intPart + numerator / divisor));
  }
}

/// FVSCHEDULE(principal, schedule) - Future value with variable rates.
class FvScheduleFunction extends FormulaFunction {
  @override
  String get name => 'FVSCHEDULE';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final principal = values[0].toNumber()?.toDouble();
    if (principal == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    final scheduleValue = values[1];
    final rates = <double>[];
    if (scheduleValue is RangeValue) {
      for (final cell in scheduleValue.flat) {
        final n = cell.toNumber();
        if (n != null) rates.add(n.toDouble());
      }
    } else {
      final n = scheduleValue.toNumber();
      if (n != null) rates.add(n.toDouble());
    }

    var result = principal;
    for (final rate in rates) {
      result *= (1 + rate);
    }
    return FormulaValue.number(result);
  }
}

/// NPV(rate, value1, [value2], ...) - Net present value.
class NpvFunction extends FormulaFunction {
  @override
  String get name => 'NPV';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final rateValue = args[0].evaluate(context);
    final rate = rateValue.toNumber()?.toDouble();
    if (rate == null) return const FormulaValue.error(FormulaError.value);

    final cashFlows = <double>[];
    for (var i = 1; i < args.length; i++) {
      cashFlows.addAll(_collectValuesFromArg(args[i], context));
    }

    var npv = 0.0;
    for (var i = 0; i < cashFlows.length; i++) {
      npv += cashFlows[i] / math.pow(1 + rate, i + 1);
    }
    return FormulaValue.number(npv);
  }
}

/// TBILLEQ(settlement, maturity, discount) - T-Bill bond-equivalent yield.
class TbillEqFunction extends FormulaFunction {
  @override
  String get name => 'TBILLEQ';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final settlement = values[0].toNumber()?.toInt();
    final maturity = values[1].toNumber()?.toInt();
    final disc = values[2].toNumber()?.toDouble();

    if (settlement == null || maturity == null || disc == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    final dsm = maturity - settlement;
    if (dsm > 365) return const FormulaValue.error(FormulaError.num);
    if (dsm <= 0) return const FormulaValue.error(FormulaError.num);

    return FormulaValue.number(365 * disc / (360 - disc * dsm));
  }
}

/// TBILLPRICE(settlement, maturity, discount) - T-Bill price per $100.
class TbillPriceFunction extends FormulaFunction {
  @override
  String get name => 'TBILLPRICE';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final settlement = values[0].toNumber()?.toInt();
    final maturity = values[1].toNumber()?.toInt();
    final disc = values[2].toNumber()?.toDouble();

    if (settlement == null || maturity == null || disc == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    final dsm = maturity - settlement;
    if (dsm > 365) return const FormulaValue.error(FormulaError.num);
    if (dsm <= 0) return const FormulaValue.error(FormulaError.num);

    return FormulaValue.number(100 * (1 - disc * dsm / 360));
  }
}

/// TBILLYIELD(settlement, maturity, pr) - T-Bill yield.
class TbillYieldFunction extends FormulaFunction {
  @override
  String get name => 'TBILLYIELD';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final settlement = values[0].toNumber()?.toInt();
    final maturity = values[1].toNumber()?.toInt();
    final pr = values[2].toNumber()?.toDouble();

    if (settlement == null || maturity == null || pr == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (pr <= 0) return const FormulaValue.error(FormulaError.num);

    final dsm = maturity - settlement;
    if (dsm > 365) return const FormulaValue.error(FormulaError.num);
    if (dsm <= 0) return const FormulaValue.error(FormulaError.num);

    return FormulaValue.number(((100 - pr) / pr) * (360 / dsm));
  }
}

// ═══════════════════════════════════════════════════════════════
// Wave 2 — TVM Annuity
// ═══════════════════════════════════════════════════════════════

/// PMT(rate, nper, pv, [fv], [type]) - Payment for a loan.
class PmtFunction extends FormulaFunction {
  @override
  String get name => 'PMT';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 5;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final rate = values[0].toNumber()?.toDouble();
    final nper = values[1].toNumber()?.toDouble();
    final pv = values[2].toNumber()?.toDouble();
    final fv = args.length > 3 ? values[3].toNumber()?.toDouble() ?? 0 : 0.0;
    final type = args.length > 4 ? values[4].toNumber()?.toInt() ?? 0 : 0;

    if (rate == null || nper == null || pv == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    return FormulaValue.number(_pmt(rate, nper, pv, fv, type));
  }
}

/// FV(rate, nper, pmt, [pv], [type]) - Future value.
class FvFunction extends FormulaFunction {
  @override
  String get name => 'FV';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 5;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final rate = values[0].toNumber()?.toDouble();
    final nper = values[1].toNumber()?.toDouble();
    final pmt = values[2].toNumber()?.toDouble();
    final pv = args.length > 3 ? values[3].toNumber()?.toDouble() ?? 0 : 0.0;
    final type = args.length > 4 ? values[4].toNumber()?.toInt() ?? 0 : 0;

    if (rate == null || nper == null || pmt == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    return FormulaValue.number(_fv(rate, nper, pmt, pv, type));
  }
}

/// PV(rate, nper, pmt, [fv], [type]) - Present value.
class PvFunction extends FormulaFunction {
  @override
  String get name => 'PV';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 5;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final rate = values[0].toNumber()?.toDouble();
    final nper = values[1].toNumber()?.toDouble();
    final pmt = values[2].toNumber()?.toDouble();
    final fv = args.length > 3 ? values[3].toNumber()?.toDouble() ?? 0 : 0.0;
    final type = args.length > 4 ? values[4].toNumber()?.toInt() ?? 0 : 0;

    if (rate == null || nper == null || pmt == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    if (rate == 0) {
      return FormulaValue.number(-(fv + pmt * nper));
    }

    final pvif = math.pow(1 + rate, nper);
    final pvResult =
        (-fv - pmt * (1 + rate * type) * (pvif - 1) / rate) / pvif;
    return FormulaValue.number(pvResult);
  }
}

/// NPER(rate, pmt, pv, [fv], [type]) - Number of periods.
class NperFunction extends FormulaFunction {
  @override
  String get name => 'NPER';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 5;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final rate = values[0].toNumber()?.toDouble();
    final pmt = values[1].toNumber()?.toDouble();
    final pv = values[2].toNumber()?.toDouble();
    final fv = args.length > 3 ? values[3].toNumber()?.toDouble() ?? 0 : 0.0;
    final type = args.length > 4 ? values[4].toNumber()?.toInt() ?? 0 : 0;

    if (rate == null || pmt == null || pv == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    if (rate == 0) {
      if (pmt == 0) return const FormulaValue.error(FormulaError.num);
      return FormulaValue.number(-(pv + fv) / pmt);
    }

    final pmtAdj = pmt * (1 + rate * type);
    final num1 = pmtAdj - fv * rate;
    final num2 = pmtAdj + pv * rate;
    if (num1 / num2 <= 0) return const FormulaValue.error(FormulaError.num);

    return FormulaValue.number(math.log(num1 / num2) / math.log(1 + rate));
  }
}

/// IPMT(rate, per, nper, pv, [fv], [type]) - Interest payment for a period.
class IpmtFunction extends FormulaFunction {
  @override
  String get name => 'IPMT';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 6;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final rate = values[0].toNumber()?.toDouble();
    final per = values[1].toNumber()?.toDouble();
    final nper = values[2].toNumber()?.toDouble();
    final pv = values[3].toNumber()?.toDouble();
    final fv = args.length > 4 ? values[4].toNumber()?.toDouble() ?? 0 : 0.0;
    final type = args.length > 5 ? values[5].toNumber()?.toInt() ?? 0 : 0;

    if (rate == null || per == null || nper == null || pv == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (per < 1 || per > nper) {
      return const FormulaValue.error(FormulaError.num);
    }

    return FormulaValue.number(_ipmt(rate, per, nper, pv, fv, type));
  }
}

/// PPMT(rate, per, nper, pv, [fv], [type]) - Principal payment for a period.
class PpmtFunction extends FormulaFunction {
  @override
  String get name => 'PPMT';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 6;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final rate = values[0].toNumber()?.toDouble();
    final per = values[1].toNumber()?.toDouble();
    final nper = values[2].toNumber()?.toDouble();
    final pv = values[3].toNumber()?.toDouble();
    final fv = args.length > 4 ? values[4].toNumber()?.toDouble() ?? 0 : 0.0;
    final type = args.length > 5 ? values[5].toNumber()?.toInt() ?? 0 : 0;

    if (rate == null || per == null || nper == null || pv == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (per < 1 || per > nper) {
      return const FormulaValue.error(FormulaError.num);
    }

    final payment = _pmt(rate, nper, pv, fv, type);
    final interest = _ipmt(rate, per, nper, pv, fv, type);
    return FormulaValue.number(payment - interest);
  }
}

/// CUMIPMT(rate, nper, pv, start, end, type) - Cumulative interest.
class CumipmtFunction extends FormulaFunction {
  @override
  String get name => 'CUMIPMT';
  @override
  int get minArgs => 6;
  @override
  int get maxArgs => 6;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final rate = values[0].toNumber()?.toDouble();
    final nper = values[1].toNumber()?.toDouble();
    final pv = values[2].toNumber()?.toDouble();
    final startPer = values[3].toNumber()?.toInt();
    final endPer = values[4].toNumber()?.toInt();
    final type = values[5].toNumber()?.toInt();

    if (rate == null ||
        nper == null ||
        pv == null ||
        startPer == null ||
        endPer == null ||
        type == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (rate <= 0 || nper <= 0 || pv <= 0) {
      return const FormulaValue.error(FormulaError.num);
    }
    if (startPer < 1 || endPer < 1 || startPer > endPer) {
      return const FormulaValue.error(FormulaError.num);
    }

    var cumInterest = 0.0;
    for (var per = startPer; per <= endPer; per++) {
      cumInterest += _ipmt(rate, per.toDouble(), nper, pv, 0, type);
    }
    return FormulaValue.number(cumInterest);
  }
}

/// CUMPRINC(rate, nper, pv, start, end, type) - Cumulative principal.
class CumprincFunction extends FormulaFunction {
  @override
  String get name => 'CUMPRINC';
  @override
  int get minArgs => 6;
  @override
  int get maxArgs => 6;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final rate = values[0].toNumber()?.toDouble();
    final nper = values[1].toNumber()?.toDouble();
    final pv = values[2].toNumber()?.toDouble();
    final startPer = values[3].toNumber()?.toInt();
    final endPer = values[4].toNumber()?.toInt();
    final type = values[5].toNumber()?.toInt();

    if (rate == null ||
        nper == null ||
        pv == null ||
        startPer == null ||
        endPer == null ||
        type == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (rate <= 0 || nper <= 0 || pv <= 0) {
      return const FormulaValue.error(FormulaError.num);
    }
    if (startPer < 1 || endPer < 1 || startPer > endPer) {
      return const FormulaValue.error(FormulaError.num);
    }

    final payment = _pmt(rate, nper, pv, 0, type);
    var cumPrinc = 0.0;
    for (var per = startPer; per <= endPer; per++) {
      final interest = _ipmt(rate, per.toDouble(), nper, pv, 0, type);
      cumPrinc += payment - interest;
    }
    return FormulaValue.number(cumPrinc);
  }
}

// ═══════════════════════════════════════════════════════════════
// Wave 3 — Iterative Solvers
// ═══════════════════════════════════════════════════════════════

/// RATE(nper, pmt, pv, [fv], [type], [guess]) - Interest rate per period.
class RateFunction extends FormulaFunction {
  @override
  String get name => 'RATE';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 6;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final nper = values[0].toNumber()?.toDouble();
    final pmt = values[1].toNumber()?.toDouble();
    final pv = values[2].toNumber()?.toDouble();
    final fv = args.length > 3 ? values[3].toNumber()?.toDouble() ?? 0 : 0.0;
    final type = args.length > 4 ? values[4].toNumber()?.toInt() ?? 0 : 0;
    final guess =
        args.length > 5 ? values[5].toNumber()?.toDouble() ?? 0.1 : 0.1;

    if (nper == null || pmt == null || pv == null) {
      return const FormulaValue.error(FormulaError.value);
    }

    double f(double rate) {
      if (rate.abs() < 1e-14) {
        return pv + pmt * nper + fv;
      }
      final pvif = math.pow(1 + rate, nper);
      return pv * pvif +
          pmt * (1 + rate * type) * (pvif - 1) / rate +
          fv;
    }

    double df(double rate) {
      if (rate.abs() < 1e-14) {
        return pmt * nper * (nper - 1) / 2 + pv * nper;
      }
      final pvif = math.pow(1 + rate, nper);
      final dpvif = nper * math.pow(1 + rate, nper - 1);
      return pv * dpvif +
          pmt *
              (type * (pvif - 1) / rate +
                  (1 + rate * type) *
                      (dpvif * rate - (pvif - 1)) /
                      (rate * rate));
    }

    final result = _newtonRaphson(f, df, guess);
    if (result == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(result);
  }
}

/// IRR(values, [guess]) - Internal rate of return.
class IrrFunction extends FormulaFunction {
  @override
  String get name => 'IRR';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final cashFlows = _collectValuesFromArg(args[0], context);
    final guess = args.length > 1
        ? args[1].evaluate(context).toNumber()?.toDouble() ?? 0.1
        : 0.1;

    if (cashFlows.isEmpty) return const FormulaValue.error(FormulaError.num);

    double f(double rate) {
      var npv = 0.0;
      for (var i = 0; i < cashFlows.length; i++) {
        npv += cashFlows[i] / math.pow(1 + rate, i);
      }
      return npv;
    }

    double df(double rate) {
      var d = 0.0;
      for (var i = 1; i < cashFlows.length; i++) {
        d -= i * cashFlows[i] / math.pow(1 + rate, i + 1);
      }
      return d;
    }

    final result = _newtonRaphson(f, df, guess);
    if (result == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(result);
  }
}

/// XNPV(rate, values, dates) - NPV for irregular cash flows.
class XnpvFunction extends FormulaFunction {
  @override
  String get name => 'XNPV';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final rateValue = args[0].evaluate(context);
    final rate = rateValue.toNumber()?.toDouble();
    if (rate == null) return const FormulaValue.error(FormulaError.value);

    final cashFlows = _collectValuesFromArg(args[1], context);
    final dates = _collectValuesFromArg(args[2], context);

    if (cashFlows.length != dates.length || cashFlows.isEmpty) {
      return const FormulaValue.error(FormulaError.num);
    }

    final d0 = dates[0];
    var npv = 0.0;
    for (var i = 0; i < cashFlows.length; i++) {
      npv += cashFlows[i] / math.pow(1 + rate, (dates[i] - d0) / 365);
    }
    return FormulaValue.number(npv);
  }
}

/// XIRR(values, dates, [guess]) - IRR for irregular cash flows.
class XirrFunction extends FormulaFunction {
  @override
  String get name => 'XIRR';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final cashFlows = _collectValuesFromArg(args[0], context);
    final dates = _collectValuesFromArg(args[1], context);
    final guess = args.length > 2
        ? args[2].evaluate(context).toNumber()?.toDouble() ?? 0.1
        : 0.1;

    if (cashFlows.length != dates.length || cashFlows.isEmpty) {
      return const FormulaValue.error(FormulaError.num);
    }

    final d0 = dates[0];

    double f(double rate) {
      var npv = 0.0;
      for (var i = 0; i < cashFlows.length; i++) {
        final t = (dates[i] - d0) / 365;
        npv += cashFlows[i] / math.pow(1 + rate, t);
      }
      return npv;
    }

    double df(double rate) {
      var d = 0.0;
      for (var i = 0; i < cashFlows.length; i++) {
        final t = (dates[i] - d0) / 365;
        d -= t * cashFlows[i] / math.pow(1 + rate, t + 1);
      }
      return d;
    }

    final result = _newtonRaphson(f, df, guess);
    if (result == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(result);
  }
}

// ═══════════════════════════════════════════════════════════════
// Wave 4 — MIRR + Complex Depreciation
// ═══════════════════════════════════════════════════════════════

/// MIRR(values, finance_rate, reinvest_rate) - Modified internal rate of return.
class MirrFunction extends FormulaFunction {
  @override
  String get name => 'MIRR';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final cashFlows = _collectValuesFromArg(args[0], context);
    final values = evaluateArgs(args, context);
    final financeRate = values[1].toNumber()?.toDouble();
    final reinvestRate = values[2].toNumber()?.toDouble();

    if (financeRate == null || reinvestRate == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (cashFlows.isEmpty) return const FormulaValue.error(FormulaError.num);

    final n = cashFlows.length;

    // NPV of negative cash flows at finance rate
    var npvNeg = 0.0;
    var npvPos = 0.0;
    for (var i = 0; i < n; i++) {
      if (cashFlows[i] < 0) {
        npvNeg += cashFlows[i] / math.pow(1 + financeRate, i);
      } else if (cashFlows[i] > 0) {
        // FV of positive cash flows at reinvest rate
        npvPos += cashFlows[i] * math.pow(1 + reinvestRate, n - 1 - i);
      }
    }

    if (npvNeg == 0 || npvPos == 0) {
      return const FormulaValue.error(FormulaError.divZero);
    }

    return FormulaValue.number(
        math.pow(-npvPos / npvNeg, 1.0 / (n - 1)) - 1);
  }
}

/// DB(cost, salvage, life, period, [month]) - Fixed declining balance.
class DbFunction extends FormulaFunction {
  @override
  String get name => 'DB';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 5;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final cost = values[0].toNumber()?.toDouble();
    final salvage = values[1].toNumber()?.toDouble();
    final life = values[2].toNumber()?.toInt();
    final period = values[3].toNumber()?.toInt();
    final month = args.length > 4 ? values[4].toNumber()?.toInt() ?? 12 : 12;

    if (cost == null || salvage == null || life == null || period == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (life <= 0 || period <= 0) {
      return const FormulaValue.error(FormulaError.num);
    }
    if (cost == 0) return const FormulaValue.number(0);

    // Rate rounded to 3 decimal places
    final rate =
        (1 - math.pow(salvage / cost, 1.0 / life)).toDouble();
    final rateRounded = (rate * 1000).round() / 1000;

    var totalDepreciation = 0.0;
    for (var p = 1; p <= period; p++) {
      double depreciation;
      if (p == 1) {
        depreciation = cost * rateRounded * month / 12;
      } else if (p == life + 1) {
        // Last partial year
        depreciation =
            (cost - totalDepreciation) * rateRounded * (12 - month) / 12;
      } else {
        depreciation = (cost - totalDepreciation) * rateRounded;
      }
      if (p == period) return FormulaValue.number(depreciation);
      totalDepreciation += depreciation;
    }

    return const FormulaValue.number(0);
  }
}

/// DDB(cost, salvage, life, period, [factor]) - Double declining balance.
class DdbFunction extends FormulaFunction {
  @override
  String get name => 'DDB';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 5;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final cost = values[0].toNumber()?.toDouble();
    final salvage = values[1].toNumber()?.toDouble();
    final life = values[2].toNumber()?.toDouble();
    final period = values[3].toNumber()?.toDouble();
    final factor =
        args.length > 4 ? values[4].toNumber()?.toDouble() ?? 2 : 2.0;

    if (cost == null ||
        salvage == null ||
        life == null ||
        period == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (life <= 0 || period <= 0 || period > life) {
      return const FormulaValue.error(FormulaError.num);
    }

    final rate = factor / life;
    var totalDepreciation = 0.0;
    for (var p = 1; p <= period; p++) {
      var depreciation = (cost - totalDepreciation) * rate;
      // Cannot depreciate below salvage
      if (totalDepreciation + depreciation > cost - salvage) {
        depreciation = cost - salvage - totalDepreciation;
        if (depreciation < 0) depreciation = 0;
      }
      if (p == period.toInt() && period == period.toInt().toDouble()) {
        return FormulaValue.number(depreciation);
      }
      totalDepreciation += depreciation;
    }

    return const FormulaValue.number(0);
  }
}

/// VDB(cost, salvage, life, start_period, end_period, [factor], [no_switch])
class VdbFunction extends FormulaFunction {
  @override
  String get name => 'VDB';
  @override
  int get minArgs => 5;
  @override
  int get maxArgs => 7;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final cost = values[0].toNumber()?.toDouble();
    final salvage = values[1].toNumber()?.toDouble();
    final life = values[2].toNumber()?.toDouble();
    final startPeriod = values[3].toNumber()?.toDouble();
    final endPeriod = values[4].toNumber()?.toDouble();
    final factor =
        args.length > 5 ? values[5].toNumber()?.toDouble() ?? 2 : 2.0;
    final noSwitch =
        args.length > 6 ? values[6].isTruthy : false;

    if (cost == null ||
        salvage == null ||
        life == null ||
        startPeriod == null ||
        endPeriod == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (life <= 0 || startPeriod < 0 || endPeriod < startPeriod) {
      return const FormulaValue.error(FormulaError.num);
    }

    final rate = factor / life;
    var totalDepreciation = 0.0;
    var accumulated = 0.0;

    // Walk through each unit period up to endPeriod
    final endInt =
        endPeriod.ceil() > life.ceil() ? life.ceil() : endPeriod.ceil();

    for (var p = 1; p <= endInt; p++) {
      // DDB depreciation for this period
      var ddbDep = (cost - accumulated) * rate;
      if (accumulated + ddbDep > cost - salvage) {
        ddbDep = cost - salvage - accumulated;
        if (ddbDep < 0) ddbDep = 0;
      }

      // SLN depreciation for remaining life
      double slnDep;
      if (!noSwitch) {
        final remainingLife = life - p + 1;
        if (remainingLife > 0) {
          slnDep = (cost - salvage - accumulated) / remainingLife;
          if (slnDep < 0) slnDep = 0;
        } else {
          slnDep = 0;
        }
      } else {
        slnDep = 0;
      }

      final dep = noSwitch ? ddbDep : math.max(ddbDep, slnDep);

      // Handle fractional start and end periods
      final periodStart = (p - 1).toDouble();
      final periodEnd = p.toDouble();

      if (periodEnd <= startPeriod) {
        // Before our window
        accumulated += dep;
        continue;
      }
      if (periodStart >= endPeriod) {
        break;
      }

      // Calculate fractional overlap
      final overlapStart = math.max(periodStart, startPeriod);
      final overlapEnd = math.min(periodEnd, endPeriod);
      final fraction = overlapEnd - overlapStart;

      totalDepreciation += dep * fraction;
      accumulated += dep;
    }

    return FormulaValue.number(totalDepreciation);
  }
}

// ═══════════════════════════════════════════════════════════════
// Wave 5 — Simple Bond/Security Functions
// ═══════════════════════════════════════════════════════════════

/// DISC(settlement, maturity, pr, redemption, [basis])
class DiscFunction extends FormulaFunction {
  @override
  String get name => 'DISC';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 5;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final settlement = values[0].toNumber()?.toInt();
    final maturity = values[1].toNumber()?.toInt();
    final pr = values[2].toNumber()?.toDouble();
    final redemption = values[3].toNumber()?.toDouble();
    final basis = args.length > 4 ? values[4].toNumber()?.toInt() ?? 0 : 0;

    if (settlement == null ||
        maturity == null ||
        pr == null ||
        redemption == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (basis < 0 || basis > 4) {
      return const FormulaValue.error(FormulaError.num);
    }

    final startDate = _serialToDate(settlement);
    final endDate = _serialToDate(maturity);
    final yf = _yearFrac(startDate, endDate, basis);
    if (yf == 0) return const FormulaValue.error(FormulaError.divZero);

    return FormulaValue.number((redemption - pr) / redemption / yf);
  }
}

/// INTRATE(settlement, maturity, investment, redemption, [basis])
class IntrateFunction extends FormulaFunction {
  @override
  String get name => 'INTRATE';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 5;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final settlement = values[0].toNumber()?.toInt();
    final maturity = values[1].toNumber()?.toInt();
    final investment = values[2].toNumber()?.toDouble();
    final redemption = values[3].toNumber()?.toDouble();
    final basis = args.length > 4 ? values[4].toNumber()?.toInt() ?? 0 : 0;

    if (settlement == null ||
        maturity == null ||
        investment == null ||
        redemption == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (investment <= 0) return const FormulaValue.error(FormulaError.num);
    if (basis < 0 || basis > 4) {
      return const FormulaValue.error(FormulaError.num);
    }

    final startDate = _serialToDate(settlement);
    final endDate = _serialToDate(maturity);
    final yf = _yearFrac(startDate, endDate, basis);
    if (yf == 0) return const FormulaValue.error(FormulaError.divZero);

    return FormulaValue.number(
        (redemption - investment) / investment / yf);
  }
}

/// RECEIVED(settlement, maturity, investment, discount, [basis])
class ReceivedFunction extends FormulaFunction {
  @override
  String get name => 'RECEIVED';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 5;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final settlement = values[0].toNumber()?.toInt();
    final maturity = values[1].toNumber()?.toInt();
    final investment = values[2].toNumber()?.toDouble();
    final discount = values[3].toNumber()?.toDouble();
    final basis = args.length > 4 ? values[4].toNumber()?.toInt() ?? 0 : 0;

    if (settlement == null ||
        maturity == null ||
        investment == null ||
        discount == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (basis < 0 || basis > 4) {
      return const FormulaValue.error(FormulaError.num);
    }

    final startDate = _serialToDate(settlement);
    final endDate = _serialToDate(maturity);
    final yf = _yearFrac(startDate, endDate, basis);
    final denom = 1 - discount * yf;
    if (denom == 0) return const FormulaValue.error(FormulaError.divZero);

    return FormulaValue.number(investment / denom);
  }
}

/// PRICEDISC(settlement, maturity, discount, redemption, [basis])
class PriceDiscFunction extends FormulaFunction {
  @override
  String get name => 'PRICEDISC';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 5;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final settlement = values[0].toNumber()?.toInt();
    final maturity = values[1].toNumber()?.toInt();
    final discount = values[2].toNumber()?.toDouble();
    final redemption = values[3].toNumber()?.toDouble();
    final basis = args.length > 4 ? values[4].toNumber()?.toInt() ?? 0 : 0;

    if (settlement == null ||
        maturity == null ||
        discount == null ||
        redemption == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (basis < 0 || basis > 4) {
      return const FormulaValue.error(FormulaError.num);
    }

    final startDate = _serialToDate(settlement);
    final endDate = _serialToDate(maturity);
    final yf = _yearFrac(startDate, endDate, basis);

    return FormulaValue.number(
        redemption - discount * redemption * yf);
  }
}

/// PRICEMAT(settlement, maturity, issue, rate, yld, [basis])
class PriceMatFunction extends FormulaFunction {
  @override
  String get name => 'PRICEMAT';
  @override
  int get minArgs => 5;
  @override
  int get maxArgs => 6;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final settlement = values[0].toNumber()?.toInt();
    final maturity = values[1].toNumber()?.toInt();
    final issue = values[2].toNumber()?.toInt();
    final rate = values[3].toNumber()?.toDouble();
    final yld = values[4].toNumber()?.toDouble();
    final basis = args.length > 5 ? values[5].toNumber()?.toInt() ?? 0 : 0;

    if (settlement == null ||
        maturity == null ||
        issue == null ||
        rate == null ||
        yld == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (basis < 0 || basis > 4) {
      return const FormulaValue.error(FormulaError.num);
    }

    final settlementDate = _serialToDate(settlement);
    final maturityDate = _serialToDate(maturity);
    final issueDate = _serialToDate(issue);

    final b = _yearFrac(issueDate, maturityDate, basis);
    final dsm = _yearFrac(settlementDate, maturityDate, basis);
    final a = _yearFrac(issueDate, settlementDate, basis);

    final denom = 1 + dsm * yld;
    if (denom == 0) return const FormulaValue.error(FormulaError.divZero);

    return FormulaValue.number(
        100 * (1 + b * rate) / denom - 100 * a * rate);
  }
}

/// ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis])
class AccrintFunction extends FormulaFunction {
  @override
  String get name => 'ACCRINT';
  @override
  int get minArgs => 6;
  @override
  int get maxArgs => 7;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final issue = values[0].toNumber()?.toInt();
    // first_interest at index 1 (used for multi-period accrual but simplified here)
    final settlement = values[2].toNumber()?.toInt();
    final rate = values[3].toNumber()?.toDouble();
    final par = values[4].toNumber()?.toDouble();
    final frequency = values[5].toNumber()?.toInt();
    final basis = args.length > 6 ? values[6].toNumber()?.toInt() ?? 0 : 0;

    if (issue == null ||
        settlement == null ||
        rate == null ||
        par == null ||
        frequency == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (frequency != 1 && frequency != 2 && frequency != 4) {
      return const FormulaValue.error(FormulaError.num);
    }
    if (basis < 0 || basis > 4) {
      return const FormulaValue.error(FormulaError.num);
    }

    final issueDate = _serialToDate(issue);
    final settlementDate = _serialToDate(settlement);
    final yf = _yearFrac(issueDate, settlementDate, basis);

    return FormulaValue.number(par * rate * yf);
  }
}

// ═══════════════════════════════════════════════════════════════
// Wave 6 — Complex Bond Functions
// ═══════════════════════════════════════════════════════════════

/// Compute bond price from yield (shared by PRICE and YIELD).
double _bondPrice(DateTime settlement, DateTime maturity, double rate,
    double yld, double redemption, int frequency, int basis) {
  final n = _coupNum(settlement, maturity, frequency);
  final dsc = _coupDays(settlement, maturity, frequency, basis) -
      _coupDaysBs(settlement, maturity, frequency, basis);
  final e = _coupDays(settlement, maturity, frequency, basis);
  final a = _coupDaysBs(settlement, maturity, frequency, basis);

  if (e == 0) return 0;

  final coupon = 100 * rate / frequency;
  final dscFrac = dsc.toDouble() / e.toDouble();

  if (n == 1) {
    // Special case for last coupon period
    final t1 = coupon + redemption;
    final t2 = yld / frequency * dscFrac + 1;
    final t3 = coupon * a / e;
    return t1 / t2 - t3;
  }

  var price = redemption / math.pow(1 + yld / frequency, n - 1 + dscFrac);
  for (var k = 1; k <= n; k++) {
    price +=
        coupon / math.pow(1 + yld / frequency, k - 1 + dscFrac);
  }
  price -= coupon * a / e;

  return price;
}

/// PRICE(settlement, maturity, rate, yld, redemption, frequency, [basis])
class PriceFunction extends FormulaFunction {
  @override
  String get name => 'PRICE';
  @override
  int get minArgs => 6;
  @override
  int get maxArgs => 7;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final settlement = values[0].toNumber()?.toInt();
    final maturity = values[1].toNumber()?.toInt();
    final rate = values[2].toNumber()?.toDouble();
    final yld = values[3].toNumber()?.toDouble();
    final redemption = values[4].toNumber()?.toDouble();
    final frequency = values[5].toNumber()?.toInt();
    final basis = args.length > 6 ? values[6].toNumber()?.toInt() ?? 0 : 0;

    if (settlement == null ||
        maturity == null ||
        rate == null ||
        yld == null ||
        redemption == null ||
        frequency == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (frequency != 1 && frequency != 2 && frequency != 4) {
      return const FormulaValue.error(FormulaError.num);
    }
    if (basis < 0 || basis > 4) {
      return const FormulaValue.error(FormulaError.num);
    }

    final settlementDate = _serialToDate(settlement);
    final maturityDate = _serialToDate(maturity);

    return FormulaValue.number(
        _bondPrice(settlementDate, maturityDate, rate, yld, redemption,
            frequency, basis));
  }
}

/// YIELD(settlement, maturity, rate, pr, redemption, frequency, [basis])
class YieldFunction extends FormulaFunction {
  @override
  String get name => 'YIELD';
  @override
  int get minArgs => 6;
  @override
  int get maxArgs => 7;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final settlement = values[0].toNumber()?.toInt();
    final maturity = values[1].toNumber()?.toInt();
    final rate = values[2].toNumber()?.toDouble();
    final pr = values[3].toNumber()?.toDouble();
    final redemption = values[4].toNumber()?.toDouble();
    final frequency = values[5].toNumber()?.toInt();
    final basis = args.length > 6 ? values[6].toNumber()?.toInt() ?? 0 : 0;

    if (settlement == null ||
        maturity == null ||
        rate == null ||
        pr == null ||
        redemption == null ||
        frequency == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (frequency != 1 && frequency != 2 && frequency != 4) {
      return const FormulaValue.error(FormulaError.num);
    }
    if (basis < 0 || basis > 4) {
      return const FormulaValue.error(FormulaError.num);
    }

    final settlementDate = _serialToDate(settlement);
    final maturityDate = _serialToDate(maturity);

    double f(double yld) {
      return _bondPrice(
              settlementDate, maturityDate, rate, yld, redemption,
              frequency, basis) -
          pr;
    }

    double df(double yld) {
      const h = 1e-7;
      return (f(yld + h) - f(yld - h)) / (2 * h);
    }

    // Start guess at coupon rate
    final result = _newtonRaphson(f, df, rate);
    if (result == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(result);
  }
}

/// DURATION(settlement, maturity, coupon, yld, frequency, [basis])
class DurationFunction extends FormulaFunction {
  @override
  String get name => 'DURATION';
  @override
  int get minArgs => 5;
  @override
  int get maxArgs => 6;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final settlement = values[0].toNumber()?.toInt();
    final maturity = values[1].toNumber()?.toInt();
    final coupon = values[2].toNumber()?.toDouble();
    final yld = values[3].toNumber()?.toDouble();
    final frequency = values[4].toNumber()?.toInt();
    final basis = args.length > 5 ? values[5].toNumber()?.toInt() ?? 0 : 0;

    if (settlement == null ||
        maturity == null ||
        coupon == null ||
        yld == null ||
        frequency == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (frequency != 1 && frequency != 2 && frequency != 4) {
      return const FormulaValue.error(FormulaError.num);
    }
    if (basis < 0 || basis > 4) {
      return const FormulaValue.error(FormulaError.num);
    }

    final settlementDate = _serialToDate(settlement);
    final maturityDate = _serialToDate(maturity);

    final n = _coupNum(settlementDate, maturityDate, frequency);
    final dsc = _coupDays(settlementDate, maturityDate, frequency, basis) -
        _coupDaysBs(settlementDate, maturityDate, frequency, basis);
    final e = _coupDays(settlementDate, maturityDate, frequency, basis);

    if (e == 0) return const FormulaValue.error(FormulaError.divZero);

    final c = 100 * coupon / frequency;
    final dscFrac = dsc.toDouble() / e.toDouble();

    var weightedCf = 0.0;
    var totalCf = 0.0;

    for (var k = 1; k <= n; k++) {
      final t = (k - 1 + dscFrac) / frequency;
      final discFactor =
          1 / math.pow(1 + yld / frequency, k - 1 + dscFrac);
      final cf = k == n ? c + 100 : c;
      weightedCf += t * cf * discFactor;
      totalCf += cf * discFactor;
    }

    if (totalCf == 0) return const FormulaValue.error(FormulaError.divZero);

    return FormulaValue.number(weightedCf / totalCf);
  }
}

/// MDURATION(settlement, maturity, coupon, yld, frequency, [basis])
class MdurationFunction extends FormulaFunction {
  @override
  String get name => 'MDURATION';
  @override
  int get minArgs => 5;
  @override
  int get maxArgs => 6;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final yld = values[3].toNumber()?.toDouble();
    final frequency = values[4].toNumber()?.toInt();

    if (yld == null || frequency == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (frequency != 1 && frequency != 2 && frequency != 4) {
      return const FormulaValue.error(FormulaError.num);
    }

    // Compute Macaulay duration first
    final durationFn = DurationFunction();
    final duration = durationFn.call(args, context);
    if (duration.isError) return duration;

    final dur = (duration as NumberValue).value.toDouble();
    return FormulaValue.number(dur / (1 + yld / frequency));
  }
}
