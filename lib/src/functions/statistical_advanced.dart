import 'dart:math' as math;

import '../ast/nodes.dart';
import '../evaluation/context.dart';
import '../evaluation/errors.dart';
import '../evaluation/value.dart';
import 'function.dart';
import 'registry.dart';

/// Register all advanced statistical functions.
void registerAdvancedStatisticalFunctions(FunctionRegistry registry) {
  registry.registerAll([
    FisherFunction(),
    FisherInvFunction(),
    StandardizeFunction(),
    PermutFunction(),
    PermutationaFunction(),
    DevsqFunction(),
    KurtFunction(),
    SkewFunction(),
    SkewPFunction(),
    CovariancePFunction(),
    CovarianceSFunction(),
    CorrelFunction(),
    PearsonFunction(),
    RsqFunction(),
    SlopeFunction(),
    InterceptFunction(),
    SteyxFunction(),
    ForecastLinearFunction(),
    ProbFunction(),
    ModeMultFunction(),
    StdevaFunction(),
    StdevpaFunction(),
    VaraFunction(),
    VarpaFunction(),
    GammaFunction(),
    GammaLnFunction(),
    GammaLnPreciseFunction(),
    GaussFunction(),
    PhiFunction(),
    NormSDistFunction(),
    NormSInvFunction(),
    NormDistFunction(),
    NormInvFunction(),
    BinomDistFunction(),
    BinomInvFunction(),
    BinomDistRangeFunction(),
    NegBinomDistFunction(),
    HypgeomDistFunction(),
    PoissonDistFunction(),
    ExponDistFunction(),
    GammaDistFunction(),
    GammaInvFunction(),
    BetaDistFunction(),
    BetaInvFunction(),
    ChisqDistFunction(),
    ChisqInvFunction(),
    ChisqDistRtFunction(),
    ChisqInvRtFunction(),
    TDistFunction(),
    TInvFunction(),
    TDist2TFunction(),
    TInv2TFunction(),
    TDistRtFunction(),
    FDistFunction(),
    FInvFunction(),
    FDistRtFunction(),
    FInvRtFunction(),
    WeibullDistFunction(),
    LognormDistFunction(),
    LognormInvFunction(),
    ConfidenceNormFunction(),
    ConfidenceTFunction(),
    ZTestFunction(),
    TTestFunction(),
    ChisqTestFunction(),
    FTestFunction(),
    LinestFunction(),
    LogestFunction(),
    TrendFunction(),
    GrowthFunction(),
  ]);
}

// =============================================================================
// Private helpers — math primitives
// =============================================================================

const _lanczosG = 7;
const _lanczosCoefficients = <double>[
  0.99999999999980993,
  676.5203681218851,
  -1259.1392167224028,
  771.32342877765313,
  -176.61502916214059,
  12.507343278686905,
  -0.13857109526572012,
  9.9843695780195716e-6,
  1.5056327351493116e-7,
];

double _lnGamma(double x) {
  if (x < 0.5) {
    return math.log(math.pi / math.sin(math.pi * x)) - _lnGamma(1 - x);
  }
  x -= 1;
  var a = _lanczosCoefficients[0];
  final t = x + _lanczosG + 0.5;
  for (var i = 1; i < _lanczosCoefficients.length; i++) {
    a += _lanczosCoefficients[i] / (x + i);
  }
  return 0.5 * math.log(2 * math.pi) +
      (x + 0.5) * math.log(t) -
      t +
      math.log(a);
}

double _gamma(double x) {
  if (x <= 0 && x == x.truncateToDouble()) return double.infinity;
  return math.exp(_lnGamma(x));
}

double _beta(double a, double b) =>
    math.exp(_lnGamma(a) + _lnGamma(b) - _lnGamma(a + b));

double _erf(double x) {
  final sign = x < 0 ? -1.0 : 1.0;
  x = x.abs();
  const a1 = 0.254829592;
  const a2 = -0.284496736;
  const a3 = 1.421413741;
  const a4 = -1.453152027;
  const a5 = 1.061405429;
  const p = 0.3275911;
  final t = 1.0 / (1.0 + p * x);
  final y = 1.0 -
      (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) *
          t *
          math.exp(-x * x);
  return sign * y;
}

double _regularizedGammaP(double a, double x) {
  if (x < 0 || a <= 0) return 0;
  if (x == 0) return 0;
  if (x < a + 1) {
    var sum = 1.0 / a;
    var term = 1.0 / a;
    for (var n = 1; n < 200; n++) {
      term *= x / (a + n);
      sum += term;
      if (term.abs() < sum.abs() * 1e-14) break;
    }
    return sum * math.exp(-x + a * math.log(x) - _lnGamma(a));
  } else {
    return 1.0 - _regularizedGammaQ(a, x);
  }
}

double _regularizedGammaQ(double a, double x) {
  if (x < a + 1) return 1.0 - _regularizedGammaP(a, x);
  const tiny = 1e-30;
  var f = tiny;
  var c = 1.0 / tiny;
  var d = 1.0 / (x - a + 1);
  f = d;
  for (var n = 1; n < 200; n++) {
    final an = -n * (n - a);
    final bn = x - a + 1 + 2 * n;
    d = bn + an * d;
    if (d.abs() < tiny) d = tiny;
    c = bn + an / c;
    if (c.abs() < tiny) c = tiny;
    d = 1.0 / d;
    final delta = c * d;
    f *= delta;
    if ((delta - 1.0).abs() < 1e-14) break;
  }
  return f * math.exp(-x + a * math.log(x) - _lnGamma(a));
}

double _regularizedBeta(double x, double a, double b) {
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  if (x > (a + 1) / (a + b + 2)) {
    return 1.0 - _regularizedBeta(1 - x, b, a);
  }
  final lnBeta = _lnGamma(a + b) - _lnGamma(a) - _lnGamma(b);
  final front =
      math.exp(lnBeta + a * math.log(x) + b * math.log(1 - x)) / a;
  const tiny = 1e-30;
  var f = tiny;
  var c = 1.0;
  var d = 1.0 - (a + b) * x / (a + 1);
  if (d.abs() < tiny) d = tiny;
  d = 1.0 / d;
  f = d;
  for (var m = 1; m < 200; m++) {
    var numerator = m * (b - m) * x / ((a + 2 * m - 1) * (a + 2 * m));
    d = 1.0 + numerator * d;
    if (d.abs() < tiny) d = tiny;
    c = 1.0 + numerator / c;
    if (c.abs() < tiny) c = tiny;
    d = 1.0 / d;
    f *= c * d;
    numerator =
        -(a + m) * (a + b + m) * x / ((a + 2 * m) * (a + 2 * m + 1));
    d = 1.0 + numerator * d;
    if (d.abs() < tiny) d = tiny;
    c = 1.0 + numerator / c;
    if (c.abs() < tiny) c = tiny;
    d = 1.0 / d;
    final delta = c * d;
    f *= delta;
    if ((delta - 1.0).abs() < 1e-14) break;
  }
  return front * f;
}

double? _inverseCdf(
    double Function(double) cdf, double p, double lo, double hi,
    {double tol = 1e-10, int maxIter = 200}) {
  for (var i = 0; i < maxIter; i++) {
    final mid = (lo + hi) / 2;
    if ((hi - lo) < tol) return mid;
    if (cdf(mid) < p) {
      lo = mid;
    } else {
      hi = mid;
    }
  }
  return (lo + hi) / 2;
}

// Distribution helpers
double _normSCdf(double z) => 0.5 * (1 + _erf(z / math.sqrt(2)));
double _normSPdf(double z) =>
    math.exp(-0.5 * z * z) / math.sqrt(2 * math.pi);

double _normSInv(double p) => _inverseCdf(_normSCdf, p, -8, 8)!;

double _tCdf(double x, double df) {
  final t2 = x * x;
  final ib = _regularizedBeta(df / (df + t2), df / 2, 0.5);
  return x >= 0 ? 1.0 - 0.5 * ib : 0.5 * ib;
}

double _fCdf(double x, double d1, double d2) =>
    _regularizedBeta(d1 * x / (d1 * x + d2), d1 / 2, d2 / 2);

// Paired array extraction
(List<double>, List<double>)? _extractPairedArrays(
    FormulaNode arg1, FormulaNode arg2, EvaluationContext context) {
  final v1 = arg1.evaluate(context);
  final v2 = arg2.evaluate(context);
  if (v1 is! RangeValue || v2 is! RangeValue) return null;
  final flat1 = v1.flat.toList();
  final flat2 = v2.flat.toList();
  if (flat1.length != flat2.length) return null;
  final xs = <double>[];
  final ys = <double>[];
  for (var i = 0; i < flat1.length; i++) {
    final n1 = flat1[i].toNumber();
    final n2 = flat2[i].toNumber();
    if (n1 != null &&
        n2 != null &&
        flat1[i] is NumberValue &&
        flat2[i] is NumberValue) {
      xs.add(n1.toDouble());
      ys.add(n2.toDouble());
    }
  }
  if (xs.isEmpty) return null;
  return (xs, ys);
}

// Flatten arguments to numbers (skip non-numeric)
List<num> _flattenToNumbers(
    List<FormulaNode> args, EvaluationContext context) {
  final result = <num>[];
  for (final arg in args) {
    final value = arg.evaluate(context);
    switch (value) {
      case NumberValue(value: final n):
        result.add(n);
      case RangeValue(values: final matrix):
        for (final row in matrix) {
          for (final cell in row) {
            if (cell is NumberValue) result.add(cell.value);
          }
        }
      default:
        break;
    }
  }
  return result;
}

// Collect all values: text=0, TRUE=1, FALSE=0
List<num> _collectAllValuesAdvanced(
    List<FormulaNode> args, EvaluationContext context) {
  final result = <num>[];
  for (final arg in args) {
    final value = arg.evaluate(context);
    switch (value) {
      case NumberValue(value: final n):
        result.add(n);
      case BooleanValue(value: final b):
        result.add(b ? 1 : 0);
      case TextValue():
        result.add(0);
      case RangeValue(values: final matrix):
        for (final row in matrix) {
          for (final cell in row) {
            switch (cell) {
              case NumberValue(value: final n):
                result.add(n);
              case BooleanValue(value: final b):
                result.add(b ? 1 : 0);
              case TextValue():
                result.add(0);
              default:
                break;
            }
          }
        }
      default:
        break;
    }
  }
  return result;
}

// Binomial coefficient
double _binomCoeff(int n, int k) {
  if (k < 0 || k > n) return 0;
  if (k == 0 || k == n) return 1;
  if (k > n - k) k = n - k;
  var result = 1.0;
  for (var i = 0; i < k; i++) {
    result *= (n - i) / (i + 1);
  }
  return result;
}

// =============================================================================
// Wave 1 — 24 function classes
// =============================================================================

/// FISHER(x) — Fisher transformation.
class FisherFunction extends FormulaFunction {
  @override
  String get name => 'FISHER';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final x = v.toNumber()?.toDouble();
    if (x == null) return const FormulaValue.error(FormulaError.value);
    if (x.abs() >= 1) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(0.5 * math.log((1 + x) / (1 - x)));
  }
}

/// FISHERINV(y) — Inverse Fisher transformation.
class FisherInvFunction extends FormulaFunction {
  @override
  String get name => 'FISHERINV';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final y = v.toNumber()?.toDouble();
    if (y == null) return const FormulaValue.error(FormulaError.value);
    final e2y = math.exp(2 * y);
    return FormulaValue.number((e2y - 1) / (e2y + 1));
  }
}

/// STANDARDIZE(x, mean, standard_dev) — Normalized value.
class StandardizeFunction extends FormulaFunction {
  @override
  String get name => 'STANDARDIZE';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final x = vals[0].toNumber()?.toDouble();
    final mean = vals[1].toNumber()?.toDouble();
    final stddev = vals[2].toNumber()?.toDouble();
    if (x == null || mean == null || stddev == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (stddev <= 0) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number((x - mean) / stddev);
  }
}

/// PERMUT(number, number_chosen) — Permutations.
class PermutFunction extends FormulaFunction {
  @override
  String get name => 'PERMUT';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final n = vals[0].toNumber()?.toInt();
    final k = vals[1].toNumber()?.toInt();
    if (n == null || k == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (n < 0 || k < 0 || k > n) {
      return const FormulaValue.error(FormulaError.num);
    }
    var result = 1.0;
    for (var i = 0; i < k; i++) {
      result *= (n - i);
    }
    return FormulaValue.number(result);
  }
}

/// PERMUTATIONA(number, number_chosen) — Permutations with repetition.
class PermutationaFunction extends FormulaFunction {
  @override
  String get name => 'PERMUTATIONA';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final n = vals[0].toNumber()?.toInt();
    final k = vals[1].toNumber()?.toInt();
    if (n == null || k == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (n < 0 || k < 0) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(math.pow(n, k));
  }
}

/// DEVSQ(number1, [number2], ...) — Sum of squared deviations.
class DevsqFunction extends FormulaFunction {
  @override
  String get name => 'DEVSQ';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = _flattenToNumbers(args, context);
    if (numbers.isEmpty) return const FormulaValue.error(FormulaError.num);
    final mean = numbers.fold(0.0, (a, b) => a + b) / numbers.length;
    var sum = 0.0;
    for (final x in numbers) {
      final d = x - mean;
      sum += d * d;
    }
    return FormulaValue.number(sum);
  }
}

/// KURT(number1, [number2], ...) — Excess kurtosis.
class KurtFunction extends FormulaFunction {
  @override
  String get name => 'KURT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = _flattenToNumbers(args, context);
    final n = numbers.length;
    if (n < 4) return const FormulaValue.error(FormulaError.divZero);
    final mean = numbers.fold(0.0, (a, b) => a + b) / n;
    var s2 = 0.0;
    for (final x in numbers) {
      final d = x - mean;
      s2 += d * d;
    }
    s2 /= (n - 1);
    final s = math.sqrt(s2);
    if (s == 0) return const FormulaValue.error(FormulaError.divZero);
    var sum4 = 0.0;
    for (final x in numbers) {
      final d = (x - mean) / s;
      sum4 += d * d * d * d;
    }
    final kurt = n *
            (n + 1) /
            ((n - 1) * (n - 2) * (n - 3)) *
            sum4 -
        3 * (n - 1) * (n - 1) / ((n - 2) * (n - 3));
    return FormulaValue.number(kurt);
  }
}

/// SKEW(number1, [number2], ...) — Sample skewness.
class SkewFunction extends FormulaFunction {
  @override
  String get name => 'SKEW';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = _flattenToNumbers(args, context);
    final n = numbers.length;
    if (n < 3) return const FormulaValue.error(FormulaError.divZero);
    final mean = numbers.fold(0.0, (a, b) => a + b) / n;
    var s2 = 0.0;
    for (final x in numbers) {
      final d = x - mean;
      s2 += d * d;
    }
    s2 /= (n - 1);
    final s = math.sqrt(s2);
    if (s == 0) return const FormulaValue.error(FormulaError.divZero);
    var sum3 = 0.0;
    for (final x in numbers) {
      final d = (x - mean) / s;
      sum3 += d * d * d;
    }
    return FormulaValue.number(n / ((n - 1) * (n - 2)) * sum3);
  }
}

/// SKEW.P(number1, [number2], ...) — Population skewness.
class SkewPFunction extends FormulaFunction {
  @override
  String get name => 'SKEW.P';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = _flattenToNumbers(args, context);
    final n = numbers.length;
    if (n < 3) return const FormulaValue.error(FormulaError.divZero);
    final mean = numbers.fold(0.0, (a, b) => a + b) / n;
    var s2 = 0.0;
    for (final x in numbers) {
      final d = x - mean;
      s2 += d * d;
    }
    final sp = math.sqrt(s2 / n);
    if (sp == 0) return const FormulaValue.error(FormulaError.divZero);
    var sum3 = 0.0;
    for (final x in numbers) {
      final d = (x - mean) / sp;
      sum3 += d * d * d;
    }
    return FormulaValue.number(sum3 / n);
  }
}

/// COVARIANCE.P(array1, array2) — Population covariance.
class CovariancePFunction extends FormulaFunction {
  @override
  String get name => 'COVARIANCE.P';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final paired = _extractPairedArrays(args[0], args[1], context);
    if (paired == null) return const FormulaValue.error(FormulaError.na);
    final (xs, ys) = paired;
    final n = xs.length;
    if (n == 0) return const FormulaValue.error(FormulaError.divZero);
    final mx = xs.fold(0.0, (a, b) => a + b) / n;
    final my = ys.fold(0.0, (a, b) => a + b) / n;
    var sum = 0.0;
    for (var i = 0; i < n; i++) {
      sum += (xs[i] - mx) * (ys[i] - my);
    }
    return FormulaValue.number(sum / n);
  }
}

/// COVARIANCE.S(array1, array2) — Sample covariance.
class CovarianceSFunction extends FormulaFunction {
  @override
  String get name => 'COVARIANCE.S';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final paired = _extractPairedArrays(args[0], args[1], context);
    if (paired == null) return const FormulaValue.error(FormulaError.na);
    final (xs, ys) = paired;
    final n = xs.length;
    if (n < 2) return const FormulaValue.error(FormulaError.divZero);
    final mx = xs.fold(0.0, (a, b) => a + b) / n;
    final my = ys.fold(0.0, (a, b) => a + b) / n;
    var sum = 0.0;
    for (var i = 0; i < n; i++) {
      sum += (xs[i] - mx) * (ys[i] - my);
    }
    return FormulaValue.number(sum / (n - 1));
  }
}

/// CORREL(array1, array2) — Pearson correlation coefficient.
class CorrelFunction extends FormulaFunction {
  @override
  String get name => 'CORREL';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final r = _computeCorrel(args, context);
    if (r == null) return const FormulaValue.error(FormulaError.na);
    return FormulaValue.number(r);
  }
}

double? _computeCorrel(List<FormulaNode> args, EvaluationContext context) {
  final paired = _extractPairedArrays(args[0], args[1], context);
  if (paired == null) return null;
  final (xs, ys) = paired;
  final n = xs.length;
  final mx = xs.fold(0.0, (a, b) => a + b) / n;
  final my = ys.fold(0.0, (a, b) => a + b) / n;
  var sumXY = 0.0, sumX2 = 0.0, sumY2 = 0.0;
  for (var i = 0; i < n; i++) {
    final dx = xs[i] - mx;
    final dy = ys[i] - my;
    sumXY += dx * dy;
    sumX2 += dx * dx;
    sumY2 += dy * dy;
  }
  final denom = math.sqrt(sumX2 * sumY2);
  if (denom == 0) return null;
  return sumXY / denom;
}

/// PEARSON(array1, array2) — Same as CORREL.
class PearsonFunction extends FormulaFunction {
  @override
  String get name => 'PEARSON';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final r = _computeCorrel(args, context);
    if (r == null) return const FormulaValue.error(FormulaError.na);
    return FormulaValue.number(r);
  }
}

/// RSQ(known_ys, known_xs) — R-squared.
class RsqFunction extends FormulaFunction {
  @override
  String get name => 'RSQ';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final r = _computeCorrel(args, context);
    if (r == null) return const FormulaValue.error(FormulaError.divZero);
    return FormulaValue.number(r * r);
  }
}

// Shared regression helper: returns (slope, intercept, mx, my, xs, ys)
(double, double, double, double, List<double>, List<double>)?
    _computeRegression(
        FormulaNode ysArg, FormulaNode xsArg, EvaluationContext context) {
  final paired = _extractPairedArrays(xsArg, ysArg, context);
  if (paired == null) return null;
  final (xs, ys) = paired;
  final n = xs.length;
  final mx = xs.fold(0.0, (a, b) => a + b) / n;
  final my = ys.fold(0.0, (a, b) => a + b) / n;
  var sumXY = 0.0, sumX2 = 0.0;
  for (var i = 0; i < n; i++) {
    sumXY += (xs[i] - mx) * (ys[i] - my);
    sumX2 += (xs[i] - mx) * (xs[i] - mx);
  }
  if (sumX2 == 0) return null;
  final slope = sumXY / sumX2;
  final intercept = my - slope * mx;
  return (slope, intercept, mx, my, xs, ys);
}

/// SLOPE(known_ys, known_xs) — Slope of linear regression.
class SlopeFunction extends FormulaFunction {
  @override
  String get name => 'SLOPE';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final reg = _computeRegression(args[0], args[1], context);
    if (reg == null) return const FormulaValue.error(FormulaError.divZero);
    return FormulaValue.number(reg.$1);
  }
}

/// INTERCEPT(known_ys, known_xs) — Intercept of linear regression.
class InterceptFunction extends FormulaFunction {
  @override
  String get name => 'INTERCEPT';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final reg = _computeRegression(args[0], args[1], context);
    if (reg == null) return const FormulaValue.error(FormulaError.divZero);
    return FormulaValue.number(reg.$2);
  }
}

/// STEYX(known_ys, known_xs) — Standard error of regression.
class SteyxFunction extends FormulaFunction {
  @override
  String get name => 'STEYX';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final reg = _computeRegression(args[0], args[1], context);
    if (reg == null) return const FormulaValue.error(FormulaError.divZero);
    final (slope, intercept, _, _, xs, ys) = reg;
    final n = xs.length;
    if (n < 3) return const FormulaValue.error(FormulaError.divZero);
    var sse = 0.0;
    for (var i = 0; i < n; i++) {
      final yhat = intercept + slope * xs[i];
      final e = ys[i] - yhat;
      sse += e * e;
    }
    return FormulaValue.number(math.sqrt(sse / (n - 2)));
  }
}

/// FORECAST.LINEAR(x, known_ys, known_xs) — Predicted value.
class ForecastLinearFunction extends FormulaFunction {
  @override
  String get name => 'FORECAST.LINEAR';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final xVal = args[0].evaluate(context);
    if (xVal.isError) return xVal;
    final x = xVal.toNumber()?.toDouble();
    if (x == null) return const FormulaValue.error(FormulaError.value);
    final reg = _computeRegression(args[1], args[2], context);
    if (reg == null) return const FormulaValue.error(FormulaError.divZero);
    final (slope, intercept, _, _, _, _) = reg;
    return FormulaValue.number(intercept + slope * x);
  }
}

/// PROB(x_range, prob_range, lower_limit, [upper_limit]).
class ProbFunction extends FormulaFunction {
  @override
  String get name => 'PROB';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final paired = _extractPairedArrays(args[0], args[1], context);
    if (paired == null) return const FormulaValue.error(FormulaError.na);
    final (xVals, pVals) = paired;

    final lowerVal = args[2].evaluate(context);
    if (lowerVal.isError) return lowerVal;
    final lower = lowerVal.toNumber()?.toDouble();
    if (lower == null) return const FormulaValue.error(FormulaError.value);

    double upper;
    if (args.length > 3) {
      final upperVal = args[3].evaluate(context);
      if (upperVal.isError) return upperVal;
      final u = upperVal.toNumber()?.toDouble();
      if (u == null) return const FormulaValue.error(FormulaError.value);
      upper = u;
    } else {
      upper = lower;
    }

    var sum = 0.0;
    for (var i = 0; i < xVals.length; i++) {
      if (xVals[i] >= lower && xVals[i] <= upper) {
        sum += pVals[i];
      }
    }
    return FormulaValue.number(sum);
  }
}

/// MODE.MULT(number1, [number2], ...) — All modal values.
class ModeMultFunction extends FormulaFunction {
  @override
  String get name => 'MODE.MULT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = _flattenToNumbers(args, context);
    if (numbers.isEmpty) return const FormulaValue.error(FormulaError.na);

    final freq = <num, int>{};
    for (final n in numbers) {
      freq[n] = (freq[n] ?? 0) + 1;
    }

    var maxCount = 1;
    for (final c in freq.values) {
      if (c > maxCount) maxCount = c;
    }
    if (maxCount <= 1) return const FormulaValue.error(FormulaError.na);

    final modes = freq.entries
        .where((e) => e.value == maxCount)
        .map((e) => e.key)
        .toList()
      ..sort();

    return FormulaValue.range([
      for (final m in modes) [FormulaValue.number(m)]
    ]);
  }
}

/// STDEVA(value1, [value2], ...) — Sample stdev including text/logical.
class StdevaFunction extends FormulaFunction {
  @override
  String get name => 'STDEVA';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = _collectAllValuesAdvanced(args, context);
    final n = numbers.length;
    if (n < 2) return const FormulaValue.error(FormulaError.divZero);
    final mean = numbers.fold(0.0, (a, b) => a + b) / n;
    var sum = 0.0;
    for (final x in numbers) {
      final d = x - mean;
      sum += d * d;
    }
    return FormulaValue.number(math.sqrt(sum / (n - 1)));
  }
}

/// STDEVPA(value1, [value2], ...) — Population stdev including text/logical.
class StdevpaFunction extends FormulaFunction {
  @override
  String get name => 'STDEVPA';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = _collectAllValuesAdvanced(args, context);
    final n = numbers.length;
    if (n < 1) return const FormulaValue.error(FormulaError.divZero);
    final mean = numbers.fold(0.0, (a, b) => a + b) / n;
    var sum = 0.0;
    for (final x in numbers) {
      final d = x - mean;
      sum += d * d;
    }
    return FormulaValue.number(math.sqrt(sum / n));
  }
}

/// VARA(value1, [value2], ...) — Sample variance including text/logical.
class VaraFunction extends FormulaFunction {
  @override
  String get name => 'VARA';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = _collectAllValuesAdvanced(args, context);
    final n = numbers.length;
    if (n < 2) return const FormulaValue.error(FormulaError.divZero);
    final mean = numbers.fold(0.0, (a, b) => a + b) / n;
    var sum = 0.0;
    for (final x in numbers) {
      final d = x - mean;
      sum += d * d;
    }
    return FormulaValue.number(sum / (n - 1));
  }
}

/// VARPA(value1, [value2], ...) — Population variance including text/logical.
class VarpaFunction extends FormulaFunction {
  @override
  String get name => 'VARPA';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final numbers = _collectAllValuesAdvanced(args, context);
    final n = numbers.length;
    if (n < 1) return const FormulaValue.error(FormulaError.divZero);
    final mean = numbers.fold(0.0, (a, b) => a + b) / n;
    var sum = 0.0;
    for (final x in numbers) {
      final d = x - mean;
      sum += d * d;
    }
    return FormulaValue.number(sum / n);
  }
}

// =============================================================================
// Wave 2 — 5 function classes
// =============================================================================

/// GAMMA(x) — Gamma function.
class GammaFunction extends FormulaFunction {
  @override
  String get name => 'GAMMA';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final x = v.toNumber()?.toDouble();
    if (x == null) return const FormulaValue.error(FormulaError.value);
    if (x <= 0 && x == x.truncateToDouble()) {
      return const FormulaValue.error(FormulaError.num);
    }
    return FormulaValue.number(_gamma(x));
  }
}

/// GAMMALN(x) — Natural log of Gamma function.
class GammaLnFunction extends FormulaFunction {
  @override
  String get name => 'GAMMALN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final x = v.toNumber()?.toDouble();
    if (x == null) return const FormulaValue.error(FormulaError.value);
    if (x <= 0) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(_lnGamma(x));
  }
}

/// GAMMALN.PRECISE(x) — Same as GAMMALN.
class GammaLnPreciseFunction extends FormulaFunction {
  @override
  String get name => 'GAMMALN.PRECISE';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final x = v.toNumber()?.toDouble();
    if (x == null) return const FormulaValue.error(FormulaError.value);
    if (x <= 0) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(_lnGamma(x));
  }
}

/// GAUSS(z) — Probability between mean and z standard deviations.
class GaussFunction extends FormulaFunction {
  @override
  String get name => 'GAUSS';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final z = v.toNumber()?.toDouble();
    if (z == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(0.5 * _erf(z / math.sqrt(2)));
  }
}

/// PHI(x) — Standard normal density function.
class PhiFunction extends FormulaFunction {
  @override
  String get name => 'PHI';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final x = v.toNumber()?.toDouble();
    if (x == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(_normSPdf(x));
  }
}

// =============================================================================
// Wave 3 — Normal Distribution (4 classes)
// =============================================================================

/// NORM.S.DIST(z, cumulative)
class NormSDistFunction extends FormulaFunction {
  @override
  String get name => 'NORM.S.DIST';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final z = vals[0].toNumber()?.toDouble();
    if (z == null) return const FormulaValue.error(FormulaError.value);
    final cum = vals[1].toBool();
    return FormulaValue.number(cum ? _normSCdf(z) : _normSPdf(z));
  }
}

/// NORM.S.INV(probability)
class NormSInvFunction extends FormulaFunction {
  @override
  String get name => 'NORM.S.INV';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final p = v.toNumber()?.toDouble();
    if (p == null) return const FormulaValue.error(FormulaError.value);
    if (p <= 0 || p >= 1) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(_normSInv(p));
  }
}

/// NORM.DIST(x, mean, standard_dev, cumulative)
class NormDistFunction extends FormulaFunction {
  @override
  String get name => 'NORM.DIST';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final x = vals[0].toNumber()?.toDouble();
    final mean = vals[1].toNumber()?.toDouble();
    final stddev = vals[2].toNumber()?.toDouble();
    if (x == null || mean == null || stddev == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (stddev <= 0) return const FormulaValue.error(FormulaError.num);
    final z = (x - mean) / stddev;
    final cum = vals[3].toBool();
    return FormulaValue.number(cum ? _normSCdf(z) : _normSPdf(z) / stddev);
  }
}

/// NORM.INV(probability, mean, standard_dev)
class NormInvFunction extends FormulaFunction {
  @override
  String get name => 'NORM.INV';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final p = vals[0].toNumber()?.toDouble();
    final mean = vals[1].toNumber()?.toDouble();
    final stddev = vals[2].toNumber()?.toDouble();
    if (p == null || mean == null || stddev == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (p <= 0 || p >= 1) return const FormulaValue.error(FormulaError.num);
    if (stddev <= 0) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(mean + stddev * _normSInv(p));
  }
}

// =============================================================================
// Wave 4 — Discrete Distributions (6 classes)
// =============================================================================

double _binomPmf(int k, int n, double p) {
  return _binomCoeff(n, k) * math.pow(p, k) * math.pow(1 - p, n - k);
}

double _binomCdf(int s, int n, double p) {
  var sum = 0.0;
  for (var k = 0; k <= s; k++) {
    sum += _binomPmf(k, n, p);
  }
  return sum;
}

/// BINOM.DIST(number_s, trials, probability_s, cumulative)
class BinomDistFunction extends FormulaFunction {
  @override
  String get name => 'BINOM.DIST';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final s = vals[0].toNumber()?.toInt();
    final n = vals[1].toNumber()?.toInt();
    final p = vals[2].toNumber()?.toDouble();
    if (s == null || n == null || p == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (s < 0 || n < 0 || s > n || p < 0 || p > 1) {
      return const FormulaValue.error(FormulaError.num);
    }
    final cum = vals[3].toBool();
    return FormulaValue.number(cum ? _binomCdf(s, n, p) : _binomPmf(s, n, p));
  }
}

/// BINOM.INV(trials, probability_s, alpha)
class BinomInvFunction extends FormulaFunction {
  @override
  String get name => 'BINOM.INV';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final n = vals[0].toNumber()?.toInt();
    final p = vals[1].toNumber()?.toDouble();
    final alpha = vals[2].toNumber()?.toDouble();
    if (n == null || p == null || alpha == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (n < 0 || p < 0 || p > 1 || alpha < 0 || alpha > 1) {
      return const FormulaValue.error(FormulaError.num);
    }
    var cumulative = 0.0;
    for (var s = 0; s <= n; s++) {
      cumulative += _binomPmf(s, n, p);
      if (cumulative >= alpha) return FormulaValue.number(s);
    }
    return FormulaValue.number(n);
  }
}

/// BINOM.DIST.RANGE(trials, probability_s, number_s, [number_s2])
class BinomDistRangeFunction extends FormulaFunction {
  @override
  String get name => 'BINOM.DIST.RANGE';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final n = vals[0].toNumber()?.toInt();
    final p = vals[1].toNumber()?.toDouble();
    final s1 = vals[2].toNumber()?.toInt();
    if (n == null || p == null || s1 == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    final s2 = args.length > 3 ? vals[3].toNumber()?.toInt() ?? s1 : s1;
    if (n < 0 || p < 0 || p > 1) {
      return const FormulaValue.error(FormulaError.num);
    }
    var sum = 0.0;
    for (var k = s1; k <= s2; k++) {
      sum += _binomPmf(k, n, p);
    }
    return FormulaValue.number(sum);
  }
}

/// NEGBINOM.DIST(number_f, number_s, probability_s, cumulative)
class NegBinomDistFunction extends FormulaFunction {
  @override
  String get name => 'NEGBINOM.DIST';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final f = vals[0].toNumber()?.toInt();
    final s = vals[1].toNumber()?.toInt();
    final p = vals[2].toNumber()?.toDouble();
    if (f == null || s == null || p == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (f < 0 || s < 1 || p <= 0 || p >= 1) {
      return const FormulaValue.error(FormulaError.num);
    }
    final cum = vals[3].toBool();
    double pmf(int k) =>
        _binomCoeff(k + s - 1, s - 1) *
        math.pow(p, s) *
        math.pow(1 - p, k);
    if (!cum) return FormulaValue.number(pmf(f));
    var sum = 0.0;
    for (var k = 0; k <= f; k++) {
      sum += pmf(k);
    }
    return FormulaValue.number(sum);
  }
}

/// HYPGEOM.DIST(sample_s, number_sample, population_s, number_pop, cumulative)
class HypgeomDistFunction extends FormulaFunction {
  @override
  String get name => 'HYPGEOM.DIST';
  @override
  int get minArgs => 5;
  @override
  int get maxArgs => 5;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final sampleS = vals[0].toNumber()?.toInt();
    final n = vals[1].toNumber()?.toInt();
    final popS = vals[2].toNumber()?.toInt();
    final popN = vals[3].toNumber()?.toInt();
    if (sampleS == null || n == null || popS == null || popN == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (sampleS < 0 ||
        n < 0 ||
        popS < 0 ||
        popN < 0 ||
        n > popN ||
        popS > popN ||
        sampleS > n ||
        sampleS > popS) {
      return const FormulaValue.error(FormulaError.num);
    }
    final cum = vals[4].toBool();
    double pmf(int k) =>
        _binomCoeff(popS, k) *
        _binomCoeff(popN - popS, n - k) /
        _binomCoeff(popN, n);
    if (!cum) return FormulaValue.number(pmf(sampleS));
    var sum = 0.0;
    for (var k = 0; k <= sampleS; k++) {
      sum += pmf(k);
    }
    return FormulaValue.number(sum);
  }
}

/// POISSON.DIST(x, mean, cumulative)
class PoissonDistFunction extends FormulaFunction {
  @override
  String get name => 'POISSON.DIST';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final x = vals[0].toNumber()?.toInt();
    final mean = vals[1].toNumber()?.toDouble();
    if (x == null || mean == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (x < 0 || mean < 0) return const FormulaValue.error(FormulaError.num);
    final cum = vals[2].toBool();
    double pmf(int k) =>
        math.exp(-mean) * math.pow(mean, k) / _gamma(k + 1);
    if (!cum) return FormulaValue.number(pmf(x));
    var sum = 0.0;
    for (var k = 0; k <= x; k++) {
      sum += pmf(k);
    }
    return FormulaValue.number(sum);
  }
}

// =============================================================================
// Wave 5 — Continuous Distributions (14 classes)
// =============================================================================

/// EXPON.DIST(x, lambda, cumulative)
class ExponDistFunction extends FormulaFunction {
  @override
  String get name => 'EXPON.DIST';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final x = vals[0].toNumber()?.toDouble();
    final lambda = vals[1].toNumber()?.toDouble();
    if (x == null || lambda == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (lambda <= 0) return const FormulaValue.error(FormulaError.num);
    if (x < 0) return const FormulaValue.error(FormulaError.num);
    final cum = vals[2].toBool();
    return FormulaValue.number(
        cum ? 1 - math.exp(-lambda * x) : lambda * math.exp(-lambda * x));
  }
}

/// GAMMA.DIST(x, alpha, beta, cumulative)
class GammaDistFunction extends FormulaFunction {
  @override
  String get name => 'GAMMA.DIST';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final x = vals[0].toNumber()?.toDouble();
    final alpha = vals[1].toNumber()?.toDouble();
    final beta = vals[2].toNumber()?.toDouble();
    if (x == null || alpha == null || beta == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (x < 0 || alpha <= 0 || beta <= 0) {
      return const FormulaValue.error(FormulaError.num);
    }
    final cum = vals[3].toBool();
    if (cum) return FormulaValue.number(_regularizedGammaP(alpha, x / beta));
    final pdf = math.pow(x, alpha - 1) *
        math.exp(-x / beta) /
        (math.pow(beta, alpha) * _gamma(alpha));
    return FormulaValue.number(pdf);
  }
}

/// GAMMA.INV(probability, alpha, beta)
class GammaInvFunction extends FormulaFunction {
  @override
  String get name => 'GAMMA.INV';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final p = vals[0].toNumber()?.toDouble();
    final alpha = vals[1].toNumber()?.toDouble();
    final beta = vals[2].toNumber()?.toDouble();
    if (p == null || alpha == null || beta == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (p <= 0 || p >= 1 || alpha <= 0 || beta <= 0) {
      return const FormulaValue.error(FormulaError.num);
    }
    final hi = math.max(1.0, alpha * beta) * 50;
    final result =
        _inverseCdf((x) => _regularizedGammaP(alpha, x / beta), p, 0, hi);
    return FormulaValue.number(result!);
  }
}

/// BETA.DIST(x, alpha, beta, cumulative, [A], [B])
class BetaDistFunction extends FormulaFunction {
  @override
  String get name => 'BETA.DIST';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 6;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final x = vals[0].toNumber()?.toDouble();
    final alpha = vals[1].toNumber()?.toDouble();
    final beta = vals[2].toNumber()?.toDouble();
    if (x == null || alpha == null || beta == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (alpha <= 0 || beta <= 0) {
      return const FormulaValue.error(FormulaError.num);
    }
    final cum = vals[3].toBool();
    final aLow = args.length > 4 ? vals[4].toNumber()?.toDouble() ?? 0.0 : 0.0;
    final bHigh =
        args.length > 5 ? vals[5].toNumber()?.toDouble() ?? 1.0 : 1.0;
    if (x < aLow || x > bHigh) {
      return const FormulaValue.error(FormulaError.num);
    }
    final z = (x - aLow) / (bHigh - aLow);
    if (cum) {
      if (z <= 0) return const FormulaValue.number(0);
      if (z >= 1) return const FormulaValue.number(1);
      return FormulaValue.number(_regularizedBeta(z, alpha, beta));
    }
    if (z <= 0 || z >= 1) return const FormulaValue.number(0);
    final pdf = math.pow(z, alpha - 1) *
        math.pow(1 - z, beta - 1) /
        _beta(alpha, beta) /
        (bHigh - aLow);
    return FormulaValue.number(pdf);
  }
}

/// BETA.INV(probability, alpha, beta, [A], [B])
class BetaInvFunction extends FormulaFunction {
  @override
  String get name => 'BETA.INV';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 5;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final p = vals[0].toNumber()?.toDouble();
    final alpha = vals[1].toNumber()?.toDouble();
    final beta = vals[2].toNumber()?.toDouble();
    if (p == null || alpha == null || beta == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (p <= 0 || p >= 1 || alpha <= 0 || beta <= 0) {
      return const FormulaValue.error(FormulaError.num);
    }
    final aLow = args.length > 3 ? vals[3].toNumber()?.toDouble() ?? 0.0 : 0.0;
    final bHigh =
        args.length > 4 ? vals[4].toNumber()?.toDouble() ?? 1.0 : 1.0;
    final z = _inverseCdf(
        (t) => _regularizedBeta(t, alpha, beta), p, 0, 1)!;
    return FormulaValue.number(aLow + z * (bHigh - aLow));
  }
}

/// CHISQ.DIST(x, degrees_freedom, cumulative)
class ChisqDistFunction extends FormulaFunction {
  @override
  String get name => 'CHISQ.DIST';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final x = vals[0].toNumber()?.toDouble();
    final df = vals[1].toNumber()?.toDouble();
    if (x == null || df == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (x < 0 || df < 1) return const FormulaValue.error(FormulaError.num);
    final cum = vals[2].toBool();
    if (cum) {
      return FormulaValue.number(_regularizedGammaP(df / 2, x / 2));
    }
    final pdf = math.pow(x, df / 2 - 1) *
        math.exp(-x / 2) /
        (math.pow(2, df / 2) * _gamma(df / 2));
    return FormulaValue.number(pdf);
  }
}

/// CHISQ.INV(probability, degrees_freedom)
class ChisqInvFunction extends FormulaFunction {
  @override
  String get name => 'CHISQ.INV';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final p = vals[0].toNumber()?.toDouble();
    final df = vals[1].toNumber()?.toDouble();
    if (p == null || df == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (p <= 0 || p >= 1 || df < 1) {
      return const FormulaValue.error(FormulaError.num);
    }
    final hi = math.max(1.0, df) * 10;
    return FormulaValue.number(
        _inverseCdf((x) => _regularizedGammaP(df / 2, x / 2), p, 0, hi)!);
  }
}

/// CHISQ.DIST.RT(x, degrees_freedom)
class ChisqDistRtFunction extends FormulaFunction {
  @override
  String get name => 'CHISQ.DIST.RT';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final x = vals[0].toNumber()?.toDouble();
    final df = vals[1].toNumber()?.toDouble();
    if (x == null || df == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (x < 0 || df < 1) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(1 - _regularizedGammaP(df / 2, x / 2));
  }
}

/// CHISQ.INV.RT(probability, degrees_freedom)
class ChisqInvRtFunction extends FormulaFunction {
  @override
  String get name => 'CHISQ.INV.RT';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final p = vals[0].toNumber()?.toDouble();
    final df = vals[1].toNumber()?.toDouble();
    if (p == null || df == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (p <= 0 || p >= 1 || df < 1) {
      return const FormulaValue.error(FormulaError.num);
    }
    final hi = math.max(1.0, df) * 10;
    return FormulaValue.number(
        _inverseCdf((x) => _regularizedGammaP(df / 2, x / 2), 1 - p, 0, hi)!);
  }
}

/// T.DIST(x, degrees_freedom, cumulative)
class TDistFunction extends FormulaFunction {
  @override
  String get name => 'T.DIST';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final x = vals[0].toNumber()?.toDouble();
    final df = vals[1].toNumber()?.toDouble();
    if (x == null || df == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (df < 1) return const FormulaValue.error(FormulaError.num);
    final cum = vals[2].toBool();
    if (cum) return FormulaValue.number(_tCdf(x, df));
    // PDF of t-distribution
    final coeff =
        _gamma((df + 1) / 2) / (math.sqrt(df * math.pi) * _gamma(df / 2));
    final pdf = coeff * math.pow(1 + x * x / df, -(df + 1) / 2);
    return FormulaValue.number(pdf);
  }
}

/// T.INV(probability, degrees_freedom)
class TInvFunction extends FormulaFunction {
  @override
  String get name => 'T.INV';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final p = vals[0].toNumber()?.toDouble();
    final df = vals[1].toNumber()?.toDouble();
    if (p == null || df == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (p <= 0 || p >= 1 || df < 1) {
      return const FormulaValue.error(FormulaError.num);
    }
    return FormulaValue.number(
        _inverseCdf((x) => _tCdf(x, df), p, -200, 200)!);
  }
}

/// T.DIST.2T(x, degrees_freedom)
class TDist2TFunction extends FormulaFunction {
  @override
  String get name => 'T.DIST.2T';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final x = vals[0].toNumber()?.toDouble();
    final df = vals[1].toNumber()?.toDouble();
    if (x == null || df == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (x < 0 || df < 1) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(2 * (1 - _tCdf(x, df)));
  }
}

/// T.INV.2T(probability, degrees_freedom)
class TInv2TFunction extends FormulaFunction {
  @override
  String get name => 'T.INV.2T';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final p = vals[0].toNumber()?.toDouble();
    final df = vals[1].toNumber()?.toDouble();
    if (p == null || df == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (p <= 0 || p > 1 || df < 1) {
      return const FormulaValue.error(FormulaError.num);
    }
    return FormulaValue.number(
        _inverseCdf((x) => _tCdf(x, df), 1 - p / 2, 0, 200)!);
  }
}

/// T.DIST.RT(x, degrees_freedom)
class TDistRtFunction extends FormulaFunction {
  @override
  String get name => 'T.DIST.RT';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final x = vals[0].toNumber()?.toDouble();
    final df = vals[1].toNumber()?.toDouble();
    if (x == null || df == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (df < 1) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(1 - _tCdf(x, df));
  }
}

// =============================================================================
// Wave 6 — F, Weibull, Lognormal (7 classes)
// =============================================================================

/// F.DIST(x, deg_freedom1, deg_freedom2, cumulative)
class FDistFunction extends FormulaFunction {
  @override
  String get name => 'F.DIST';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final x = vals[0].toNumber()?.toDouble();
    final d1 = vals[1].toNumber()?.toDouble();
    final d2 = vals[2].toNumber()?.toDouble();
    if (x == null || d1 == null || d2 == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (x < 0 || d1 < 1 || d2 < 1) {
      return const FormulaValue.error(FormulaError.num);
    }
    final cum = vals[3].toBool();
    if (cum) return FormulaValue.number(_fCdf(x, d1, d2));
    // F-distribution PDF
    final num1 = math.pow(d1 * x, d1) * math.pow(d2, d2);
    final den1 = math.pow(d1 * x + d2, d1 + d2);
    final pdf = math.sqrt(num1 / den1) / (x * _beta(d1 / 2, d2 / 2));
    return FormulaValue.number(pdf.isFinite ? pdf : 0);
  }
}

/// F.INV(probability, deg_freedom1, deg_freedom2)
class FInvFunction extends FormulaFunction {
  @override
  String get name => 'F.INV';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final p = vals[0].toNumber()?.toDouble();
    final d1 = vals[1].toNumber()?.toDouble();
    final d2 = vals[2].toNumber()?.toDouble();
    if (p == null || d1 == null || d2 == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (p <= 0 || p >= 1 || d1 < 1 || d2 < 1) {
      return const FormulaValue.error(FormulaError.num);
    }
    return FormulaValue.number(
        _inverseCdf((x) => _fCdf(x, d1, d2), p, 0, 1000)!);
  }
}

/// F.DIST.RT(x, deg_freedom1, deg_freedom2)
class FDistRtFunction extends FormulaFunction {
  @override
  String get name => 'F.DIST.RT';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final x = vals[0].toNumber()?.toDouble();
    final d1 = vals[1].toNumber()?.toDouble();
    final d2 = vals[2].toNumber()?.toDouble();
    if (x == null || d1 == null || d2 == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (x < 0 || d1 < 1 || d2 < 1) {
      return const FormulaValue.error(FormulaError.num);
    }
    return FormulaValue.number(1 - _fCdf(x, d1, d2));
  }
}

/// F.INV.RT(probability, deg_freedom1, deg_freedom2)
class FInvRtFunction extends FormulaFunction {
  @override
  String get name => 'F.INV.RT';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final p = vals[0].toNumber()?.toDouble();
    final d1 = vals[1].toNumber()?.toDouble();
    final d2 = vals[2].toNumber()?.toDouble();
    if (p == null || d1 == null || d2 == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (p <= 0 || p >= 1 || d1 < 1 || d2 < 1) {
      return const FormulaValue.error(FormulaError.num);
    }
    return FormulaValue.number(
        _inverseCdf((x) => _fCdf(x, d1, d2), 1 - p, 0, 1000)!);
  }
}

/// WEIBULL.DIST(x, alpha, beta, cumulative)
class WeibullDistFunction extends FormulaFunction {
  @override
  String get name => 'WEIBULL.DIST';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final x = vals[0].toNumber()?.toDouble();
    final alpha = vals[1].toNumber()?.toDouble();
    final beta = vals[2].toNumber()?.toDouble();
    if (x == null || alpha == null || beta == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (x < 0 || alpha <= 0 || beta <= 0) {
      return const FormulaValue.error(FormulaError.num);
    }
    final cum = vals[3].toBool();
    final xb = math.pow(x / beta, alpha);
    if (cum) return FormulaValue.number(1 - math.exp(-xb));
    final pdf = (alpha / beta) *
        math.pow(x / beta, alpha - 1) *
        math.exp(-xb);
    return FormulaValue.number(pdf);
  }
}

/// LOGNORM.DIST(x, mean, standard_dev, cumulative)
class LognormDistFunction extends FormulaFunction {
  @override
  String get name => 'LOGNORM.DIST';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final x = vals[0].toNumber()?.toDouble();
    final mean = vals[1].toNumber()?.toDouble();
    final stddev = vals[2].toNumber()?.toDouble();
    if (x == null || mean == null || stddev == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (x <= 0 || stddev <= 0) {
      return const FormulaValue.error(FormulaError.num);
    }
    final z = (math.log(x) - mean) / stddev;
    final cum = vals[3].toBool();
    return FormulaValue.number(
        cum ? _normSCdf(z) : _normSPdf(z) / (x * stddev));
  }
}

/// LOGNORM.INV(probability, mean, standard_dev)
class LognormInvFunction extends FormulaFunction {
  @override
  String get name => 'LOGNORM.INV';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final p = vals[0].toNumber()?.toDouble();
    final mean = vals[1].toNumber()?.toDouble();
    final stddev = vals[2].toNumber()?.toDouble();
    if (p == null || mean == null || stddev == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (p <= 0 || p >= 1 || stddev <= 0) {
      return const FormulaValue.error(FormulaError.num);
    }
    return FormulaValue.number(math.exp(mean + stddev * _normSInv(p)));
  }
}

// =============================================================================
// Wave 7 — Tests, Confidence, Regression (10 classes)
// =============================================================================

/// CONFIDENCE.NORM(alpha, standard_dev, size)
class ConfidenceNormFunction extends FormulaFunction {
  @override
  String get name => 'CONFIDENCE.NORM';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final alpha = vals[0].toNumber()?.toDouble();
    final stddev = vals[1].toNumber()?.toDouble();
    final size = vals[2].toNumber()?.toDouble();
    if (alpha == null || stddev == null || size == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (alpha <= 0 || alpha >= 1 || stddev <= 0 || size < 1) {
      return const FormulaValue.error(FormulaError.num);
    }
    return FormulaValue.number(
        _normSInv(1 - alpha / 2).abs() * stddev / math.sqrt(size));
  }
}

/// CONFIDENCE.T(alpha, standard_dev, size)
class ConfidenceTFunction extends FormulaFunction {
  @override
  String get name => 'CONFIDENCE.T';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final vals = evaluateArgs(args, context);
    for (final v in vals) {
      if (v.isError) return v;
    }
    final alpha = vals[0].toNumber()?.toDouble();
    final stddev = vals[1].toNumber()?.toDouble();
    final size = vals[2].toNumber()?.toInt();
    if (alpha == null || stddev == null || size == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (alpha <= 0 || alpha >= 1 || stddev <= 0 || size < 1) {
      return const FormulaValue.error(FormulaError.num);
    }
    final df = size - 1;
    if (df < 1) return const FormulaValue.error(FormulaError.divZero);
    final t = _inverseCdf(
        (x) => _tCdf(x, df.toDouble()), 1 - alpha / 2, 0, 200)!;
    return FormulaValue.number(t * stddev / math.sqrt(size));
  }
}

/// Z.TEST(array, x, [sigma])
class ZTestFunction extends FormulaFunction {
  @override
  String get name => 'Z.TEST';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final arrVal = args[0].evaluate(context);
    if (arrVal.isError) return arrVal;
    final numbers = <double>[];
    if (arrVal is RangeValue) {
      for (final cell in arrVal.flat) {
        if (cell is NumberValue) numbers.add(cell.value.toDouble());
      }
    } else {
      final n = arrVal.toNumber()?.toDouble();
      if (n != null) numbers.add(n);
    }
    if (numbers.isEmpty) return const FormulaValue.error(FormulaError.na);

    final xVal = args[1].evaluate(context);
    if (xVal.isError) return xVal;
    final x = xVal.toNumber()?.toDouble();
    if (x == null) return const FormulaValue.error(FormulaError.value);

    final n = numbers.length;
    final mean = numbers.fold(0.0, (a, b) => a + b) / n;

    double sigma;
    if (args.length > 2) {
      final sigVal = args[2].evaluate(context);
      if (sigVal.isError) return sigVal;
      final s = sigVal.toNumber()?.toDouble();
      if (s == null) return const FormulaValue.error(FormulaError.value);
      sigma = s;
    } else {
      // Sample standard deviation
      var ss = 0.0;
      for (final v in numbers) {
        final d = v - mean;
        ss += d * d;
      }
      sigma = math.sqrt(ss / (n - 1));
    }
    if (sigma == 0) return const FormulaValue.error(FormulaError.divZero);
    final z = (mean - x) / (sigma / math.sqrt(n));
    return FormulaValue.number(1 - _normSCdf(z));
  }
}

/// T.TEST(array1, array2, tails, type)
class TTestFunction extends FormulaFunction {
  @override
  String get name => 'T.TEST';
  @override
  int get minArgs => 4;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final tailsVal = args[2].evaluate(context);
    final typeVal = args[3].evaluate(context);
    if (tailsVal.isError) return tailsVal;
    if (typeVal.isError) return typeVal;
    final tails = tailsVal.toNumber()?.toInt();
    final type = typeVal.toNumber()?.toInt();
    if (tails == null || type == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (tails != 1 && tails != 2) {
      return const FormulaValue.error(FormulaError.num);
    }
    if (type < 1 || type > 3) {
      return const FormulaValue.error(FormulaError.num);
    }

    List<double> collect(FormulaNode arg) {
      final v = arg.evaluate(context);
      final result = <double>[];
      if (v is RangeValue) {
        for (final cell in v.flat) {
          if (cell is NumberValue) result.add(cell.value.toDouble());
        }
      } else {
        final n = v.toNumber()?.toDouble();
        if (n != null) result.add(n);
      }
      return result;
    }

    final a1 = collect(args[0]);
    final a2 = collect(args[1]);

    double t;
    double df;

    if (type == 1) {
      // Paired
      if (a1.length != a2.length || a1.length < 2) {
        return const FormulaValue.error(FormulaError.na);
      }
      final n = a1.length;
      final diffs = <double>[for (var i = 0; i < n; i++) a1[i] - a2[i]];
      final meanD = diffs.fold(0.0, (a, b) => a + b) / n;
      var ss = 0.0;
      for (final d in diffs) {
        ss += (d - meanD) * (d - meanD);
      }
      final sd = math.sqrt(ss / (n - 1));
      if (sd == 0) return const FormulaValue.error(FormulaError.divZero);
      t = meanD / (sd / math.sqrt(n));
      df = (n - 1).toDouble();
    } else if (type == 2) {
      // Two-sample equal variance
      final n1 = a1.length;
      final n2 = a2.length;
      if (n1 < 2 || n2 < 2) {
        return const FormulaValue.error(FormulaError.na);
      }
      final m1 = a1.fold(0.0, (a, b) => a + b) / n1;
      final m2 = a2.fold(0.0, (a, b) => a + b) / n2;
      var ss1 = 0.0, ss2 = 0.0;
      for (final v in a1) {
        ss1 += (v - m1) * (v - m1);
      }
      for (final v in a2) {
        ss2 += (v - m2) * (v - m2);
      }
      final sp2 = (ss1 + ss2) / (n1 + n2 - 2);
      if (sp2 == 0) return const FormulaValue.error(FormulaError.divZero);
      t = (m1 - m2) / math.sqrt(sp2 * (1 / n1 + 1 / n2));
      df = (n1 + n2 - 2).toDouble();
    } else {
      // Type 3 - Welch's t-test
      final n1 = a1.length;
      final n2 = a2.length;
      if (n1 < 2 || n2 < 2) {
        return const FormulaValue.error(FormulaError.na);
      }
      final m1 = a1.fold(0.0, (a, b) => a + b) / n1;
      final m2 = a2.fold(0.0, (a, b) => a + b) / n2;
      var ss1 = 0.0, ss2 = 0.0;
      for (final v in a1) {
        ss1 += (v - m1) * (v - m1);
      }
      for (final v in a2) {
        ss2 += (v - m2) * (v - m2);
      }
      final v1 = ss1 / (n1 - 1);
      final v2 = ss2 / (n2 - 1);
      final vn1 = v1 / n1;
      final vn2 = v2 / n2;
      if (vn1 + vn2 == 0) {
        return const FormulaValue.error(FormulaError.divZero);
      }
      t = (m1 - m2) / math.sqrt(vn1 + vn2);
      df = (vn1 + vn2) * (vn1 + vn2) /
          (vn1 * vn1 / (n1 - 1) + vn2 * vn2 / (n2 - 1));
    }

    final pVal = tails == 2
        ? 2 * (1 - _tCdf(t.abs(), df))
        : 1 - _tCdf(t.abs(), df);
    return FormulaValue.number(pVal);
  }
}

/// CHISQ.TEST(actual_range, expected_range)
class ChisqTestFunction extends FormulaFunction {
  @override
  String get name => 'CHISQ.TEST';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v1 = args[0].evaluate(context);
    final v2 = args[1].evaluate(context);
    if (v1.isError) return v1;
    if (v2.isError) return v2;
    if (v1 is! RangeValue || v2 is! RangeValue) {
      return const FormulaValue.error(FormulaError.value);
    }
    final rows = v1.rowCount;
    final cols = v1.columnCount;
    if (rows != v2.rowCount || cols != v2.columnCount) {
      return const FormulaValue.error(FormulaError.na);
    }
    var chi2 = 0.0;
    var count = 0;
    for (var r = 0; r < rows; r++) {
      for (var c = 0; c < cols; c++) {
        final a = v1.values[r][c].toNumber()?.toDouble();
        final e = v2.values[r][c].toNumber()?.toDouble();
        if (a == null || e == null) continue;
        if (e == 0) return const FormulaValue.error(FormulaError.divZero);
        chi2 += (a - e) * (a - e) / e;
        count++;
      }
    }
    final df = (rows > 1 && cols > 1)
        ? (rows - 1) * (cols - 1)
        : (count - 1);
    if (df < 1) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(
        1 - _regularizedGammaP(df / 2, chi2 / 2));
  }
}

/// F.TEST(array1, array2)
class FTestFunction extends FormulaFunction {
  @override
  String get name => 'F.TEST';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    List<double> collect(FormulaNode arg) {
      final v = arg.evaluate(context);
      final result = <double>[];
      if (v is RangeValue) {
        for (final cell in v.flat) {
          if (cell is NumberValue) result.add(cell.value.toDouble());
        }
      } else {
        final n = v.toNumber()?.toDouble();
        if (n != null) result.add(n);
      }
      return result;
    }

    final a1 = collect(args[0]);
    final a2 = collect(args[1]);
    final n1 = a1.length;
    final n2 = a2.length;
    if (n1 < 2 || n2 < 2) {
      return const FormulaValue.error(FormulaError.divZero);
    }
    final m1 = a1.fold(0.0, (a, b) => a + b) / n1;
    final m2 = a2.fold(0.0, (a, b) => a + b) / n2;
    var ss1 = 0.0, ss2 = 0.0;
    for (final v in a1) {
      ss1 += (v - m1) * (v - m1);
    }
    for (final v in a2) {
      ss2 += (v - m2) * (v - m2);
    }
    final var1 = ss1 / (n1 - 1);
    final var2 = ss2 / (n2 - 1);
    if (var1 == 0 && var2 == 0) {
      return const FormulaValue.error(FormulaError.divZero);
    }
    // Put larger variance on top
    final double f;
    final int df1, df2;
    if (var1 >= var2) {
      if (var2 == 0) return const FormulaValue.error(FormulaError.divZero);
      f = var1 / var2;
      df1 = n1 - 1;
      df2 = n2 - 1;
    } else {
      if (var1 == 0) return const FormulaValue.error(FormulaError.divZero);
      f = var2 / var1;
      df1 = n2 - 1;
      df2 = n1 - 1;
    }
    var pVal = 2 * (1 - _fCdf(f, df1.toDouble(), df2.toDouble()));
    if (pVal > 1) pVal = 1.0;
    return FormulaValue.number(pVal);
  }
}

/// LINEST(known_ys, [known_xs], [const], [stats])
class LinestFunction extends FormulaFunction {
  @override
  String get name => 'LINEST';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final ysVal = args[0].evaluate(context);
    if (ysVal.isError) return ysVal;
    final ys = <double>[];
    if (ysVal is RangeValue) {
      for (final cell in ysVal.flat) {
        if (cell is NumberValue) ys.add(cell.value.toDouble());
      }
    } else {
      final n = ysVal.toNumber()?.toDouble();
      if (n != null) ys.add(n);
    }
    if (ys.isEmpty) return const FormulaValue.error(FormulaError.value);

    List<double> xs;
    if (args.length > 1) {
      final xsVal = args[1].evaluate(context);
      if (xsVal.isError) return xsVal;
      xs = <double>[];
      if (xsVal is RangeValue) {
        for (final cell in xsVal.flat) {
          if (cell is NumberValue) xs.add(cell.value.toDouble());
        }
      } else {
        final n = xsVal.toNumber()?.toDouble();
        if (n != null) xs.add(n);
      }
    } else {
      xs = [for (var i = 1; i <= ys.length; i++) i.toDouble()];
    }
    if (xs.length != ys.length) {
      return const FormulaValue.error(FormulaError.na);
    }

    final n = xs.length;
    final mx = xs.fold(0.0, (a, b) => a + b) / n;
    final my = ys.fold(0.0, (a, b) => a + b) / n;
    var sumXY = 0.0, sumX2 = 0.0;
    for (var i = 0; i < n; i++) {
      sumXY += (xs[i] - mx) * (ys[i] - my);
      sumX2 += (xs[i] - mx) * (xs[i] - mx);
    }
    if (sumX2 == 0) return const FormulaValue.error(FormulaError.divZero);
    final slope = sumXY / sumX2;
    final intercept = my - slope * mx;

    final showStats = args.length > 3 && args[3].evaluate(context).toBool();
    if (!showStats) {
      return FormulaValue.range([
        [FormulaValue.number(slope), FormulaValue.number(intercept)]
      ]);
    }

    // Compute regression statistics
    var sse = 0.0, sst = 0.0;
    for (var i = 0; i < n; i++) {
      final yhat = intercept + slope * xs[i];
      sse += (ys[i] - yhat) * (ys[i] - yhat);
      sst += (ys[i] - my) * (ys[i] - my);
    }
    final r2 = sst == 0 ? 1.0 : 1 - sse / sst;
    final seY = n > 2 ? math.sqrt(sse / (n - 2)) : 0.0;
    final seSlope = sumX2 > 0 && n > 2 ? seY / math.sqrt(sumX2) : 0.0;
    final seIntercept = n > 2
        ? seY * math.sqrt(xs.fold(0.0, (a, b) => a + b * b) / (n * sumX2))
        : 0.0;
    final fStat = (n > 2 && sse > 0) ? (sst - sse) / (sse / (n - 2)) : 0.0;
    final dfVal = (n - 2).toDouble();

    return FormulaValue.range([
      [FormulaValue.number(slope), FormulaValue.number(intercept)],
      [FormulaValue.number(seSlope), FormulaValue.number(seIntercept)],
      [FormulaValue.number(r2), FormulaValue.number(seY)],
      [FormulaValue.number(fStat), FormulaValue.number(dfVal)],
      [
        FormulaValue.number(sst - sse),
        FormulaValue.number(sse),
      ],
    ]);
  }
}

/// LOGEST(known_ys, [known_xs], [const], [stats])
class LogestFunction extends FormulaFunction {
  @override
  String get name => 'LOGEST';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final ysVal = args[0].evaluate(context);
    if (ysVal.isError) return ysVal;
    final ys = <double>[];
    if (ysVal is RangeValue) {
      for (final cell in ysVal.flat) {
        if (cell is NumberValue) ys.add(cell.value.toDouble());
      }
    } else {
      final n = ysVal.toNumber()?.toDouble();
      if (n != null) ys.add(n);
    }
    if (ys.isEmpty) return const FormulaValue.error(FormulaError.value);
    // All ys must be positive for log
    for (final y in ys) {
      if (y <= 0) return const FormulaValue.error(FormulaError.num);
    }
    final logYs = [for (final y in ys) math.log(y)];

    List<double> xs;
    if (args.length > 1) {
      final xsVal = args[1].evaluate(context);
      if (xsVal.isError) return xsVal;
      xs = <double>[];
      if (xsVal is RangeValue) {
        for (final cell in xsVal.flat) {
          if (cell is NumberValue) xs.add(cell.value.toDouble());
        }
      } else {
        final n = xsVal.toNumber()?.toDouble();
        if (n != null) xs.add(n);
      }
    } else {
      xs = [for (var i = 1; i <= ys.length; i++) i.toDouble()];
    }
    if (xs.length != logYs.length) {
      return const FormulaValue.error(FormulaError.na);
    }

    final n = xs.length;
    final mx = xs.fold(0.0, (a, b) => a + b) / n;
    final my = logYs.fold(0.0, (a, b) => a + b) / n;
    var sumXY = 0.0, sumX2 = 0.0;
    for (var i = 0; i < n; i++) {
      sumXY += (xs[i] - mx) * (logYs[i] - my);
      sumX2 += (xs[i] - mx) * (xs[i] - mx);
    }
    if (sumX2 == 0) return const FormulaValue.error(FormulaError.divZero);
    final slope = sumXY / sumX2;
    final intercept = my - slope * mx;

    final m = math.exp(slope);
    final b = math.exp(intercept);

    final showStats = args.length > 3 && args[3].evaluate(context).toBool();
    if (!showStats) {
      return FormulaValue.range([
        [FormulaValue.number(m), FormulaValue.number(b)]
      ]);
    }

    var sse = 0.0, sst = 0.0;
    for (var i = 0; i < n; i++) {
      final yhat = intercept + slope * xs[i];
      sse += (logYs[i] - yhat) * (logYs[i] - yhat);
      sst += (logYs[i] - my) * (logYs[i] - my);
    }
    final r2 = sst == 0 ? 1.0 : 1 - sse / sst;
    final seY = n > 2 ? math.sqrt(sse / (n - 2)) : 0.0;
    final seSlope = sumX2 > 0 && n > 2 ? seY / math.sqrt(sumX2) : 0.0;
    final seIntercept = n > 2
        ? seY * math.sqrt(xs.fold(0.0, (a, b) => a + b * b) / (n * sumX2))
        : 0.0;
    final fStat = (n > 2 && sse > 0) ? (sst - sse) / (sse / (n - 2)) : 0.0;
    final dfVal = (n - 2).toDouble();

    return FormulaValue.range([
      [FormulaValue.number(m), FormulaValue.number(b)],
      [FormulaValue.number(seSlope), FormulaValue.number(seIntercept)],
      [FormulaValue.number(r2), FormulaValue.number(seY)],
      [FormulaValue.number(fStat), FormulaValue.number(dfVal)],
      [FormulaValue.number(sst - sse), FormulaValue.number(sse)],
    ]);
  }
}

/// TREND(known_ys, [known_xs], [new_xs], [const])
class TrendFunction extends FormulaFunction {
  @override
  String get name => 'TREND';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final ysVal = args[0].evaluate(context);
    if (ysVal.isError) return ysVal;
    final ys = <double>[];
    if (ysVal is RangeValue) {
      for (final cell in ysVal.flat) {
        if (cell is NumberValue) ys.add(cell.value.toDouble());
      }
    } else {
      final n = ysVal.toNumber()?.toDouble();
      if (n != null) ys.add(n);
    }
    if (ys.isEmpty) return const FormulaValue.error(FormulaError.value);

    List<double> xs;
    if (args.length > 1) {
      final xsVal = args[1].evaluate(context);
      if (xsVal.isError) return xsVal;
      xs = <double>[];
      if (xsVal is RangeValue) {
        for (final cell in xsVal.flat) {
          if (cell is NumberValue) xs.add(cell.value.toDouble());
        }
      } else {
        final n = xsVal.toNumber()?.toDouble();
        if (n != null) xs.add(n);
      }
    } else {
      xs = [for (var i = 1; i <= ys.length; i++) i.toDouble()];
    }
    if (xs.length != ys.length) {
      return const FormulaValue.error(FormulaError.na);
    }

    final n = xs.length;
    final mx = xs.fold(0.0, (a, b) => a + b) / n;
    final my = ys.fold(0.0, (a, b) => a + b) / n;
    var sumXY = 0.0, sumX2 = 0.0;
    for (var i = 0; i < n; i++) {
      sumXY += (xs[i] - mx) * (ys[i] - my);
      sumX2 += (xs[i] - mx) * (xs[i] - mx);
    }
    if (sumX2 == 0) return const FormulaValue.error(FormulaError.divZero);
    final slope = sumXY / sumX2;
    final intercept = my - slope * mx;

    List<double> newXs;
    if (args.length > 2) {
      final nxVal = args[2].evaluate(context);
      if (nxVal.isError) return nxVal;
      newXs = <double>[];
      if (nxVal is RangeValue) {
        for (final cell in nxVal.flat) {
          if (cell is NumberValue) newXs.add(cell.value.toDouble());
        }
      } else {
        final v = nxVal.toNumber()?.toDouble();
        if (v != null) newXs.add(v);
      }
    } else {
      newXs = xs;
    }

    return FormulaValue.range([
      for (final nx in newXs)
        [FormulaValue.number(intercept + slope * nx)]
    ]);
  }
}

/// GROWTH(known_ys, [known_xs], [new_xs], [const])
class GrowthFunction extends FormulaFunction {
  @override
  String get name => 'GROWTH';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 4;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final ysVal = args[0].evaluate(context);
    if (ysVal.isError) return ysVal;
    final ys = <double>[];
    if (ysVal is RangeValue) {
      for (final cell in ysVal.flat) {
        if (cell is NumberValue) ys.add(cell.value.toDouble());
      }
    } else {
      final n = ysVal.toNumber()?.toDouble();
      if (n != null) ys.add(n);
    }
    if (ys.isEmpty) return const FormulaValue.error(FormulaError.value);
    for (final y in ys) {
      if (y <= 0) return const FormulaValue.error(FormulaError.num);
    }
    final logYs = [for (final y in ys) math.log(y)];

    List<double> xs;
    if (args.length > 1) {
      final xsVal = args[1].evaluate(context);
      if (xsVal.isError) return xsVal;
      xs = <double>[];
      if (xsVal is RangeValue) {
        for (final cell in xsVal.flat) {
          if (cell is NumberValue) xs.add(cell.value.toDouble());
        }
      } else {
        final n = xsVal.toNumber()?.toDouble();
        if (n != null) xs.add(n);
      }
    } else {
      xs = [for (var i = 1; i <= ys.length; i++) i.toDouble()];
    }
    if (xs.length != logYs.length) {
      return const FormulaValue.error(FormulaError.na);
    }

    final n = xs.length;
    final mx = xs.fold(0.0, (a, b) => a + b) / n;
    final my = logYs.fold(0.0, (a, b) => a + b) / n;
    var sumXY = 0.0, sumX2 = 0.0;
    for (var i = 0; i < n; i++) {
      sumXY += (xs[i] - mx) * (logYs[i] - my);
      sumX2 += (xs[i] - mx) * (xs[i] - mx);
    }
    if (sumX2 == 0) return const FormulaValue.error(FormulaError.divZero);
    final slope = sumXY / sumX2;
    final intercept = my - slope * mx;

    List<double> newXs;
    if (args.length > 2) {
      final nxVal = args[2].evaluate(context);
      if (nxVal.isError) return nxVal;
      newXs = <double>[];
      if (nxVal is RangeValue) {
        for (final cell in nxVal.flat) {
          if (cell is NumberValue) newXs.add(cell.value.toDouble());
        }
      } else {
        final v = nxVal.toNumber()?.toDouble();
        if (v != null) newXs.add(v);
      }
    } else {
      newXs = xs;
    }

    return FormulaValue.range([
      for (final nx in newXs)
        [FormulaValue.number(math.exp(intercept + slope * nx))]
    ]);
  }
}
