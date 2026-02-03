import 'dart:math' as math;

import '../ast/nodes.dart';
import '../evaluation/context.dart';
import '../evaluation/errors.dart';
import '../evaluation/value.dart';
import 'function.dart';
import 'registry.dart';

// =============================================================================
// Shared Helpers
// =============================================================================

/// Abramowitz & Stegun approximation of the error function.
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

// --------------- Two's Complement Base Conversion Helpers -------------------

const int _binBits = 10; // 10-bit two's complement
const int _octBits = 30; // 30-bit two's complement
const int _hexBits = 40; // 40-bit two's complement

int? _parseBinary(String s) {
  if (s.length > 10) return null;
  for (final c in s.codeUnits) {
    if (c != 0x30 && c != 0x31) return null; // only 0 and 1
  }
  final raw = int.tryParse(s, radix: 2);
  if (raw == null) return null;
  // If 10 digits and leading bit is 1, it's negative (two's complement)
  if (s.length == 10 && s[0] == '1') {
    return raw - (1 << _binBits);
  }
  return raw;
}

int? _parseOctal(String s) {
  if (s.length > 10) return null;
  for (final c in s.codeUnits) {
    if (c < 0x30 || c > 0x37) return null; // only 0-7
  }
  final raw = int.tryParse(s, radix: 8);
  if (raw == null) return null;
  if (s.length == 10 && s.codeUnitAt(0) >= 0x34) {
    // leading digit >= 4 means negative in 30-bit two's complement
    return raw - (1 << _octBits);
  }
  return raw;
}

int? _parseHex(String s) {
  if (s.length > 10) return null;
  final upper = s.toUpperCase();
  for (final c in upper.codeUnits) {
    if (!((c >= 0x30 && c <= 0x39) || (c >= 0x41 && c <= 0x46))) return null;
  }
  final raw = int.tryParse(upper, radix: 16);
  if (raw == null) return null;
  if (s.length == 10 && upper.codeUnitAt(0) >= 0x38) {
    // leading digit >= 8 means negative in 40-bit two's complement
    return raw - (1 << _hexBits);
  }
  return raw;
}

String? _toBinary(int value, [int? places]) {
  if (value < -(1 << (_binBits - 1)) || value > (1 << (_binBits - 1)) - 1) {
    return null; // -512..511
  }
  String result;
  if (value < 0) {
    result = ((1 << _binBits) + value).toRadixString(2);
  } else {
    result = value.toRadixString(2);
  }
  if (places != null) {
    if (places < 1 || places > 10) return null;
    if (result.length > places) return null;
    result = result.padLeft(places, '0');
  }
  return result;
}

String? _toOctal(int value, [int? places]) {
  if (value < -(1 << (_octBits - 1)) || value > (1 << (_octBits - 1)) - 1) {
    return null;
  }
  String result;
  if (value < 0) {
    result = ((1 << _octBits) + value).toRadixString(8);
  } else {
    result = value.toRadixString(8);
  }
  if (places != null) {
    if (places < 1 || places > 10) return null;
    if (result.length > places) return null;
    result = result.padLeft(places, '0');
  }
  return result;
}

String? _toHex(int value, [int? places]) {
  if (value < -(1 << (_hexBits - 1)) || value > (1 << (_hexBits - 1)) - 1) {
    return null;
  }
  String result;
  if (value < 0) {
    result = ((1 << _hexBits) + value).toRadixString(16).toUpperCase();
  } else {
    result = value.toRadixString(16).toUpperCase();
  }
  if (places != null) {
    if (places < 1 || places > 10) return null;
    if (result.length > places) return null;
    result = result.padLeft(places, '0');
  }
  return result;
}

// --------------- Complex Number Helpers -----------------------------------

/// Parse a complex number string like "3+4i", "3-4i", "5", "4i", "i", "-i",
/// "3+4j", etc. Returns (real, imag, suffix) or null on failure.
({double real, double imag, String suffix})? _parseComplex(String text) {
  text = text.trim();
  if (text.isEmpty) return null;

  // Determine suffix
  String suffix = 'i';
  if (text.endsWith('j')) {
    suffix = 'j';
    text = text.substring(0, text.length - 1);
  } else if (text.endsWith('i')) {
    suffix = 'i';
    text = text.substring(0, text.length - 1);
  } else {
    // Pure real number
    final val = double.tryParse(text);
    if (val == null) return null;
    return (real: val, imag: 0.0, suffix: 'i');
  }

  // Now text has suffix stripped. It could be:
  // "" -> means just "i" or "j" -> 0+1i
  // "-" -> means "-i" or "-j" -> 0-1i
  // "3+4" -> real=3, imag=4
  // "3-4" -> real=3, imag=-4
  // "4" -> imag=4
  // "+4" -> imag=4
  // "-4" -> imag=-4
  // "3+" -> real=3, imag=1
  // "3-" -> real=3, imag=-1

  if (text.isEmpty) {
    return (real: 0.0, imag: 1.0, suffix: suffix);
  }
  if (text == '+') {
    return (real: 0.0, imag: 1.0, suffix: suffix);
  }
  if (text == '-') {
    return (real: 0.0, imag: -1.0, suffix: suffix);
  }

  // Find the last + or - that splits real and imaginary parts.
  // Skip the first character (it could be a sign for the real part).
  int splitPos = -1;
  for (int i = text.length - 1; i >= 1; i--) {
    if (text[i] == '+' || text[i] == '-') {
      // Make sure this is not part of an exponent (e.g., "1e+2")
      if (i > 0 && (text[i - 1] == 'e' || text[i - 1] == 'E')) continue;
      splitPos = i;
      break;
    }
  }

  if (splitPos == -1) {
    // No split found => pure imaginary
    final imag = double.tryParse(text);
    if (imag == null) return null;
    return (real: 0.0, imag: imag, suffix: suffix);
  }

  final realPart = text.substring(0, splitPos);
  final imagPart = text.substring(splitPos);

  final realVal = double.tryParse(realPart);
  if (realVal == null) return null;

  double imagVal;
  if (imagPart == '+' || imagPart == '') {
    imagVal = 1.0;
  } else if (imagPart == '-') {
    imagVal = -1.0;
  } else {
    final parsed = double.tryParse(imagPart);
    if (parsed == null) return null;
    imagVal = parsed;
  }

  return (real: realVal, imag: imagVal, suffix: suffix);
}

/// Format a complex number to text. Rounds near-zero values.
String _formatComplex(double real, double imag, [String suffix = 'i']) {
  // Round values very close to integer
  double roundNearZero(double v) {
    if ((v - v.roundToDouble()).abs() < 1e-12) return v.roundToDouble();
    return v;
  }

  real = roundNearZero(real);
  imag = roundNearZero(imag);

  String formatNum(double v) {
    if (v == v.truncateToDouble()) return v.toInt().toString();
    return v.toString();
  }

  if (imag == 0) {
    return formatNum(real);
  }
  if (real == 0) {
    if (imag == 1) return suffix;
    if (imag == -1) return '-$suffix';
    return '${formatNum(imag)}$suffix';
  }
  // Both parts
  if (imag == 1) return '${formatNum(real)}+$suffix';
  if (imag == -1) return '${formatNum(real)}-$suffix';
  if (imag > 0) return '${formatNum(real)}+${formatNum(imag)}$suffix';
  return '${formatNum(real)}${formatNum(imag)}$suffix'; // imag already has minus sign
}

// --------------- CONVERT Unit Table -----------------------------------------

enum _UnitCategory {
  weight,
  distance,
  time,
  pressure,
  force,
  energy,
  power,
  temperature,
  volume,
  area,
  information,
  speed,
  magnetism,
}

class _UnitDef {
  final _UnitCategory category;
  final double toBase; // multiply by this to convert to base unit
  final bool allowMetricPrefix;
  final bool allowBinaryPrefix;

  const _UnitDef(this.category, this.toBase,
      {this.allowMetricPrefix = false, this.allowBinaryPrefix = false});
}

// Base units: kg, m, sec, Pa, N, J, W, K, m^3, m^2, bit, m/s, T
final Map<String, _UnitDef> _units = {
  // Weight (base: kg)
  'g': const _UnitDef(_UnitCategory.weight, 0.001, allowMetricPrefix: true),
  'kg': const _UnitDef(_UnitCategory.weight, 1),
  'mg': const _UnitDef(_UnitCategory.weight, 0.000001),
  'lbm': const _UnitDef(_UnitCategory.weight, 0.45359237),
  'ozm': const _UnitDef(_UnitCategory.weight, 0.028349523125),
  'stone': const _UnitDef(_UnitCategory.weight, 6.35029318),
  'ton': const _UnitDef(_UnitCategory.weight, 907.18474),
  'sg': const _UnitDef(_UnitCategory.weight, 14.593903),
  'u': const _UnitDef(_UnitCategory.weight, 1.6605390666e-27),
  'grain': const _UnitDef(_UnitCategory.weight, 0.00006479891),
  'cwt': const _UnitDef(_UnitCategory.weight, 45.359237),
  'shweight': const _UnitDef(_UnitCategory.weight, 45.359237),
  'uk_cwt': const _UnitDef(_UnitCategory.weight, 50.80234544),
  'lcwt': const _UnitDef(_UnitCategory.weight, 50.80234544),
  'hweight': const _UnitDef(_UnitCategory.weight, 50.80234544),
  'brton': const _UnitDef(_UnitCategory.weight, 1016.0469088),
  'LTON': const _UnitDef(_UnitCategory.weight, 1016.0469088),
  'uk_ton': const _UnitDef(_UnitCategory.weight, 1016.0469088),

  // Distance (base: m)
  'm': const _UnitDef(_UnitCategory.distance, 1, allowMetricPrefix: true),
  'km': const _UnitDef(_UnitCategory.distance, 1000),
  'cm': const _UnitDef(_UnitCategory.distance, 0.01),
  'mm': const _UnitDef(_UnitCategory.distance, 0.001),
  'mi': const _UnitDef(_UnitCategory.distance, 1609.344),
  'Nmi': const _UnitDef(_UnitCategory.distance, 1852),
  'in': const _UnitDef(_UnitCategory.distance, 0.0254),
  'ft': const _UnitDef(_UnitCategory.distance, 0.3048),
  'yd': const _UnitDef(_UnitCategory.distance, 0.9144),
  'ang': const _UnitDef(_UnitCategory.distance, 1e-10),
  'ell': const _UnitDef(_UnitCategory.distance, 1.143),
  'ly': const _UnitDef(_UnitCategory.distance, 9.46073047258e15),
  'parsec': const _UnitDef(_UnitCategory.distance, 3.0856775814914e16),
  'pc': const _UnitDef(_UnitCategory.distance, 3.0856775814914e16),
  'Pica': const _UnitDef(_UnitCategory.distance, 0.0254 / 6),
  'Picapt': const _UnitDef(_UnitCategory.distance, 0.0254 / 72),
  'survey_mi': const _UnitDef(_UnitCategory.distance, 1609.3472186944),

  // Time (base: sec)
  'sec': const _UnitDef(_UnitCategory.time, 1, allowMetricPrefix: true),
  's': const _UnitDef(_UnitCategory.time, 1, allowMetricPrefix: true),
  'mn': const _UnitDef(_UnitCategory.time, 60),
  'min': const _UnitDef(_UnitCategory.time, 60),
  'hr': const _UnitDef(_UnitCategory.time, 3600),
  'day': const _UnitDef(_UnitCategory.time, 86400),
  'yr': const _UnitDef(_UnitCategory.time, 365.25 * 86400),

  // Pressure (base: Pa)
  'Pa': const _UnitDef(_UnitCategory.pressure, 1, allowMetricPrefix: true),
  'p': const _UnitDef(_UnitCategory.pressure, 1, allowMetricPrefix: true),
  'atm': const _UnitDef(_UnitCategory.pressure, 101325),
  'at': const _UnitDef(_UnitCategory.pressure, 98066.5),
  'mmHg': const _UnitDef(_UnitCategory.pressure, 133.322387415),
  'psi': const _UnitDef(_UnitCategory.pressure, 6894.757293168),
  'Torr': const _UnitDef(_UnitCategory.pressure, 133.3223684211),

  // Force (base: N)
  'N': const _UnitDef(_UnitCategory.force, 1, allowMetricPrefix: true),
  'dyn': const _UnitDef(_UnitCategory.force, 0.00001, allowMetricPrefix: true),
  'dy': const _UnitDef(_UnitCategory.force, 0.00001),
  'lbf': const _UnitDef(_UnitCategory.force, 4.4482216152605),
  'pond': const _UnitDef(_UnitCategory.force, 0.00980665, allowMetricPrefix: true),

  // Energy (base: J)
  'J': const _UnitDef(_UnitCategory.energy, 1, allowMetricPrefix: true),
  'e': const _UnitDef(_UnitCategory.energy, 1e-7, allowMetricPrefix: true),
  'eV': const _UnitDef(_UnitCategory.energy, 1.602176634e-19, allowMetricPrefix: true),
  'cal': const _UnitDef(_UnitCategory.energy, 4.1868, allowMetricPrefix: true),
  'c': const _UnitDef(_UnitCategory.energy, 4.1868, allowMetricPrefix: true),
  'Cal': const _UnitDef(_UnitCategory.energy, 4186.8),
  'BTU': const _UnitDef(_UnitCategory.energy, 1055.05585262),
  'btu': const _UnitDef(_UnitCategory.energy, 1055.05585262),
  'Wh': const _UnitDef(_UnitCategory.energy, 3600, allowMetricPrefix: true),
  'wh': const _UnitDef(_UnitCategory.energy, 3600, allowMetricPrefix: true),
  'HPh': const _UnitDef(_UnitCategory.energy, 2684519.5368856),
  'hph': const _UnitDef(_UnitCategory.energy, 2684519.5368856),
  'flb': const _UnitDef(_UnitCategory.energy, 1.3558179483314),

  // Power (base: W)
  'W': const _UnitDef(_UnitCategory.power, 1, allowMetricPrefix: true),
  'w': const _UnitDef(_UnitCategory.power, 1, allowMetricPrefix: true),
  'HP': const _UnitDef(_UnitCategory.power, 745.69987158227),
  'h': const _UnitDef(_UnitCategory.power, 745.69987158227),
  'PS': const _UnitDef(_UnitCategory.power, 735.49875),

  // Temperature (base: K) — special-cased in conversion
  'C': const _UnitDef(_UnitCategory.temperature, 1),
  'cel': const _UnitDef(_UnitCategory.temperature, 1),
  'F': const _UnitDef(_UnitCategory.temperature, 1),
  'fah': const _UnitDef(_UnitCategory.temperature, 1),
  'K': const _UnitDef(_UnitCategory.temperature, 1),
  'kel': const _UnitDef(_UnitCategory.temperature, 1),
  'Rank': const _UnitDef(_UnitCategory.temperature, 1),
  'Reau': const _UnitDef(_UnitCategory.temperature, 1),

  // Volume (base: m^3)
  'l': const _UnitDef(_UnitCategory.volume, 0.001, allowMetricPrefix: true),
  'L': const _UnitDef(_UnitCategory.volume, 0.001, allowMetricPrefix: true),
  'lt': const _UnitDef(_UnitCategory.volume, 0.001, allowMetricPrefix: true),
  'tsp': const _UnitDef(_UnitCategory.volume, 4.92892159375e-6),
  'tspm': const _UnitDef(_UnitCategory.volume, 5e-6),
  'tbs': const _UnitDef(_UnitCategory.volume, 1.478676478125e-5),
  'oz': const _UnitDef(_UnitCategory.volume, 2.95735295625e-5),
  'cup': const _UnitDef(_UnitCategory.volume, 2.365882365e-4),
  'pt': const _UnitDef(_UnitCategory.volume, 4.73176473e-4),
  'us_pt': const _UnitDef(_UnitCategory.volume, 4.73176473e-4),
  'uk_pt': const _UnitDef(_UnitCategory.volume, 5.6826125e-4),
  'qt': const _UnitDef(_UnitCategory.volume, 9.46352946e-4),
  'uk_qt': const _UnitDef(_UnitCategory.volume, 1.1365225e-3),
  'gal': const _UnitDef(_UnitCategory.volume, 3.785411784e-3),
  'uk_gal': const _UnitDef(_UnitCategory.volume, 4.54609e-3),
  'ang3': const _UnitDef(_UnitCategory.volume, 1e-30),
  'barrel': const _UnitDef(_UnitCategory.volume, 0.158987294928),
  'bushel': const _UnitDef(_UnitCategory.volume, 0.03523907016688),
  'in3': const _UnitDef(_UnitCategory.volume, 1.6387064e-5),
  'ft3': const _UnitDef(_UnitCategory.volume, 0.028316846592),
  'ly3': const _UnitDef(_UnitCategory.volume, 8.46786664623715e47),
  'm3': const _UnitDef(_UnitCategory.volume, 1),
  'mi3': const _UnitDef(_UnitCategory.volume, 4168181825.44058),
  'yd3': const _UnitDef(_UnitCategory.volume, 0.764554857984),
  'Nmi3': const _UnitDef(_UnitCategory.volume, 6352182208),
  'Pica3': const _UnitDef(_UnitCategory.volume, 7.58857917e-8),
  'GRT': const _UnitDef(_UnitCategory.volume, 2.8316846592),
  'regton': const _UnitDef(_UnitCategory.volume, 2.8316846592),
  'MTON': const _UnitDef(_UnitCategory.volume, 1.13267386368),

  // Area (base: m^2)
  'ha': const _UnitDef(_UnitCategory.area, 10000),
  'uk_acre': const _UnitDef(_UnitCategory.area, 4046.8564224),
  'us_acre': const _UnitDef(_UnitCategory.area, 4046.8564224),
  'ang2': const _UnitDef(_UnitCategory.area, 1e-20),
  'ar': const _UnitDef(_UnitCategory.area, 100, allowMetricPrefix: true),
  'ft2': const _UnitDef(_UnitCategory.area, 0.09290304),
  'in2': const _UnitDef(_UnitCategory.area, 0.00064516),
  'ly2': const _UnitDef(_UnitCategory.area, 8.9502087256e31),
  'm2': const _UnitDef(_UnitCategory.area, 1),
  'Morgen': const _UnitDef(_UnitCategory.area, 2500),
  'mi2': const _UnitDef(_UnitCategory.area, 2589988.110336),
  'Nmi2': const _UnitDef(_UnitCategory.area, 3429904),
  'Pica2': const _UnitDef(_UnitCategory.area, 1.79211111e-5),
  'yd2': const _UnitDef(_UnitCategory.area, 0.83612736),

  // Information (base: bit)
  'bit': const _UnitDef(_UnitCategory.information, 1,
      allowMetricPrefix: true, allowBinaryPrefix: true),
  'byte': const _UnitDef(_UnitCategory.information, 8,
      allowMetricPrefix: true, allowBinaryPrefix: true),

  // Speed (base: m/s)
  'm/s': const _UnitDef(_UnitCategory.speed, 1),
  'm/h': const _UnitDef(_UnitCategory.speed, 1 / 3600),
  'mph': const _UnitDef(_UnitCategory.speed, 0.44704),
  'kn': const _UnitDef(_UnitCategory.speed, 0.514444444444),
  'admkn': const _UnitDef(_UnitCategory.speed, 0.514773333333),
  'km/h': const _UnitDef(_UnitCategory.speed, 1 / 3.6),
  'km/s': const _UnitDef(_UnitCategory.speed, 1000),

  // Magnetism (base: T)
  'T': const _UnitDef(_UnitCategory.magnetism, 1, allowMetricPrefix: true),
  'ga': const _UnitDef(_UnitCategory.magnetism, 0.0001, allowMetricPrefix: true),
};

final Map<String, double> _metricPrefixes = {
  'Y': 1e24,
  'Z': 1e21,
  'E': 1e18,
  'P': 1e15,
  'T': 1e12,
  'G': 1e9,
  'M': 1e6,
  'k': 1e3,
  'h': 1e2,
  'da': 1e1,
  'd': 1e-1,
  'c': 1e-2,
  'm': 1e-3,
  'u': 1e-6,
  'n': 1e-9,
  'p': 1e-12,
  'f': 1e-15,
  'a': 1e-18,
  'z': 1e-21,
  'y': 1e-24,
};

final Map<String, double> _binaryPrefixes = {
  'Yi': math.pow(2, 80).toDouble(),
  'Zi': math.pow(2, 70).toDouble(),
  'Ei': math.pow(2, 60).toDouble(),
  'Pi': math.pow(2, 50).toDouble(),
  'Ti': math.pow(2, 40).toDouble(),
  'Gi': math.pow(2, 30).toDouble(),
  'Mi': math.pow(2, 20).toDouble(),
  'ki': math.pow(2, 10).toDouble(),
};

/// Look up a unit, trying exact match, then metric prefixes, then binary prefixes.
/// Returns (category, conversionFactorToBase) or null.
({_UnitCategory category, double factor})? _lookupUnit(String unit) {
  // Exact match
  final exact = _units[unit];
  if (exact != null) {
    return (category: exact.category, factor: exact.toBase);
  }

  // Try binary prefixes first (longer prefixes like "ki", "Mi" etc.)
  for (final entry in _binaryPrefixes.entries) {
    if (unit.startsWith(entry.key)) {
      final baseUnit = unit.substring(entry.key.length);
      final baseDef = _units[baseUnit];
      if (baseDef != null && baseDef.allowBinaryPrefix) {
        return (category: baseDef.category, factor: baseDef.toBase * entry.value);
      }
    }
  }

  // Try metric prefixes (try longest prefix first: "da" is 2 chars)
  // Sort by prefix length descending
  final sortedPrefixes = _metricPrefixes.entries.toList()
    ..sort((a, b) => b.key.length.compareTo(a.key.length));
  for (final entry in sortedPrefixes) {
    if (unit.startsWith(entry.key)) {
      final baseUnit = unit.substring(entry.key.length);
      final baseDef = _units[baseUnit];
      if (baseDef != null && baseDef.allowMetricPrefix) {
        return (category: baseDef.category, factor: baseDef.toBase * entry.value);
      }
    }
  }

  return null;
}

/// Convert temperature to Kelvin.
double _toKelvin(double value, String unit) {
  switch (unit) {
    case 'C':
    case 'cel':
      return value + 273.15;
    case 'F':
    case 'fah':
      return (value - 32) * 5 / 9 + 273.15;
    case 'K':
    case 'kel':
      return value;
    case 'Rank':
      return value * 5 / 9;
    case 'Reau':
      return value * 5 / 4 + 273.15;
    default:
      return value;
  }
}

/// Convert Kelvin to target temperature unit.
double _fromKelvin(double kelvin, String unit) {
  switch (unit) {
    case 'C':
    case 'cel':
      return kelvin - 273.15;
    case 'F':
    case 'fah':
      return (kelvin - 273.15) * 9 / 5 + 32;
    case 'K':
    case 'kel':
      return kelvin;
    case 'Rank':
      return kelvin * 9 / 5;
    case 'Reau':
      return (kelvin - 273.15) * 4 / 5;
    default:
      return kelvin;
  }
}

// Bit operation maximum value: 2^48 - 1
const int _maxBitVal = (1 << 48) - 1;

// =============================================================================
// Wave 1 — Comparison + Bitwise
// =============================================================================

class DeltaFunction extends FormulaFunction {
  @override
  String get name => 'DELTA';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final a = values[0].toNumber()?.toDouble();
    if (a == null) return const FormulaValue.error(FormulaError.value);
    final b = args.length > 1 ? values[1].toNumber()?.toDouble() : 0.0;
    if (b == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(a == b ? 1 : 0);
  }
}

class GestepFunction extends FormulaFunction {
  @override
  String get name => 'GESTEP';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final number = values[0].toNumber()?.toDouble();
    if (number == null) return const FormulaValue.error(FormulaError.value);
    final step = args.length > 1 ? values[1].toNumber()?.toDouble() : 0.0;
    if (step == null) return const FormulaValue.error(FormulaError.value);
    return FormulaValue.number(number >= step ? 1 : 0);
  }
}

FormulaValue _validateBitArgs(List<FormulaValue> values) {
  final a = values[0].toNumber()?.toDouble();
  final b = values[1].toNumber()?.toDouble();
  if (a == null || b == null) {
    return const FormulaValue.error(FormulaError.value);
  }
  if (a < 0 || b < 0 || a != a.truncateToDouble() || b != b.truncateToDouble()) {
    return const FormulaValue.error(FormulaError.num);
  }
  if (a > _maxBitVal || b > _maxBitVal) {
    return const FormulaValue.error(FormulaError.num);
  }
  return const FormulaValue.number(0); // sentinel: OK
}

class BitandFunction extends FormulaFunction {
  @override
  String get name => 'BITAND';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final check = _validateBitArgs(values);
    if (check.isError) return check;
    final a = values[0].toNumber()!.toInt();
    final b = values[1].toNumber()!.toInt();
    return FormulaValue.number(a & b);
  }
}

class BitorFunction extends FormulaFunction {
  @override
  String get name => 'BITOR';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final check = _validateBitArgs(values);
    if (check.isError) return check;
    final a = values[0].toNumber()!.toInt();
    final b = values[1].toNumber()!.toInt();
    return FormulaValue.number(a | b);
  }
}

class BitxorFunction extends FormulaFunction {
  @override
  String get name => 'BITXOR';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final check = _validateBitArgs(values);
    if (check.isError) return check;
    final a = values[0].toNumber()!.toInt();
    final b = values[1].toNumber()!.toInt();
    return FormulaValue.number(a ^ b);
  }
}

class BitlshiftFunction extends FormulaFunction {
  @override
  String get name => 'BITLSHIFT';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final numVal = values[0].toNumber()?.toDouble();
    final shift = values[1].toNumber()?.toInt();
    if (numVal == null || shift == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (numVal < 0 || numVal != numVal.truncateToDouble() || numVal > _maxBitVal) {
      return const FormulaValue.error(FormulaError.num);
    }
    final n = numVal.toInt();
    int result;
    if (shift >= 0) {
      result = n << shift;
    } else {
      result = n >> (-shift);
    }
    if (result < 0 || result > _maxBitVal) {
      return const FormulaValue.error(FormulaError.num);
    }
    return FormulaValue.number(result);
  }
}

class BitrshiftFunction extends FormulaFunction {
  @override
  String get name => 'BITRSHIFT';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final numVal = values[0].toNumber()?.toDouble();
    final shift = values[1].toNumber()?.toInt();
    if (numVal == null || shift == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (numVal < 0 || numVal != numVal.truncateToDouble() || numVal > _maxBitVal) {
      return const FormulaValue.error(FormulaError.num);
    }
    final n = numVal.toInt();
    int result;
    if (shift >= 0) {
      result = n >> shift;
    } else {
      result = n << (-shift);
    }
    if (result < 0 || result > _maxBitVal) {
      return const FormulaValue.error(FormulaError.num);
    }
    return FormulaValue.number(result);
  }
}

// =============================================================================
// Wave 2 — Base Conversion
// =============================================================================

class Bin2decFunction extends FormulaFunction {
  @override
  String get name => 'BIN2DEC';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final s = v.toText();
    final result = _parseBinary(s);
    if (result == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(result);
  }
}

class Bin2hexFunction extends FormulaFunction {
  @override
  String get name => 'BIN2HEX';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final s = values[0].toText();
    final dec = _parseBinary(s);
    if (dec == null) return const FormulaValue.error(FormulaError.num);
    final places = args.length > 1 ? values[1].toNumber()?.toInt() : null;
    if (args.length > 1 && places == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    final result = _toHex(dec, places);
    if (result == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.text(result);
  }
}

class Bin2octFunction extends FormulaFunction {
  @override
  String get name => 'BIN2OCT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final s = values[0].toText();
    final dec = _parseBinary(s);
    if (dec == null) return const FormulaValue.error(FormulaError.num);
    final places = args.length > 1 ? values[1].toNumber()?.toInt() : null;
    if (args.length > 1 && places == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    final result = _toOctal(dec, places);
    if (result == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.text(result);
  }
}

class Dec2binFunction extends FormulaFunction {
  @override
  String get name => 'DEC2BIN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final dec = values[0].toNumber()?.toInt();
    if (dec == null) return const FormulaValue.error(FormulaError.value);
    final places = args.length > 1 ? values[1].toNumber()?.toInt() : null;
    if (args.length > 1 && places == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    final result = _toBinary(dec, places);
    if (result == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.text(result);
  }
}

class Dec2hexFunction extends FormulaFunction {
  @override
  String get name => 'DEC2HEX';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final dec = values[0].toNumber()?.toInt();
    if (dec == null) return const FormulaValue.error(FormulaError.value);
    final places = args.length > 1 ? values[1].toNumber()?.toInt() : null;
    if (args.length > 1 && places == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    final result = _toHex(dec, places);
    if (result == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.text(result);
  }
}

class Dec2octFunction extends FormulaFunction {
  @override
  String get name => 'DEC2OCT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final dec = values[0].toNumber()?.toInt();
    if (dec == null) return const FormulaValue.error(FormulaError.value);
    final places = args.length > 1 ? values[1].toNumber()?.toInt() : null;
    if (args.length > 1 && places == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    final result = _toOctal(dec, places);
    if (result == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.text(result);
  }
}

class Hex2binFunction extends FormulaFunction {
  @override
  String get name => 'HEX2BIN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final s = values[0].toText();
    final dec = _parseHex(s);
    if (dec == null) return const FormulaValue.error(FormulaError.num);
    final places = args.length > 1 ? values[1].toNumber()?.toInt() : null;
    if (args.length > 1 && places == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    final result = _toBinary(dec, places);
    if (result == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.text(result);
  }
}

class Hex2decFunction extends FormulaFunction {
  @override
  String get name => 'HEX2DEC';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final s = v.toText();
    final result = _parseHex(s);
    if (result == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(result);
  }
}

class Hex2octFunction extends FormulaFunction {
  @override
  String get name => 'HEX2OCT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final s = values[0].toText();
    final dec = _parseHex(s);
    if (dec == null) return const FormulaValue.error(FormulaError.num);
    final places = args.length > 1 ? values[1].toNumber()?.toInt() : null;
    if (args.length > 1 && places == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    final result = _toOctal(dec, places);
    if (result == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.text(result);
  }
}

class Oct2binFunction extends FormulaFunction {
  @override
  String get name => 'OCT2BIN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final s = values[0].toText();
    final dec = _parseOctal(s);
    if (dec == null) return const FormulaValue.error(FormulaError.num);
    final places = args.length > 1 ? values[1].toNumber()?.toInt() : null;
    if (args.length > 1 && places == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    final result = _toBinary(dec, places);
    if (result == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.text(result);
  }
}

class Oct2decFunction extends FormulaFunction {
  @override
  String get name => 'OCT2DEC';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final s = v.toText();
    final result = _parseOctal(s);
    if (result == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(result);
  }
}

class Oct2hexFunction extends FormulaFunction {
  @override
  String get name => 'OCT2HEX';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final s = values[0].toText();
    final dec = _parseOctal(s);
    if (dec == null) return const FormulaValue.error(FormulaError.num);
    final places = args.length > 1 ? values[1].toNumber()?.toInt() : null;
    if (args.length > 1 && places == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    final result = _toHex(dec, places);
    if (result == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.text(result);
  }
}

// =============================================================================
// Wave 3 — Number Format + Error Functions
// =============================================================================

class BaseFunction extends FormulaFunction {
  @override
  String get name => 'BASE';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final number = values[0].toNumber()?.toDouble();
    final radix = values[1].toNumber()?.toInt();
    if (number == null || radix == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (number < 0 || number != number.truncateToDouble()) {
      return const FormulaValue.error(FormulaError.num);
    }
    if (radix < 2 || radix > 36) {
      return const FormulaValue.error(FormulaError.num);
    }
    final minLength = args.length > 2 ? values[2].toNumber()?.toInt() : null;
    if (args.length > 2 && minLength == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    var result = number.toInt().toRadixString(radix).toUpperCase();
    if (minLength != null && minLength > 0) {
      result = result.padLeft(minLength, '0');
    }
    return FormulaValue.text(result);
  }
}

class DecimalFunction extends FormulaFunction {
  @override
  String get name => 'DECIMAL';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final text = values[0].toText().toUpperCase();
    final radix = values[1].toNumber()?.toInt();
    if (radix == null) return const FormulaValue.error(FormulaError.value);
    if (radix < 2 || radix > 36) {
      return const FormulaValue.error(FormulaError.num);
    }
    final result = int.tryParse(text, radix: radix);
    if (result == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(result);
  }
}

const Map<String, int> _romanValues = {
  'M': 1000,
  'D': 500,
  'C': 100,
  'L': 50,
  'X': 10,
  'V': 5,
  'I': 1,
};

class ArabicFunction extends FormulaFunction {
  @override
  String get name => 'ARABIC';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    var text = v.toText().trim().toUpperCase();
    if (text.isEmpty) return const FormulaValue.number(0);

    var sign = 1;
    if (text.startsWith('-')) {
      sign = -1;
      text = text.substring(1);
    }

    var result = 0;
    for (int i = 0; i < text.length; i++) {
      final current = _romanValues[text[i]];
      if (current == null) {
        return const FormulaValue.error(FormulaError.value);
      }
      final next = (i + 1 < text.length) ? _romanValues[text[i + 1]] : null;
      if (next != null && next > current) {
        result -= current;
      } else {
        result += current;
      }
    }
    return FormulaValue.number(sign * result);
  }
}

class RomanFunction extends FormulaFunction {
  @override
  String get name => 'ROMAN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  static const _classicPairs = [
    (1000, 'M'),
    (900, 'CM'),
    (500, 'D'),
    (400, 'CD'),
    (100, 'C'),
    (90, 'XC'),
    (50, 'L'),
    (40, 'XL'),
    (10, 'X'),
    (9, 'IX'),
    (5, 'V'),
    (4, 'IV'),
    (1, 'I'),
  ];

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final number = values[0].toNumber()?.toInt();
    if (number == null) return const FormulaValue.error(FormulaError.value);
    if (number < 0 || number > 3999) {
      return const FormulaValue.error(FormulaError.value);
    }
    if (number == 0) return const FormulaValue.text('');

    var remaining = number;
    final buf = StringBuffer();
    for (final (value, symbol) in _classicPairs) {
      while (remaining >= value) {
        buf.write(symbol);
        remaining -= value;
      }
    }
    return FormulaValue.text(buf.toString());
  }
}

class ErfFunction extends FormulaFunction {
  @override
  String get name => 'ERF';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    if (args.length == 1) {
      final x = values[0].toNumber()?.toDouble();
      if (x == null) return const FormulaValue.error(FormulaError.value);
      return FormulaValue.number(_erf(x));
    }
    // Two args: erf(upper) - erf(lower)
    final lower = values[0].toNumber()?.toDouble();
    final upper = values[1].toNumber()?.toDouble();
    if (lower == null || upper == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    return FormulaValue.number(_erf(upper) - _erf(lower));
  }
}

class ErfPreciseFunction extends FormulaFunction {
  @override
  String get name => 'ERF.PRECISE';
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
    return FormulaValue.number(_erf(x));
  }
}

class ErfcFunction extends FormulaFunction {
  @override
  String get name => 'ERFC';
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
    return FormulaValue.number(1.0 - _erf(x));
  }
}

class ErfcPreciseFunction extends FormulaFunction {
  @override
  String get name => 'ERFC.PRECISE';
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
    return FormulaValue.number(1.0 - _erf(x));
  }
}

// =============================================================================
// Wave 4 — Complex Number Functions
// =============================================================================

class ComplexFunction extends FormulaFunction {
  @override
  String get name => 'COMPLEX';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final real = values[0].toNumber()?.toDouble();
    final imag = values[1].toNumber()?.toDouble();
    if (real == null || imag == null) {
      return const FormulaValue.error(FormulaError.value);
    }
    var suffix = 'i';
    if (args.length > 2) {
      suffix = values[2].toText();
      if (suffix != 'i' && suffix != 'j') {
        return const FormulaValue.error(FormulaError.value);
      }
    }
    return FormulaValue.text(_formatComplex(real, imag, suffix));
  }
}

class ImrealFunction extends FormulaFunction {
  @override
  String get name => 'IMREAL';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final parsed = _parseComplex(v.toText());
    if (parsed == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(parsed.real);
  }
}

class ImaginaryFunction extends FormulaFunction {
  @override
  String get name => 'IMAGINARY';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final parsed = _parseComplex(v.toText());
    if (parsed == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(parsed.imag);
  }
}

class ImabsFunction extends FormulaFunction {
  @override
  String get name => 'IMABS';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final parsed = _parseComplex(v.toText());
    if (parsed == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.number(
        math.sqrt(parsed.real * parsed.real + parsed.imag * parsed.imag));
  }
}

class ImargumentFunction extends FormulaFunction {
  @override
  String get name => 'IMARGUMENT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final parsed = _parseComplex(v.toText());
    if (parsed == null) return const FormulaValue.error(FormulaError.num);
    if (parsed.real == 0 && parsed.imag == 0) {
      return const FormulaValue.error(FormulaError.divZero);
    }
    return FormulaValue.number(math.atan2(parsed.imag, parsed.real));
  }
}

class ImconjugateFunction extends FormulaFunction {
  @override
  String get name => 'IMCONJUGATE';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final parsed = _parseComplex(v.toText());
    if (parsed == null) return const FormulaValue.error(FormulaError.num);
    return FormulaValue.text(
        _formatComplex(parsed.real, -parsed.imag, parsed.suffix));
  }
}

class ImsumFunction extends FormulaFunction {
  @override
  String get name => 'IMSUM';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    double real = 0, imag = 0;
    String suffix = 'i';
    for (final v in values) {
      if (v.isError) return v;
      final parsed = _parseComplex(v.toText());
      if (parsed == null) return const FormulaValue.error(FormulaError.num);
      real += parsed.real;
      imag += parsed.imag;
      suffix = parsed.suffix;
    }
    return FormulaValue.text(_formatComplex(real, imag, suffix));
  }
}

class ImsubFunction extends FormulaFunction {
  @override
  String get name => 'IMSUB';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final z1 = _parseComplex(values[0].toText());
    final z2 = _parseComplex(values[1].toText());
    if (z1 == null || z2 == null) {
      return const FormulaValue.error(FormulaError.num);
    }
    return FormulaValue.text(
        _formatComplex(z1.real - z2.real, z1.imag - z2.imag, z1.suffix));
  }
}

class ImproductFunction extends FormulaFunction {
  @override
  String get name => 'IMPRODUCT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => -1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    double real = 1, imag = 0;
    String suffix = 'i';
    for (final v in values) {
      if (v.isError) return v;
      final parsed = _parseComplex(v.toText());
      if (parsed == null) return const FormulaValue.error(FormulaError.num);
      final newReal = real * parsed.real - imag * parsed.imag;
      final newImag = real * parsed.imag + imag * parsed.real;
      real = newReal;
      imag = newImag;
      suffix = parsed.suffix;
    }
    return FormulaValue.text(_formatComplex(real, imag, suffix));
  }
}

class ImdivFunction extends FormulaFunction {
  @override
  String get name => 'IMDIV';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final z1 = _parseComplex(values[0].toText());
    final z2 = _parseComplex(values[1].toText());
    if (z1 == null || z2 == null) {
      return const FormulaValue.error(FormulaError.num);
    }
    final denom = z2.real * z2.real + z2.imag * z2.imag;
    if (denom == 0) return const FormulaValue.error(FormulaError.num);
    final real = (z1.real * z2.real + z1.imag * z2.imag) / denom;
    final imag = (z1.imag * z2.real - z1.real * z2.imag) / denom;
    return FormulaValue.text(_formatComplex(real, imag, z1.suffix));
  }
}

class ImpowerFunction extends FormulaFunction {
  @override
  String get name => 'IMPOWER';
  @override
  int get minArgs => 2;
  @override
  int get maxArgs => 2;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final z = _parseComplex(values[0].toText());
    final n = values[1].toNumber()?.toDouble();
    if (z == null || n == null) {
      return const FormulaValue.error(FormulaError.num);
    }
    if (z.real == 0 && z.imag == 0 && n == 0) {
      return const FormulaValue.text('1');
    }
    // Polar form: z = r * e^(i*theta)
    final r = math.sqrt(z.real * z.real + z.imag * z.imag);
    final theta = math.atan2(z.imag, z.real);
    final newR = math.pow(r, n);
    final newTheta = theta * n;
    final real = newR * math.cos(newTheta);
    final imag = newR * math.sin(newTheta);
    return FormulaValue.text(_formatComplex(real, imag, z.suffix));
  }
}

class ImsqrtFunction extends FormulaFunction {
  @override
  String get name => 'IMSQRT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final z = _parseComplex(v.toText());
    if (z == null) return const FormulaValue.error(FormulaError.num);
    final r = math.sqrt(z.real * z.real + z.imag * z.imag);
    final theta = math.atan2(z.imag, z.real);
    final newR = math.sqrt(r);
    final newTheta = theta / 2;
    final real = newR * math.cos(newTheta);
    final imag = newR * math.sin(newTheta);
    return FormulaValue.text(_formatComplex(real, imag, z.suffix));
  }
}

class ImexpFunction extends FormulaFunction {
  @override
  String get name => 'IMEXP';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final z = _parseComplex(v.toText());
    if (z == null) return const FormulaValue.error(FormulaError.num);
    // e^(a+bi) = e^a * (cos(b) + i*sin(b))
    final ea = math.exp(z.real);
    final real = ea * math.cos(z.imag);
    final imag = ea * math.sin(z.imag);
    return FormulaValue.text(_formatComplex(real, imag, z.suffix));
  }
}

class ImlnFunction extends FormulaFunction {
  @override
  String get name => 'IMLN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final z = _parseComplex(v.toText());
    if (z == null) return const FormulaValue.error(FormulaError.num);
    final r = math.sqrt(z.real * z.real + z.imag * z.imag);
    if (r == 0) return const FormulaValue.error(FormulaError.num);
    final theta = math.atan2(z.imag, z.real);
    return FormulaValue.text(_formatComplex(math.log(r), theta, z.suffix));
  }
}

class Imlog10Function extends FormulaFunction {
  @override
  String get name => 'IMLOG10';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final z = _parseComplex(v.toText());
    if (z == null) return const FormulaValue.error(FormulaError.num);
    final r = math.sqrt(z.real * z.real + z.imag * z.imag);
    if (r == 0) return const FormulaValue.error(FormulaError.num);
    final theta = math.atan2(z.imag, z.real);
    final lnR = math.log(r);
    final ln10 = math.log(10);
    return FormulaValue.text(
        _formatComplex(lnR / ln10, theta / ln10, z.suffix));
  }
}

class Imlog2Function extends FormulaFunction {
  @override
  String get name => 'IMLOG2';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final z = _parseComplex(v.toText());
    if (z == null) return const FormulaValue.error(FormulaError.num);
    final r = math.sqrt(z.real * z.real + z.imag * z.imag);
    if (r == 0) return const FormulaValue.error(FormulaError.num);
    final theta = math.atan2(z.imag, z.real);
    final lnR = math.log(r);
    final ln2 = math.log(2);
    return FormulaValue.text(
        _formatComplex(lnR / ln2, theta / ln2, z.suffix));
  }
}

class ImsinFunction extends FormulaFunction {
  @override
  String get name => 'IMSIN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final z = _parseComplex(v.toText());
    if (z == null) return const FormulaValue.error(FormulaError.num);
    // sin(a+bi) = sin(a)cosh(b) + i*cos(a)sinh(b)
    final real = math.sin(z.real) * _cosh(z.imag);
    final imag = math.cos(z.real) * _sinh(z.imag);
    return FormulaValue.text(_formatComplex(real, imag, z.suffix));
  }
}

class ImcosFunction extends FormulaFunction {
  @override
  String get name => 'IMCOS';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final z = _parseComplex(v.toText());
    if (z == null) return const FormulaValue.error(FormulaError.num);
    // cos(a+bi) = cos(a)cosh(b) - i*sin(a)sinh(b)
    final real = math.cos(z.real) * _cosh(z.imag);
    final imag = -math.sin(z.real) * _sinh(z.imag);
    return FormulaValue.text(_formatComplex(real, imag, z.suffix));
  }
}

class ImtanFunction extends FormulaFunction {
  @override
  String get name => 'IMTAN';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final z = _parseComplex(v.toText());
    if (z == null) return const FormulaValue.error(FormulaError.num);
    // tan = sin/cos
    final sinR = math.sin(z.real) * _cosh(z.imag);
    final sinI = math.cos(z.real) * _sinh(z.imag);
    final cosR = math.cos(z.real) * _cosh(z.imag);
    final cosI = -math.sin(z.real) * _sinh(z.imag);
    final denom = cosR * cosR + cosI * cosI;
    if (denom == 0) return const FormulaValue.error(FormulaError.num);
    final real = (sinR * cosR + sinI * cosI) / denom;
    final imag = (sinI * cosR - sinR * cosI) / denom;
    return FormulaValue.text(_formatComplex(real, imag, z.suffix));
  }
}

class ImsinhFunction extends FormulaFunction {
  @override
  String get name => 'IMSINH';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final z = _parseComplex(v.toText());
    if (z == null) return const FormulaValue.error(FormulaError.num);
    // sinh(a+bi) = sinh(a)cos(b) + i*cosh(a)sin(b)
    final real = _sinh(z.real) * math.cos(z.imag);
    final imag = _cosh(z.real) * math.sin(z.imag);
    return FormulaValue.text(_formatComplex(real, imag, z.suffix));
  }
}

class ImcoshFunction extends FormulaFunction {
  @override
  String get name => 'IMCOSH';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final z = _parseComplex(v.toText());
    if (z == null) return const FormulaValue.error(FormulaError.num);
    // cosh(a+bi) = cosh(a)cos(b) + i*sinh(a)sin(b)
    final real = _cosh(z.real) * math.cos(z.imag);
    final imag = _sinh(z.real) * math.sin(z.imag);
    return FormulaValue.text(_formatComplex(real, imag, z.suffix));
  }
}

class ImsecFunction extends FormulaFunction {
  @override
  String get name => 'IMSEC';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final z = _parseComplex(v.toText());
    if (z == null) return const FormulaValue.error(FormulaError.num);
    // sec = 1/cos
    final cosR = math.cos(z.real) * _cosh(z.imag);
    final cosI = -math.sin(z.real) * _sinh(z.imag);
    final denom = cosR * cosR + cosI * cosI;
    if (denom == 0) return const FormulaValue.error(FormulaError.num);
    final real = cosR / denom;
    final imag = -cosI / denom;
    return FormulaValue.text(_formatComplex(real, imag, z.suffix));
  }
}

class ImsechFunction extends FormulaFunction {
  @override
  String get name => 'IMSECH';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final z = _parseComplex(v.toText());
    if (z == null) return const FormulaValue.error(FormulaError.num);
    // sech = 1/cosh
    final coshR = _cosh(z.real) * math.cos(z.imag);
    final coshI = _sinh(z.real) * math.sin(z.imag);
    final denom = coshR * coshR + coshI * coshI;
    if (denom == 0) return const FormulaValue.error(FormulaError.num);
    final real = coshR / denom;
    final imag = -coshI / denom;
    return FormulaValue.text(_formatComplex(real, imag, z.suffix));
  }
}

class ImcscFunction extends FormulaFunction {
  @override
  String get name => 'IMCSC';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final z = _parseComplex(v.toText());
    if (z == null) return const FormulaValue.error(FormulaError.num);
    // csc = 1/sin
    final sinR = math.sin(z.real) * _cosh(z.imag);
    final sinI = math.cos(z.real) * _sinh(z.imag);
    final denom = sinR * sinR + sinI * sinI;
    if (denom == 0) return const FormulaValue.error(FormulaError.num);
    final real = sinR / denom;
    final imag = -sinI / denom;
    return FormulaValue.text(_formatComplex(real, imag, z.suffix));
  }
}

class ImcschFunction extends FormulaFunction {
  @override
  String get name => 'IMCSCH';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final z = _parseComplex(v.toText());
    if (z == null) return const FormulaValue.error(FormulaError.num);
    // csch = 1/sinh
    final sinhR = _sinh(z.real) * math.cos(z.imag);
    final sinhI = _cosh(z.real) * math.sin(z.imag);
    final denom = sinhR * sinhR + sinhI * sinhI;
    if (denom == 0) return const FormulaValue.error(FormulaError.num);
    final real = sinhR / denom;
    final imag = -sinhI / denom;
    return FormulaValue.text(_formatComplex(real, imag, z.suffix));
  }
}

class ImcotFunction extends FormulaFunction {
  @override
  String get name => 'IMCOT';
  @override
  int get minArgs => 1;
  @override
  int get maxArgs => 1;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final v = args[0].evaluate(context);
    if (v.isError) return v;
    final z = _parseComplex(v.toText());
    if (z == null) return const FormulaValue.error(FormulaError.num);
    // cot = cos/sin
    final sinR = math.sin(z.real) * _cosh(z.imag);
    final sinI = math.cos(z.real) * _sinh(z.imag);
    final cosR = math.cos(z.real) * _cosh(z.imag);
    final cosI = -math.sin(z.real) * _sinh(z.imag);
    final denom = sinR * sinR + sinI * sinI;
    if (denom == 0) return const FormulaValue.error(FormulaError.num);
    final real = (cosR * sinR + cosI * sinI) / denom;
    final imag = (cosI * sinR - cosR * sinI) / denom;
    return FormulaValue.text(_formatComplex(real, imag, z.suffix));
  }
}

// Hyperbolic helpers
double _sinh(double x) => (math.exp(x) - math.exp(-x)) / 2;
double _cosh(double x) => (math.exp(x) + math.exp(-x)) / 2;

// =============================================================================
// Wave 5 — CONVERT
// =============================================================================

class ConvertFunction extends FormulaFunction {
  @override
  String get name => 'CONVERT';
  @override
  int get minArgs => 3;
  @override
  int get maxArgs => 3;

  @override
  FormulaValue call(List<FormulaNode> args, EvaluationContext context) {
    final values = evaluateArgs(args, context);
    final number = values[0].toNumber()?.toDouble();
    if (number == null) return const FormulaValue.error(FormulaError.value);
    final fromUnit = values[1].toText();
    final toUnit = values[2].toText();

    final from = _lookupUnit(fromUnit);
    final to = _lookupUnit(toUnit);

    if (from == null || to == null) {
      return const FormulaValue.error(FormulaError.na);
    }
    if (from.category != to.category) {
      return const FormulaValue.error(FormulaError.na);
    }

    // Special case: temperature
    if (from.category == _UnitCategory.temperature) {
      final kelvin = _toKelvin(number, fromUnit);
      final result = _fromKelvin(kelvin, toUnit);
      return FormulaValue.number(result);
    }

    // Standard conversion: value * from_factor / to_factor
    final result = number * from.factor / to.factor;
    return FormulaValue.number(result);
  }
}

// =============================================================================
// Registration
// =============================================================================

void registerEngineeringFunctions(FunctionRegistry registry) {
  registry.registerAll([
    // Wave 1
    DeltaFunction(),
    GestepFunction(),
    BitandFunction(),
    BitorFunction(),
    BitxorFunction(),
    BitlshiftFunction(),
    BitrshiftFunction(),
    // Wave 2
    Bin2decFunction(),
    Bin2hexFunction(),
    Bin2octFunction(),
    Dec2binFunction(),
    Dec2hexFunction(),
    Dec2octFunction(),
    Hex2binFunction(),
    Hex2decFunction(),
    Hex2octFunction(),
    Oct2binFunction(),
    Oct2decFunction(),
    Oct2hexFunction(),
    // Wave 3
    BaseFunction(),
    DecimalFunction(),
    ArabicFunction(),
    RomanFunction(),
    ErfFunction(),
    ErfPreciseFunction(),
    ErfcFunction(),
    ErfcPreciseFunction(),
    // Wave 4
    ComplexFunction(),
    ImrealFunction(),
    ImaginaryFunction(),
    ImabsFunction(),
    ImargumentFunction(),
    ImconjugateFunction(),
    ImsumFunction(),
    ImsubFunction(),
    ImproductFunction(),
    ImdivFunction(),
    ImpowerFunction(),
    ImsqrtFunction(),
    ImexpFunction(),
    ImlnFunction(),
    Imlog10Function(),
    Imlog2Function(),
    ImsinFunction(),
    ImcosFunction(),
    ImtanFunction(),
    ImsinhFunction(),
    ImcoshFunction(),
    ImsecFunction(),
    ImsechFunction(),
    ImcscFunction(),
    ImcschFunction(),
    ImcotFunction(),
    // Wave 5
    ConvertFunction(),
  ]);
}
