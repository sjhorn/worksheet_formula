// ignore_for_file: avoid_print

import 'package:worksheet_formula/src/parser/formula_parser.dart';

void main() {
  final parser = FormulaParser();

  print('=== Parse Benchmarks ===\n');

  _benchmark('Simple arithmetic (1+2)', () => parser.parse('1+2'));

  _benchmark('Cell references (A1+B1+C1+D1+E1)',
      () => parser.parse('A1+B1+C1+D1+E1'));

  _benchmark('Range formula (SUM(A1:Z100))',
      () => parser.parse('=SUM(A1:Z100)'));

  _benchmark('Nested functions (IF(AND(...)))',
      () => parser.parse('=IF(AND(A1>0,B1<100),SUM(C1:C10)*2,"N/A")'));

  _benchmark('String formula (CONCAT)', () => parser.parse('=CONCAT("hello"," ","world")'));

  _benchmark('Complex (chained operators)',
      () => parser.parse('=A1+B1*C1-D1/E1^2&"text"'));
}

void _benchmark(String label, void Function() fn,
    [int iterations = 10000]) {
  // Warmup
  for (var i = 0; i < 100; i++) {
    fn();
  }

  final sw = Stopwatch()..start();
  for (var i = 0; i < iterations; i++) {
    fn();
  }
  sw.stop();

  final opsPerSec = (iterations / sw.elapsedMicroseconds * 1000000).round();
  print('  $label');
  print('    ${sw.elapsedMilliseconds}ms for $iterations ops'
      ' ($opsPerSec ops/sec)\n');
}
