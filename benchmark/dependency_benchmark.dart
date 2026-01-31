// ignore_for_file: avoid_print

import 'package:a1/a1.dart';
import 'package:worksheet_formula/src/dependencies/graph.dart';

void main() {
  print('=== Dependency Graph Benchmarks ===\n');

  _benchmarkUpdateDependencies(1000);
  _benchmarkUpdateDependencies(10000);
  _benchmarkGetCellsToRecalculate(1000);
  _benchmarkGetCellsToRecalculate(10000);
  _benchmarkHasCircularReference(1000);
}

void _benchmarkUpdateDependencies(int cellCount) {
  final sw = Stopwatch()..start();
  final graph = DependencyGraph();

  for (var i = 0; i < cellCount; i++) {
    final cell = A1.fromVector(i % 26, i ~/ 26);
    final deps = <A1>{};
    // Each cell depends on 1-3 previous cells
    if (i > 0) deps.add(A1.fromVector((i - 1) % 26, (i - 1) ~/ 26));
    if (i > 1) deps.add(A1.fromVector((i - 2) % 26, (i - 2) ~/ 26));
    graph.updateDependencies(cell, deps);
  }

  sw.stop();
  print('  updateDependencies ($cellCount cells)');
  print('    ${sw.elapsedMilliseconds}ms\n');
}

void _benchmarkGetCellsToRecalculate(int cellCount) {
  final graph = DependencyGraph();

  // Build a linear chain: cell_0 -> cell_1 -> cell_2 -> ...
  for (var i = 1; i < cellCount; i++) {
    final cell = A1.fromVector(i % 26, i ~/ 26);
    final dep = A1.fromVector((i - 1) % 26, (i - 1) ~/ 26);
    graph.updateDependencies(cell, {dep});
  }

  final root = A1.fromVector(0, 0);
  const iterations = 100;

  // Warmup
  for (var i = 0; i < 5; i++) {
    graph.getCellsToRecalculate(root);
  }

  final sw = Stopwatch()..start();
  for (var i = 0; i < iterations; i++) {
    graph.getCellsToRecalculate(root);
  }
  sw.stop();

  final opsPerSec = (iterations / sw.elapsedMicroseconds * 1000000).round();
  print('  getCellsToRecalculate ($cellCount cell chain)');
  print('    ${sw.elapsedMilliseconds}ms for $iterations ops'
      ' ($opsPerSec ops/sec)\n');
}

void _benchmarkHasCircularReference(int cellCount) {
  final graph = DependencyGraph();

  // Build a linear chain (no cycles)
  for (var i = 1; i < cellCount; i++) {
    final cell = A1.fromVector(i % 26, i ~/ 26);
    final dep = A1.fromVector((i - 1) % 26, (i - 1) ~/ 26);
    graph.updateDependencies(cell, {dep});
  }

  final lastCell = A1.fromVector((cellCount - 1) % 26, (cellCount - 1) ~/ 26);
  const iterations = 100;

  // Warmup
  for (var i = 0; i < 5; i++) {
    graph.hasCircularReference(lastCell);
  }

  final sw = Stopwatch()..start();
  for (var i = 0; i < iterations; i++) {
    graph.hasCircularReference(lastCell);
  }
  sw.stop();

  final opsPerSec = (iterations / sw.elapsedMicroseconds * 1000000).round();
  print('  hasCircularReference ($cellCount cell chain, no cycle)');
  print('    ${sw.elapsedMilliseconds}ms for $iterations ops'
      ' ($opsPerSec ops/sec)\n');
}
