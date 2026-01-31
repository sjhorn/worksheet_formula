import 'package:a1/a1.dart';
import 'package:test/test.dart';
import 'package:worksheet_formula/src/dependencies/graph.dart';

void main() {
  late DependencyGraph graph;

  setUp(() {
    graph = DependencyGraph();
  });

  group('updateDependencies', () {
    test('adds dependency links', () {
      // A3 depends on A1 and A2
      graph.updateDependencies('A3'.a1, {'A1'.a1, 'A2'.a1});

      expect(graph.getDependencies('A3'.a1), {'A1'.a1, 'A2'.a1});
      expect(graph.getDependents('A1'.a1), {'A3'.a1});
      expect(graph.getDependents('A2'.a1), {'A3'.a1});
    });

    test('updates removes old links', () {
      graph.updateDependencies('A3'.a1, {'A1'.a1, 'A2'.a1});
      // Now A3 only depends on B1
      graph.updateDependencies('A3'.a1, {'B1'.a1});

      expect(graph.getDependencies('A3'.a1), {'B1'.a1});
      expect(graph.getDependents('A1'.a1), isEmpty);
      expect(graph.getDependents('A2'.a1), isEmpty);
      expect(graph.getDependents('B1'.a1), {'A3'.a1});
    });

    test('empty dependencies removes cell', () {
      graph.updateDependencies('A3'.a1, {'A1'.a1});
      graph.updateDependencies('A3'.a1, {});

      expect(graph.getDependencies('A3'.a1), isEmpty);
      expect(graph.getDependents('A1'.a1), isEmpty);
    });
  });

  group('removeCell', () {
    test('removes all dependency info for a cell', () {
      graph.updateDependencies('A3'.a1, {'A1'.a1, 'A2'.a1});
      graph.removeCell('A3'.a1);

      expect(graph.getDependencies('A3'.a1), isEmpty);
      expect(graph.getDependents('A1'.a1), isEmpty);
      expect(graph.getDependents('A2'.a1), isEmpty);
    });
  });

  group('getDependents', () {
    test('returns empty set for unknown cell', () {
      expect(graph.getDependents('A1'.a1), isEmpty);
    });

    test('returns all cells that depend on given cell', () {
      graph.updateDependencies('B1'.a1, {'A1'.a1});
      graph.updateDependencies('C1'.a1, {'A1'.a1});

      expect(graph.getDependents('A1'.a1), {'B1'.a1, 'C1'.a1});
    });
  });

  group('getDependencies', () {
    test('returns empty set for unknown cell', () {
      expect(graph.getDependencies('A1'.a1), isEmpty);
    });
  });

  group('getCellsToRecalculate', () {
    test('returns direct dependents', () {
      graph.updateDependencies('B1'.a1, {'A1'.a1});
      graph.updateDependencies('C1'.a1, {'A1'.a1});

      final cells = graph.getCellsToRecalculate('A1'.a1);
      expect(cells, containsAll(['B1'.a1, 'C1'.a1]));
    });

    test('returns transitive dependents in order', () {
      // A1 -> B1 -> C1
      graph.updateDependencies('B1'.a1, {'A1'.a1});
      graph.updateDependencies('C1'.a1, {'B1'.a1});

      final cells = graph.getCellsToRecalculate('A1'.a1);
      expect(cells.indexOf('B1'.a1), lessThan(cells.indexOf('C1'.a1)));
    });

    test('handles diamond dependencies', () {
      // A1 -> B1, A1 -> C1, B1 -> D1, C1 -> D1
      graph.updateDependencies('B1'.a1, {'A1'.a1});
      graph.updateDependencies('C1'.a1, {'A1'.a1});
      graph.updateDependencies('D1'.a1, {'B1'.a1, 'C1'.a1});

      final cells = graph.getCellsToRecalculate('A1'.a1);
      expect(cells, contains('D1'.a1));
      // D1 should come after both B1 and C1
      expect(cells.indexOf('D1'.a1), greaterThan(cells.indexOf('B1'.a1)));
      expect(cells.indexOf('D1'.a1), greaterThan(cells.indexOf('C1'.a1)));
    });

    test('handles circular references without infinite loop', () {
      // A1 -> B1 -> A1 (circular)
      graph.updateDependencies('B1'.a1, {'A1'.a1});
      graph.updateDependencies('A1'.a1, {'B1'.a1});

      // Should not throw or loop infinitely
      final cells = graph.getCellsToRecalculate('A1'.a1);
      expect(cells, isNotEmpty);
    });

    test('returns empty for cell with no dependents', () {
      graph.updateDependencies('B1'.a1, {'A1'.a1});
      expect(graph.getCellsToRecalculate('B1'.a1), isEmpty);
    });
  });

  group('hasCircularReference', () {
    test('returns false when no circular reference', () {
      graph.updateDependencies('B1'.a1, {'A1'.a1});
      expect(graph.hasCircularReference('B1'.a1), isFalse);
    });

    test('detects direct circular reference', () {
      // A1 depends on B1, B1 depends on A1
      graph.updateDependencies('A1'.a1, {'B1'.a1});
      graph.updateDependencies('B1'.a1, {'A1'.a1});

      expect(graph.hasCircularReference('A1'.a1), isTrue);
      expect(graph.hasCircularReference('B1'.a1), isTrue);
    });

    test('detects indirect circular reference', () {
      // A1 -> B1 -> C1 -> A1
      graph.updateDependencies('A1'.a1, {'C1'.a1});
      graph.updateDependencies('B1'.a1, {'A1'.a1});
      graph.updateDependencies('C1'.a1, {'B1'.a1});

      expect(graph.hasCircularReference('A1'.a1), isTrue);
    });

    test('returns false for isolated cell', () {
      expect(graph.hasCircularReference('A1'.a1), isFalse);
    });
  });

  group('clear', () {
    test('removes all dependency information', () {
      graph.updateDependencies('B1'.a1, {'A1'.a1});
      graph.updateDependencies('C1'.a1, {'A1'.a1});
      graph.clear();

      expect(graph.getDependents('A1'.a1), isEmpty);
      expect(graph.getDependencies('B1'.a1), isEmpty);
      expect(graph.getDependencies('C1'.a1), isEmpty);
    });
  });
}
