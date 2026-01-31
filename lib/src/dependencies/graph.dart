import 'package:a1/a1.dart';

/// Tracks cell dependencies for efficient recalculation.
class DependencyGraph {
  /// Map from cell -> cells that depend on it (dependents).
  final Map<A1, Set<A1>> _dependents = {};

  /// Map from cell -> cells it depends on (dependencies).
  final Map<A1, Set<A1>> _dependencies = {};

  /// Update the dependencies for a cell.
  ///
  /// Call this when a cell's formula changes.
  void updateDependencies(A1 cell, Set<A1> newDependencies) {
    // Remove old dependency links
    final oldDeps = _dependencies[cell];
    if (oldDeps != null) {
      for (final dep in oldDeps) {
        _dependents[dep]?.remove(cell);
      }
    }

    // Add new dependency links
    if (newDependencies.isEmpty) {
      _dependencies.remove(cell);
    } else {
      _dependencies[cell] = newDependencies;
      for (final dep in newDependencies) {
        (_dependents[dep] ??= {}).add(cell);
      }
    }
  }

  /// Remove a cell from the dependency graph.
  ///
  /// Call this when a cell is cleared or deleted.
  void removeCell(A1 cell) {
    updateDependencies(cell, {});
    _dependents.remove(cell);
  }

  /// Get all cells that depend on the given cell.
  Set<A1> getDependents(A1 cell) => _dependents[cell] ?? {};

  /// Get all cells that the given cell depends on.
  Set<A1> getDependencies(A1 cell) => _dependencies[cell] ?? {};

  /// Get all cells that need recalculation when a cell changes.
  ///
  /// Returns cells in topological order (dependencies before dependents).
  /// Uses iterative DFS to avoid stack overflow with deep chains.
  List<A1> getCellsToRecalculate(A1 changedCell) {
    final result = <A1>[];
    final visited = <A1>{};
    final inProgress = <A1>{};

    // Iterative post-order DFS using explicit stack
    final stack = <(A1, bool)>[]; // (cell, expanded)
    for (final dependent in _dependents[changedCell] ?? <A1>{}) {
      stack.add((dependent, false));
    }

    while (stack.isNotEmpty) {
      final (cell, expanded) = stack.removeLast();

      if (visited.contains(cell)) continue;

      if (expanded) {
        // All children have been processed
        inProgress.remove(cell);
        visited.add(cell);
        result.add(cell);
        continue;
      }

      if (inProgress.contains(cell)) {
        // Circular dependency detected - skip
        continue;
      }

      inProgress.add(cell);
      // Push this cell again as "expanded" (will be processed after children)
      stack.add((cell, true));

      // Push children
      for (final dependent in _dependents[cell] ?? <A1>{}) {
        if (!visited.contains(dependent)) {
          stack.add((dependent, false));
        }
      }
    }

    return result.reversed.toList();
  }

  /// Check if there's a circular reference involving the given cell.
  /// Uses iterative DFS to avoid stack overflow with deep chains.
  bool hasCircularReference(A1 cell) {
    final visited = <A1>{};
    final stack = <A1>[cell];

    while (stack.isNotEmpty) {
      final current = stack.removeLast();

      if (current == cell && visited.isNotEmpty) return true;
      if (visited.contains(current)) continue;
      visited.add(current);

      for (final dep in _dependencies[current] ?? <A1>{}) {
        stack.add(dep);
      }
    }

    return false;
  }

  /// Clear all dependency information.
  void clear() {
    _dependents.clear();
    _dependencies.clear();
  }
}
