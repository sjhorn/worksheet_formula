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
  List<A1> getCellsToRecalculate(A1 changedCell) {
    final result = <A1>[];
    final visited = <A1>{};
    final inProgress = <A1>{};

    void visit(A1 cell) {
      if (visited.contains(cell)) return;
      if (inProgress.contains(cell)) {
        // Circular dependency detected - skip but don't throw
        return;
      }

      inProgress.add(cell);

      for (final dependent in _dependents[cell] ?? <A1>{}) {
        visit(dependent);
      }

      inProgress.remove(cell);
      visited.add(cell);
      result.add(cell);
    }

    for (final dependent in _dependents[changedCell] ?? <A1>{}) {
      visit(dependent);
    }

    return result.reversed.toList();
  }

  /// Check if there's a circular reference involving the given cell.
  bool hasCircularReference(A1 cell) {
    final visited = <A1>{};

    bool visit(A1 current, A1 target) {
      if (current == target && visited.isNotEmpty) return true;
      if (visited.contains(current)) return false;
      visited.add(current);

      for (final dep in _dependencies[current] ?? <A1>{}) {
        if (visit(dep, target)) return true;
      }
      return false;
    }

    return visit(cell, cell);
  }

  /// Clear all dependency information.
  void clear() {
    _dependents.clear();
    _dependencies.clear();
  }
}
