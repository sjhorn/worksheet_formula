import 'array.dart';
import 'date.dart';
import 'function.dart';
import 'information.dart';
import 'logical.dart';
import 'lookup.dart';
import 'math.dart';
import 'statistical.dart';
import 'text.dart';

/// Registry of available formula functions.
class FunctionRegistry {
  final Map<String, FormulaFunction> _functions = {};

  /// Create a new registry, optionally with built-in functions pre-registered.
  FunctionRegistry({bool registerBuiltIns = true}) {
    if (registerBuiltIns) {
      registerMathFunctions(this);
      registerLogicalFunctions(this);
      registerTextFunctions(this);
      registerStatisticalFunctions(this);
      registerLookupFunctions(this);
      registerDateFunctions(this);
      registerInformationFunctions(this);
      registerArrayFunctions(this);
    }
  }

  /// Register a function.
  void register(FormulaFunction function) {
    _functions[function.name.toUpperCase()] = function;
  }

  /// Register multiple functions.
  void registerAll(Iterable<FormulaFunction> functions) {
    for (final func in functions) {
      register(func);
    }
  }

  /// Get a function by name (case-insensitive).
  FormulaFunction? get(String name) => _functions[name.toUpperCase()];

  /// Check if a function exists.
  bool has(String name) => _functions.containsKey(name.toUpperCase());

  /// Get all registered function names.
  Iterable<String> get names => _functions.keys;

  /// Create a copy with additional functions.
  FunctionRegistry copyWith(Iterable<FormulaFunction> additional) {
    final copy = FunctionRegistry(registerBuiltIns: false);
    copy._functions.addAll(_functions);
    copy.registerAll(additional);
    return copy;
  }
}
