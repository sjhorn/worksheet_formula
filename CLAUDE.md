# Worksheet Formula

## Project Overview

Explain in PLAN.md

## Development Principles

### TDD Workflow
1. **Write test first** - Define expected behavior before implementation
2. **Red → Green → Refactor** - Failing test → Pass → Optimize
3. **Test file mirrors source** - `lib/src/core/span_list.dart` → `test/core/span_list_test.dart`
4. **Minimum 80% coverage** - Critical paths require 100%

### SOLID Principles
- **S**: Each class has one responsibility 
- **O**: Extend via interfaces, not modification
- **L**: Subtypes must be substitutable
- **I**: Small, focused interfaces
- **D**: Depend on abstractions

### Dart Idioms
- Prefer `final` and immutable models
- Use factory constructors for complex initialization
- Extension methods for utility functions
- `typedef` for function signatures
- Named parameters with required keyword


## Testing Strategy

### Unit Tests
- All pure functions and models
- Mock dependencies via interfaces
- Property-based tests for math operations

### Integration Tests
- Large dataset performance
- Memory leak detection


## Commands
```bash
# Run tests with coverage
dart test --coverage

# Generate coverage report
genhtml coverage/lcov.info -o coverage/html

# Run specific test file
dart test test/core/file.dart

```

## Code Review Checklist
- [ ] Tests written before implementation
- [ ] All public APIs documented
- [ ] No magic numbers (use constants)
- [ ] Interfaces for external dependencies
- [ ] Immutable models where possible
- [ ] Memory disposal in `dispose()` methods
- [ ] Performance-critical code benchmarked

## Release Process

Follow these steps in order. Fix any issues before proceeding to the next step.

### 1. Static Analysis
```bash
# Run the analyzer — must have zero issues
dart analyze

# Apply automated fixes for any issues
dart fix --apply

# Re-run analyzer to confirm clean
dart analyze
```

### 2. Tests
```bash
# Run all tests — must all pass
dart test
```

### 3. Coverage
```bash
# Generate coverage data
dart test --coverage

# Generate HTML report and review
genhtml coverage/lcov.info -o coverage/html
open coverage/html/index.html

# Verify minimum 80% coverage (critical paths 100%)
```

### 4. Version & Changelog
- Bump version in `pubspec.yaml` following [semver](https://semver.org/)
  - **patch** (1.0.x): bug fixes
  - **minor** (1.x.0): new features, backwards compatible
  - **major** (x.0.0): breaking API changes
- Add entry to `CHANGELOG.md` under new version heading with date
- Update any version references in `README.md` if needed

### 5. Commit & Tag
```bash
git add -A
git commit -m "chore: release vX.Y.Z"
git tag vX.Y.Z
git push && git push --tags
```

### 6. Publish to pub.dev
```bash
# Dry run first — fix any issues it reports
dart pub publish --dry-run

# Publish for real
dart pub publish
```