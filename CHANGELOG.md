## 1.1.0

- **Immediate LAMBDA invocation** — `=LAMBDA(x,x+1)(5)` now returns `6`
- New `CallExpressionNode` AST node for postfix call expressions
- Parser supports `expr(args)` syntax for calling function-producing expressions
- Chained/curried lambdas work: `=LAMBDA(x,LAMBDA(y,x+y))(1)(2)` returns `3`
- Non-function invocation returns `#VALUE!` (e.g. `=(5)(3)`)
- 1634 tests passing

## 1.0.0

- **1.0 release** — 400 built-in functions across 14 categories
- Phase 9: 16 remaining & niche functions (ENCODEURL, REGEXMATCH, REGEXEXTRACT, REGEXREPLACE, SERIESSUM, SQRTPI, MULTINOMIAL, MUNIT, MMULT, MDETERM, MINVERSE, ARRAYTOTEXT, VALUETOTEXT, ASC, DBCS, BAHTTEXT)
- New category: Web & Regex (4 functions)
- Matrix functions: MUNIT, MMULT, MDETERM (LU decomposition), MINVERSE (Gauss-Jordan elimination)
- CJK text functions: ASC, DBCS, BAHTTEXT
- Phase 10 roadmap for 7 deferred functions (context-dependent and out-of-scope)
- Updated README, pubspec, and documentation to reflect 400-function state
- 2 functions deferred: GETPIVOTDATA (pivot table metadata), PHONETIC (IME/furigana data)
- 1625+ tests passing

## 0.9.0

- **Lambda & higher-order functions** (9): LAMBDA, LET, MAP, REDUCE, SCAN, MAKEARRAY, BYCOL, BYROW, ISOMITTED
- LAMBDA creates custom function closures (FunctionValue)
- LET defines named variables within formulas
- MAP, REDUCE, SCAN for functional array processing
- MAKEARRAY builds arrays from lambda with 1-based indices
- BYCOL, BYROW apply lambda to each column/row
- ISOMITTED tests if a lambda argument was omitted (OmittedValue)
- Structural changes: NameNode in AST, FunctionValue/OmittedValue in values, getVariable() on EvaluationContext, bare identifier parsing
- 384 total built-in functions across 13 categories

## 0.8.0

- **Database functions** (12): DSUM, DAVERAGE, DCOUNT, DCOUNTA, DMAX, DMIN, DGET, DPRODUCT, DSTDEV, DSTDEVP, DVAR, DVARP
- Structured criteria matching against database ranges (header row + data rows)
- Support for comparison operators in criteria (>, <, >=, <=, <>)
- DGET returns #VALUE! when multiple rows match, #VALUE! when none match
- 375 total built-in functions across 12 categories

## 0.7.0

- **Engineering functions** (54): DELTA, GESTEP, BITAND, BITOR, BITXOR, BITLSHIFT, BITRSHIFT, BIN2DEC, BIN2HEX, BIN2OCT, DEC2BIN, DEC2HEX, DEC2OCT, HEX2BIN, HEX2DEC, HEX2OCT, OCT2BIN, OCT2DEC, OCT2HEX, BASE, DECIMAL, ARABIC, ROMAN, ERF, ERF.PRECISE, ERFC, ERFC.PRECISE, COMPLEX, IMREAL, IMAGINARY, IMABS, IMARGUMENT, IMCONJUGATE, IMSUM, IMSUB, IMPRODUCT, IMDIV, IMPOWER, IMSQRT, IMEXP, IMLN, IMLOG10, IMLOG2, IMSIN, IMCOS, IMTAN, IMSINH, IMCOSH, IMSEC, IMSECH, IMCSC, IMCSCH, IMCOT, CONVERT
- Comparison & bitwise: DELTA, GESTEP, BITAND, BITOR, BITXOR, BITLSHIFT, BITRSHIFT
- Base conversion: BIN2DEC, BIN2HEX, BIN2OCT, DEC2BIN, DEC2HEX, DEC2OCT, HEX2BIN, HEX2DEC, HEX2OCT, OCT2BIN, OCT2DEC, OCT2HEX, BASE, DECIMAL
- Number format: ARABIC, ROMAN
- Error functions: ERF, ERF.PRECISE, ERFC, ERFC.PRECISE
- Complex number arithmetic: 26 IM* functions (COMPLEX, IMREAL, IMAGINARY, IMABS, etc.)
- Unit conversion: CONVERT with 13 categories, ~100 units, metric + binary prefixes
- 363 total built-in functions across 11 categories

## 0.6.0

- **Advanced statistical & probability functions** (70): FISHER, FISHERINV, STANDARDIZE, PERMUT, PERMUTATIONA, DEVSQ, KURT, SKEW, SKEW.P, COVARIANCE.P, COVARIANCE.S, CORREL, PEARSON, RSQ, SLOPE, INTERCEPT, STEYX, FORECAST.LINEAR, PROB, MODE.MULT, STDEVA, STDEVPA, VARA, VARPA, GAMMA, GAMMALN, GAMMALN.PRECISE, GAUSS, PHI, NORM.S.DIST, NORM.S.INV, NORM.DIST, NORM.INV, BINOM.DIST, BINOM.INV, BINOM.DIST.RANGE, NEGBINOM.DIST, HYPGEOM.DIST, POISSON.DIST, EXPON.DIST, GAMMA.DIST, GAMMA.INV, BETA.DIST, BETA.INV, CHISQ.DIST, CHISQ.INV, CHISQ.DIST.RT, CHISQ.INV.RT, T.DIST, T.INV, T.DIST.2T, T.INV.2T, T.DIST.RT, F.DIST, F.INV, F.DIST.RT, F.INV.RT, WEIBULL.DIST, LOGNORM.DIST, LOGNORM.INV, CONFIDENCE.NORM, CONFIDENCE.T, Z.TEST, T.TEST, CHISQ.TEST, F.TEST, LINEST, LOGEST, TREND, GROWTH
- Distributions: Normal, Student's t, Chi-squared, F, Binomial, Negative Binomial, Hypergeometric, Poisson, Exponential, Gamma, Beta, Weibull, Lognormal (CDF, PDF, inverse)
- Regression & correlation: CORREL, PEARSON, RSQ, SLOPE, INTERCEPT, STEYX, LINEST, LOGEST, TREND, GROWTH, FORECAST.LINEAR
- Hypothesis testing: T.TEST, CHISQ.TEST, F.TEST, Z.TEST, CONFIDENCE.NORM, CONFIDENCE.T
- Other: FISHER, FISHERINV, STANDARDIZE, PERMUT, PERMUTATIONA, DEVSQ, KURT, SKEW, SKEW.P, COVARIANCE.P/S, PROB, MODE.MULT, STDEVA, STDEVPA, VARA, VARPA, GAMMA, GAMMALN, GAUSS, PHI
- Lanczos approximation for log-gamma, regularized incomplete beta/gamma, bisection inverse CDF solver
- 309 total built-in functions across 10 categories

## 0.5.0

- **Financial functions** (40): PMT, FV, PV, NPER, RATE, IPMT, PPMT, CUMIPMT, CUMPRINC, NPV, XNPV, IRR, XIRR, MIRR, FVSCHEDULE, SLN, SYD, DB, DDB, VDB, PRICE, YIELD, DURATION, MDURATION, ACCRINT, DISC, INTRATE, RECEIVED, PRICEDISC, PRICEMAT, TBILLEQ, TBILLPRICE, TBILLYIELD, DOLLARDE, DOLLARFR, EFFECT, NOMINAL, PDURATION, RRI, ISPMT
- Loans & annuities (TVM): PMT, FV, PV, NPER, RATE, IPMT, PPMT, CUMIPMT, CUMPRINC
- Investment analysis: NPV, XNPV, IRR, XIRR, MIRR, FVSCHEDULE
- Depreciation: SLN, SYD, DB, DDB, VDB
- Bonds & securities: PRICE, YIELD, DURATION, MDURATION, ACCRINT, DISC, INTRATE, RECEIVED, PRICEDISC, PRICEMAT, TBILLEQ, TBILLPRICE, TBILLYIELD, DOLLARDE, DOLLARFR
- Other financial: EFFECT, NOMINAL, PDURATION, RRI, ISPMT
- Newton-Raphson solver for RATE, IRR, XIRR, YIELD
- Day count basis support (US 30/360, actual/actual, actual/360, actual/365, EU 30/360)
- 239 total built-in functions across 9 categories

## 0.4.0

- **Dynamic array functions** (17): SEQUENCE, RANDARRAY, TOCOL, TOROW, WRAPROWS, WRAPCOLS, CHOOSEROWS, CHOOSECOLS, DROP, TAKE, EXPAND, HSTACK, VSTACK, FILTER, UNIQUE, SORT, SORTBY
- Array generators, flatten/reshape, slice/select, concatenation, and filter/sort/unique operations
- Private shared helpers for matrix conversion, flattening, value comparison, and index resolution
- 199 total built-in functions across 8 categories

## 0.3.0

- **87 new built-in functions** across all existing categories (95 → 182 total)
- **Math & Trigonometry** (28 new): PI, LN, LOG, LOG10, EXP, SIN, COS, TAN, ASIN, ACOS, ATAN, ATAN2, DEGREES, RADIANS, EVEN, ODD, GCD, LCM, TRUNC, MROUND, QUOTIENT, COMBIN, COMBINA, FACT, FACTDOUBLE, SUMSQ, SUBTOTAL, AGGREGATE
- **Text** (13 new): REPT, CHAR, CODE, CLEAN, DOLLAR, FIXED, T, NUMBERVALUE, UNICHAR, UNICODE, TEXTBEFORE, TEXTAFTER, TEXTSPLIT
- **Statistical** (19 new): STDEV.S, STDEV.P, VAR.S, VAR.P, PERCENTILE.INC, PERCENTILE.EXC, PERCENTRANK.INC, PERCENTRANK.EXC, RANK.AVG, FREQUENCY, AVEDEV, AVERAGEA, MAXA, MINA, TRIMMEAN, GEOMEAN, HARMEAN, MAXIFS, MINIFS
- **Lookup & Reference** (10 new): ROW, COLUMN, ROWS, COLUMNS, ADDRESS, INDIRECT, OFFSET, TRANSPOSE, HYPERLINK, AREAS
- **Date/Time** (9 new): TIMEVALUE, WEEKNUM, ISOWEEKNUM, NETWORKDAYS, NETWORKDAYS.INTL, WORKDAY, WORKDAY.INTL, DAYS360, YEARFRAC
- **Information** (8 new): ISERR, ISNONTEXT, ISEVEN, ISODD, ISREF, N, NA, ERROR.TYPE

## 0.2.0

- **Statistical functions**: COUNT, COUNTA, COUNTBLANK, COUNTIF, SUMIF, AVERAGEIF with shared criteria matching
- **Lookup functions**: VLOOKUP (exact & approximate), INDEX, MATCH (exact, ascending, descending)
- **Date functions**: DATE, TODAY, NOW, YEAR, MONTH, DAY using Excel serial number convention
- **TEXT format codes**: `0`, `#`, decimal, thousands separator, percentage, scientific notation
- **Better parser errors**: Position-aware messages for unmatched parentheses, unexpected tokens, truncated input
- **Performance benchmarks**: Parse, evaluate, and dependency graph benchmarks
- **Iterative graph traversal**: DependencyGraph uses iterative DFS to handle deep cell chains
- **Worksheet example**: Runnable example demonstrating full API usage
- **Flutter integration example**: Standalone Flutter app integrating with the `worksheet` widget
- **MIT license**

## 0.1.0

- Initial version with core formula engine
- Formula parser with operator precedence (arithmetic, comparison, concatenation, percent, power)
- AST representation with sealed FormulaNode classes
- FormulaValue type system: Number, Text, Boolean, Error, Empty, Range
- Excel-compatible error types: #DIV/0!, #VALUE!, #REF!, #NAME?, #NUM!, #N/A, #NULL!, #CALC!, #CIRCULAR!
- Math functions: SUM, AVERAGE, MIN, MAX, ABS, ROUND, INT, MOD, SQRT, POWER
- Logical functions: IF, AND, OR, NOT, IFERROR, IFNA, TRUE, FALSE
- Text functions: CONCAT, CONCATENATE, LEFT, RIGHT, MID, LEN, LOWER, UPPER, TRIM, TEXT
- EvaluationContext interface for pluggable data sources
- FunctionRegistry with custom function registration
- DependencyGraph for cell dependency tracking and recalculation ordering
- Parse caching for performance
- Cell reference extraction from formulas
