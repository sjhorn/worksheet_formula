# Roadmap: Built-in Function Coverage

Goal: implement all commonly used Excel / Google Sheets built-in functions,
prioritised by real-world usage frequency.

## Current State

**400 functions** across 14 categories:

| Category | Count | Functions |
|----------|-------|-----------|
| Math & Trig | 50 | SUM, AVERAGE, MIN, MAX, ABS, ROUND, INT, MOD, SQRT, POWER, SUMPRODUCT, ROUNDUP, ROUNDDOWN, CEILING, FLOOR, SIGN, PRODUCT, RAND, RANDBETWEEN, PI, LN, LOG, LOG10, EXP, SIN, COS, TAN, ASIN, ACOS, ATAN, ATAN2, DEGREES, RADIANS, EVEN, ODD, GCD, LCM, TRUNC, MROUND, QUOTIENT, COMBIN, COMBINA, FACT, FACTDOUBLE, SUMSQ, SUBTOTAL, AGGREGATE, SERIESSUM, SQRTPI, MULTINOMIAL |
| Logical | 11 | IF, AND, OR, NOT, IFERROR, IFNA, TRUE, FALSE, IFS, SWITCH, XOR |
| Text | 36 | CONCAT, CONCATENATE, LEFT, RIGHT, MID, LEN, LOWER, UPPER, TRIM, TEXT, FIND, SEARCH, SUBSTITUTE, REPLACE, VALUE, TEXTJOIN, PROPER, EXACT, REPT, CHAR, CODE, CLEAN, DOLLAR, FIXED, T, NUMBERVALUE, UNICHAR, UNICODE, TEXTBEFORE, TEXTAFTER, TEXTSPLIT, ARRAYTOTEXT, VALUETOTEXT, ASC, DBCS, BAHTTEXT |
| Statistical | 35 | COUNT, COUNTA, COUNTBLANK, COUNTIF, SUMIF, AVERAGEIF, SUMIFS, COUNTIFS, AVERAGEIFS, MEDIAN, MODE.SNGL, MODE, LARGE, SMALL, RANK.EQ, RANK, STDEV.S, STDEV.P, VAR.S, VAR.P, PERCENTILE.INC, PERCENTILE.EXC, PERCENTRANK.INC, PERCENTRANK.EXC, RANK.AVG, FREQUENCY, AVEDEV, AVERAGEA, MAXA, MINA, TRIMMEAN, GEOMEAN, HARMEAN, MAXIFS, MINIFS |
| Lookup & Ref | 18 | VLOOKUP, INDEX, MATCH, HLOOKUP, LOOKUP, CHOOSE, XMATCH, XLOOKUP, ROW, COLUMN, ROWS, COLUMNS, ADDRESS, INDIRECT, OFFSET, TRANSPOSE, HYPERLINK, AREAS |
| Date/Time | 25 | DATE, TODAY, NOW, YEAR, MONTH, DAY, DAYS, DATEDIF, DATEVALUE, WEEKDAY, HOUR, MINUTE, SECOND, TIME, EDATE, EOMONTH, TIMEVALUE, WEEKNUM, ISOWEEKNUM, NETWORKDAYS, NETWORKDAYS.INTL, WORKDAY, WORKDAY.INTL, DAYS360, YEARFRAC |
| Information | 15 | ISBLANK, ISERROR, ISNUMBER, ISTEXT, ISLOGICAL, ISNA, TYPE, ISERR, ISNONTEXT, ISEVEN, ISODD, ISREF, N, NA, ERROR.TYPE |
| Dynamic Array | 21 | SEQUENCE, RANDARRAY, TOCOL, TOROW, WRAPROWS, WRAPCOLS, CHOOSEROWS, CHOOSECOLS, DROP, TAKE, EXPAND, HSTACK, VSTACK, FILTER, UNIQUE, SORT, SORTBY, MUNIT, MMULT, MDETERM, MINVERSE |
| Financial | 40 | PMT, FV, PV, NPER, RATE, IPMT, PPMT, CUMIPMT, CUMPRINC, NPV, XNPV, IRR, XIRR, MIRR, FVSCHEDULE, SLN, SYD, DB, DDB, VDB, PRICE, YIELD, DURATION, MDURATION, ACCRINT, DISC, INTRATE, RECEIVED, PRICEDISC, PRICEMAT, TBILLEQ, TBILLPRICE, TBILLYIELD, DOLLARDE, DOLLARFR, EFFECT, NOMINAL, PDURATION, RRI, ISPMT |
| Adv. Statistical | 70 | FISHER, FISHERINV, STANDARDIZE, PERMUT, PERMUTATIONA, DEVSQ, KURT, SKEW, SKEW.P, COVARIANCE.P, COVARIANCE.S, CORREL, PEARSON, RSQ, SLOPE, INTERCEPT, STEYX, FORECAST.LINEAR, PROB, MODE.MULT, STDEVA, STDEVPA, VARA, VARPA, GAMMA, GAMMALN, GAMMALN.PRECISE, GAUSS, PHI, NORM.S.DIST, NORM.S.INV, NORM.DIST, NORM.INV, BINOM.DIST, BINOM.INV, BINOM.DIST.RANGE, NEGBINOM.DIST, HYPGEOM.DIST, POISSON.DIST, EXPON.DIST, GAMMA.DIST, GAMMA.INV, BETA.DIST, BETA.INV, CHISQ.DIST, CHISQ.INV, CHISQ.DIST.RT, CHISQ.INV.RT, T.DIST, T.INV, T.DIST.2T, T.INV.2T, T.DIST.RT, F.DIST, F.INV, F.DIST.RT, F.INV.RT, WEIBULL.DIST, LOGNORM.DIST, LOGNORM.INV, CONFIDENCE.NORM, CONFIDENCE.T, Z.TEST, T.TEST, CHISQ.TEST, F.TEST, LINEST, LOGEST, TREND, GROWTH |
| Engineering | 54 | DELTA, GESTEP, BITAND, BITOR, BITXOR, BITLSHIFT, BITRSHIFT, BIN2DEC, BIN2HEX, BIN2OCT, DEC2BIN, DEC2HEX, DEC2OCT, HEX2BIN, HEX2DEC, HEX2OCT, OCT2BIN, OCT2DEC, OCT2HEX, BASE, DECIMAL, ARABIC, ROMAN, ERF, ERF.PRECISE, ERFC, ERFC.PRECISE, COMPLEX, IMREAL, IMAGINARY, IMABS, IMARGUMENT, IMCONJUGATE, IMSUM, IMSUB, IMPRODUCT, IMDIV, IMPOWER, IMSQRT, IMEXP, IMLN, IMLOG10, IMLOG2, IMSIN, IMCOS, IMTAN, IMSINH, IMCOSH, IMSEC, IMSECH, IMCSC, IMCSCH, IMCOT, CONVERT |
| Database | 12 | DSUM, DAVERAGE, DCOUNT, DCOUNTA, DMAX, DMIN, DGET, DPRODUCT, DSTDEV, DSTDEVP, DVAR, DVARP |
| Lambda & Higher-Order | 9 | LAMBDA, LET, MAP, REDUCE, SCAN, MAKEARRAY, BYCOL, BYROW, ISOMITTED |
| Web & Regex | 4 | ENCODEURL, REGEXMATCH, REGEXEXTRACT, REGEXREPLACE |

---

## Phase 1 — High-Priority Missing Functions ✅ Complete

All 52 functions implemented and tested (527 tests passing).

### Multi-Criteria Functions (extremely high demand)

| Function | Description |
|----------|-------------|
| SUMIFS | Sum with multiple criteria |
| COUNTIFS | Count with multiple criteria |
| AVERAGEIFS | Average with multiple criteria |

### Lookup & Reference

| Function | Description |
|----------|-------------|
| XLOOKUP | Modern replacement for VLOOKUP — searches any direction |
| HLOOKUP | Horizontal lookup |
| LOOKUP | Simplified lookup (vector/array form) |
| CHOOSE | Return value from a list by index |
| XMATCH | Modern MATCH with additional match modes |

### Logical

| Function | Description |
|----------|-------------|
| IFS | Chained IF without nesting |
| SWITCH | Match a value against a list of cases |
| XOR | Exclusive OR |

### Text

| Function | Description |
|----------|-------------|
| FIND | Case-sensitive position search |
| SEARCH | Case-insensitive position search (supports wildcards) |
| SUBSTITUTE | Replace occurrences of a substring |
| REPLACE | Replace characters by position |
| VALUE | Convert text to number |
| TEXTJOIN | Join text with delimiter (handles empties) |
| PROPER | Capitalize first letter of each word |
| EXACT | Case-sensitive string comparison |

### Math

| Function | Description |
|----------|-------------|
| SUMPRODUCT | Sum of element-wise products |
| ROUNDUP | Always round away from zero |
| ROUNDDOWN | Always round toward zero |
| CEILING | Round up to nearest multiple |
| FLOOR | Round down to nearest multiple |
| SIGN | Returns -1, 0, or 1 |
| PRODUCT | Multiply all arguments |
| RAND | Random number 0–1 |
| RANDBETWEEN | Random integer in range |

### Statistical

| Function | Description |
|----------|-------------|
| MEDIAN | Middle value |
| MODE / MODE.SNGL | Most frequent value |
| LARGE | K-th largest value |
| SMALL | K-th smallest value |
| RANK.EQ | Rank of a number in a list |

### Date/Time

| Function | Description |
|----------|-------------|
| DAYS | Days between two dates |
| DATEDIF | Difference in years/months/days |
| DATEVALUE | Convert date text to serial number |
| WEEKDAY | Day of week (1–7) |
| HOUR | Hour from a time |
| MINUTE | Minute from a time |
| SECOND | Second from a time |
| TIME | Construct time from h/m/s |
| EDATE | Add months to a date |
| EOMONTH | End of month after adding months |

### Information

| Function | Description |
|----------|-------------|
| ISBLANK | Cell is empty |
| ISERROR | Value is any error |
| ISNUMBER | Value is a number |
| ISTEXT | Value is text |
| ISLOGICAL | Value is boolean |
| ISNA | Value is #N/A |
| TYPE | Returns type code of a value |

**Phase 1 total: 52 functions → brought library from 43 to 95 functions**

---

## Phase 2 — Extended Essentials ✅ Complete

87 functions implemented and tested (768 tests passing).
Some context-dependent functions deferred (FORMULATEXT, ISFORMULA, CELL, SHEET, SHEETS).

### Math & Trigonometry

| Function | Description |
|----------|-------------|
| PI | The constant pi |
| LN | Natural logarithm |
| LOG | Logarithm with specified base |
| LOG10 | Base-10 logarithm |
| EXP | e raised to a power |
| SIN | Sine |
| COS | Cosine |
| TAN | Tangent |
| ASIN | Arcsine |
| ACOS | Arccosine |
| ATAN | Arctangent |
| ATAN2 | Arctangent from x and y |
| DEGREES | Radians to degrees |
| RADIANS | Degrees to radians |
| EVEN | Round up to nearest even |
| ODD | Round up to nearest odd |
| GCD | Greatest common divisor |
| LCM | Least common multiple |
| TRUNC | Truncate to integer |
| MROUND | Round to nearest multiple |
| QUOTIENT | Integer division |
| COMBIN | Combinations |
| COMBINA | Combinations with repetitions |
| FACT | Factorial |
| FACTDOUBLE | Double factorial |
| SUMSQ | Sum of squares |
| SUBTOTAL | Aggregate function with function selector |
| AGGREGATE | Extended SUBTOTAL with error-skipping options |

### Text

| Function | Description |
|----------|-------------|
| REPT | Repeat text N times |
| CHAR | Character from code |
| CODE | Code from character |
| CLEAN | Remove non-printable characters |
| DOLLAR | Format number as currency text |
| FIXED | Format number with fixed decimals |
| T | Return text if text, else empty |
| NUMBERVALUE | Locale-aware text-to-number |
| UNICHAR | Character from Unicode code point |
| UNICODE | Unicode code point from character |
| TEXTBEFORE | Text before a delimiter |
| TEXTAFTER | Text after a delimiter |
| TEXTSPLIT | Split text by delimiters into array |

### Statistical

| Function | Description |
|----------|-------------|
| STDEV.S | Sample standard deviation |
| STDEV.P | Population standard deviation |
| VAR.S | Sample variance |
| VAR.P | Population variance |
| PERCENTILE.INC | K-th percentile (inclusive) |
| PERCENTILE.EXC | K-th percentile (exclusive) |
| PERCENTRANK.INC | Percent rank (inclusive) |
| PERCENTRANK.EXC | Percent rank (exclusive) |
| RANK.AVG | Rank with averaged ties |
| FREQUENCY | Frequency distribution |
| AVEDEV | Average absolute deviation |
| AVERAGEA | Average including text/logical |
| MAXA | Max including text/logical |
| MINA | Min including text/logical |
| TRIMMEAN | Mean excluding outliers |
| GEOMEAN | Geometric mean |
| HARMEAN | Harmonic mean |
| MAXIFS | Max with criteria |
| MINIFS | Min with criteria |

### Date/Time

| Function | Description |
|----------|-------------|
| TIMEVALUE | Convert time text to serial number |
| WEEKNUM | Week number of the year |
| ISOWEEKNUM | ISO week number |
| NETWORKDAYS | Working days between two dates |
| NETWORKDAYS.INTL | Working days (custom weekends) |
| WORKDAY | Date after N working days |
| WORKDAY.INTL | Workday (custom weekends) |
| DAYS360 | Days between dates (30/360 basis) |
| YEARFRAC | Fraction of year between dates |

### Lookup & Reference

| Function | Description |
|----------|-------------|
| ROW | Row number of a reference |
| ROWS | Number of rows in a reference |
| COLUMN | Column number of a reference |
| COLUMNS | Number of columns in a reference |
| ADDRESS | Create a cell reference as text |
| INDIRECT | Convert text to a reference |
| OFFSET | Reference offset from a cell |
| TRANSPOSE | Transpose rows/columns |
| FORMULATEXT | Return formula as text |
| HYPERLINK | Create a clickable link |
| AREAS | Number of areas in a reference |

### Information

| Function | Description |
|----------|-------------|
| ISERR | Is error (excluding #N/A) |
| ISNONTEXT | Is not text |
| ISEVEN | Is even number |
| ISODD | Is odd number |
| ISREF | Is a valid reference |
| ISFORMULA | Cell contains a formula |
| N | Convert to number |
| NA | Return #N/A error |
| ERROR.TYPE | Error type number |
| CELL | Information about a cell |
| SHEET | Sheet number |
| SHEETS | Number of sheets |

**Phase 2 total: 87 functions implemented (6 deferred) → brought library from 95 to 182 functions**

---

## Phase 3 — Dynamic Array Functions ✅ Complete

All 17 functions implemented and tested (857 tests passing).
New category file: `lib/src/functions/array.dart`.

| Function | Description |
|----------|-------------|
| FILTER | Filter rows/columns by criteria |
| SORT | Sort a range |
| SORTBY | Sort by another range |
| UNIQUE | Remove duplicates |
| SEQUENCE | Generate a sequence of numbers |
| RANDARRAY | Generate array of random numbers |
| WRAPCOLS | Wrap a row into columns |
| WRAPROWS | Wrap a column into rows |
| TOCOL | Flatten to a single column |
| TOROW | Flatten to a single row |
| CHOOSECOLS | Select columns from an array |
| CHOOSEROWS | Select rows from an array |
| DROP | Remove rows/columns from edges |
| TAKE | Take rows/columns from edges |
| EXPAND | Expand array to specified dimensions |
| HSTACK | Stack arrays horizontally |
| VSTACK | Stack arrays vertically |

**Phase 3 total: 17 functions → brought library from 182 to 199 functions**

---

## Phase 4 — Financial Functions ✅ Complete

All 40 functions implemented and tested (960 tests passing).
New category file: `lib/src/functions/financial.dart`.

### Loans & Annuities

| Function | Description |
|----------|-------------|
| PMT | Payment for a loan |
| IPMT | Interest portion of a payment |
| PPMT | Principal portion of a payment |
| PV | Present value |
| FV | Future value |
| NPER | Number of periods |
| RATE | Interest rate per period |
| CUMIPMT | Cumulative interest over a range of periods |
| CUMPRINC | Cumulative principal over a range of periods |

### Investment Analysis

| Function | Description |
|----------|-------------|
| NPV | Net present value |
| XNPV | NPV for irregular cash flows |
| IRR | Internal rate of return |
| XIRR | IRR for irregular cash flows |
| MIRR | Modified IRR |
| FVSCHEDULE | Future value with variable rates |

### Depreciation

| Function | Description |
|----------|-------------|
| SLN | Straight-line depreciation |
| SYD | Sum-of-years'-digits depreciation |
| DB | Fixed-declining balance |
| DDB | Double-declining balance |
| VDB | Variable declining balance |

### Bonds & Securities

| Function | Description |
|----------|-------------|
| PRICE | Bond price |
| YIELD | Bond yield |
| DURATION | Macaulay duration |
| MDURATION | Modified duration |
| ACCRINT | Accrued interest |
| DISC | Discount rate for a security |
| INTRATE | Interest rate for a fully invested security |
| RECEIVED | Amount received at maturity |
| PRICEDISC | Price of a discounted security |
| PRICEMAT | Price of a security that pays at maturity |
| TBILLEQ | T-bill equivalent yield |
| TBILLPRICE | T-bill price |
| TBILLYIELD | T-bill yield |
| DOLLARDE | Dollar fraction to decimal |
| DOLLARFR | Dollar decimal to fraction |

### Other Financial

| Function | Description |
|----------|-------------|
| EFFECT | Effective annual interest rate |
| NOMINAL | Nominal annual interest rate |
| PDURATION | Periods to reach target value |
| RRI | Equivalent interest rate for growth |
| ISPMT | Interest on straight-line loan |

**Phase 4 total: 40 functions → brought library from 199 to 239 functions**

---

## Phase 5 — Advanced Statistical & Probability ✅ Complete

All 70 functions implemented and tested (1173 tests passing).
New category file: `lib/src/functions/statistical_advanced.dart`.

### Distributions

| Function | Description |
|----------|-------------|
| NORM.DIST | Normal distribution |
| NORM.INV | Inverse normal distribution |
| NORM.S.DIST | Standard normal distribution |
| NORM.S.INV | Inverse standard normal |
| T.DIST | Student's t-distribution |
| T.INV | Inverse t-distribution |
| T.DIST.2T | Two-tailed t-distribution |
| T.INV.2T | Two-tailed inverse t-distribution |
| T.DIST.RT | Right-tailed t-distribution |
| CHISQ.DIST | Chi-squared distribution |
| CHISQ.INV | Inverse chi-squared |
| CHISQ.DIST.RT | Right-tailed chi-squared |
| CHISQ.INV.RT | Right-tailed inverse chi-squared |
| F.DIST | F-distribution |
| F.INV | Inverse F-distribution |
| F.DIST.RT | Right-tailed F-distribution |
| F.INV.RT | Right-tailed inverse F |
| BINOM.DIST | Binomial distribution |
| BINOM.INV | Inverse binomial |
| BINOM.DIST.RANGE | Binomial probability in a range |
| POISSON.DIST | Poisson distribution |
| EXPON.DIST | Exponential distribution |
| GAMMA.DIST | Gamma distribution |
| GAMMA.INV | Inverse gamma |
| BETA.DIST | Beta distribution |
| BETA.INV | Inverse beta |
| WEIBULL.DIST | Weibull distribution |
| LOGNORM.DIST | Lognormal distribution |
| LOGNORM.INV | Inverse lognormal |
| NEGBINOM.DIST | Negative binomial |
| HYPGEOM.DIST | Hypergeometric distribution |

### Regression & Correlation

| Function | Description |
|----------|-------------|
| CORREL | Correlation coefficient |
| PEARSON | Pearson correlation |
| RSQ | R-squared |
| SLOPE | Slope of linear regression |
| INTERCEPT | Y-intercept of linear regression |
| STEYX | Standard error of estimate |
| LINEST | Full linear regression statistics |
| LOGEST | Exponential regression statistics |
| TREND | Values along a linear trend |
| GROWTH | Values along an exponential trend |
| FORECAST.LINEAR | Forecast using linear regression |

### Hypothesis Testing

| Function | Description |
|----------|-------------|
| T.TEST | Student's t-test |
| CHISQ.TEST | Chi-squared test |
| F.TEST | F-test |
| Z.TEST | Z-test |
| CONFIDENCE.NORM | Confidence interval (normal) |
| CONFIDENCE.T | Confidence interval (t-distribution) |

### Other Statistical

| Function | Description |
|----------|-------------|
| COVARIANCE.P | Population covariance |
| COVARIANCE.S | Sample covariance |
| FISHER | Fisher transformation |
| FISHERINV | Inverse Fisher |
| GAMMA | Gamma function value |
| GAMMALN | Log of gamma function |
| GAMMALN.PRECISE | Precise log gamma |
| GAUSS | Standard normal cumulative to z |
| PHI | Standard normal density at z |
| KURT | Kurtosis |
| SKEW | Skewness |
| SKEW.P | Population skewness |
| DEVSQ | Sum of squared deviations |
| PROB | Probability within range |
| PERMUT | Permutations |
| PERMUTATIONA | Permutations with repetition |
| STANDARDIZE | Z-score |
| MODE.MULT | Multiple modes |
| STDEVA | Stdev including text/logical |
| STDEVPA | Pop stdev including text/logical |
| VARA | Variance including text/logical |
| VARPA | Pop variance including text/logical |

**Phase 5 total: 70 functions → brought library from 239 to 309 functions**

---

## Phase 6 — Engineering Functions ✅ Complete

All 54 functions implemented and tested (1365 tests passing).
New category file: `lib/src/functions/engineering.dart`.

### Comparison & Bitwise

| Function | Description |
|----------|-------------|
| DELTA | Tests whether two values are equal |
| GESTEP | Tests whether a number >= step |
| BITAND | Bitwise AND |
| BITOR | Bitwise OR |
| BITXOR | Bitwise XOR |
| BITLSHIFT | Bitwise left shift |
| BITRSHIFT | Bitwise right shift |

### Base Conversion

| Function | Description |
|----------|-------------|
| BIN2DEC | Binary to decimal |
| BIN2HEX | Binary to hexadecimal |
| BIN2OCT | Binary to octal |
| DEC2BIN | Decimal to binary |
| DEC2HEX | Decimal to hexadecimal |
| DEC2OCT | Decimal to octal |
| HEX2BIN | Hexadecimal to binary |
| HEX2DEC | Hexadecimal to decimal |
| HEX2OCT | Hexadecimal to octal |
| OCT2BIN | Octal to binary |
| OCT2DEC | Octal to decimal |
| OCT2HEX | Octal to hexadecimal |

### Number Format & Error Functions

| Function | Description |
|----------|-------------|
| BASE | Convert number to text in given base |
| DECIMAL | Convert text in given base to number |
| ARABIC | Roman numeral to number |
| ROMAN | Number to Roman numeral |
| ERF | Error function |
| ERF.PRECISE | Precise error function |
| ERFC | Complementary error function |
| ERFC.PRECISE | Precise complementary error function |

### Complex Numbers

| Function | Description |
|----------|-------------|
| COMPLEX | Create complex number |
| IMREAL | Real part of complex number |
| IMAGINARY | Imaginary part |
| IMABS | Absolute value of complex number |
| IMARGUMENT | Argument (angle) |
| IMCONJUGATE | Complex conjugate |
| IMSUM | Sum of complex numbers |
| IMSUB | Subtract complex numbers |
| IMPRODUCT | Product of complex numbers |
| IMDIV | Divide complex numbers |
| IMPOWER | Complex power |
| IMSQRT | Complex square root |
| IMEXP | Complex exponential |
| IMLN | Complex natural log |
| IMLOG10 | Complex log base 10 |
| IMLOG2 | Complex log base 2 |
| IMSIN | Complex sine |
| IMCOS | Complex cosine |
| IMTAN | Complex tangent |
| IMSINH | Complex hyperbolic sine |
| IMCOSH | Complex hyperbolic cosine |
| IMSEC | Complex secant |
| IMSECH | Complex hyperbolic secant |
| IMCSC | Complex cosecant |
| IMCSCH | Complex hyperbolic cosecant |
| IMCOT | Complex cotangent |

### Unit Conversion

| Function | Description |
|----------|-------------|
| CONVERT | Unit conversion (13 categories, ~100 units, metric + binary prefixes) |

**Phase 6 total: 54 functions → brought library from 309 to 363 functions**

---

## Phase 7 — Database Functions ✅ Complete

All 12 functions implemented and tested (1447 tests passing).
New category file: `lib/src/functions/database.dart`.

| Function | Description |
|----------|-------------|
| DSUM | Sum matching rows |
| DAVERAGE | Average matching rows |
| DCOUNT | Count numeric matching rows |
| DCOUNTA | Count non-blank matching rows |
| DMAX | Max of matching rows |
| DMIN | Min of matching rows |
| DGET | Single value from matching row |
| DPRODUCT | Product of matching rows |
| DSTDEV | Sample stdev of matching rows |
| DSTDEVP | Population stdev of matching rows |
| DVAR | Sample variance of matching rows |
| DVARP | Population variance of matching rows |

**Phase 7 total: 12 functions → brought library from 363 to 375 functions**

---

## Phase 8 — Lambda & Higher-Order Functions ✅ Complete

All 9 functions implemented and tested (1545 tests passing).
New category file: `lib/src/functions/lambda.dart`.
Structural changes: added `NameNode` to AST, `FunctionValue`/`OmittedValue` to values,
`getVariable()` to `EvaluationContext`, and bare identifier parsing to the parser.

| Function | Description |
|----------|-------------|
| LAMBDA | Create a custom function (returns FunctionValue closure) |
| LET | Define named variables in a formula |
| MAP | Apply a lambda to each element |
| REDUCE | Reduce an array to a single value |
| SCAN | Running accumulation |
| MAKEARRAY | Build array from lambda (1-based indices) |
| BYCOL | Apply lambda to each column |
| BYROW | Apply lambda to each row |
| ISOMITTED | Test if lambda argument was omitted |

**Phase 8 total: 9 functions → brought library from 375 to 384 functions**

---

## Phase 9 — Remaining & Niche Functions ✅ Complete

16 functions implemented and tested (1625 tests passing).
New category file: `lib/src/functions/web.dart`.
2 functions deferred (GETPIVOTDATA — requires pivot table metadata; PHONETIC — requires IME/furigana mapping data).

### Web & Regex

| Function | Description |
|----------|-------------|
| ENCODEURL | URL-encode a string |
| REGEXMATCH | Test if text matches pattern |
| REGEXEXTRACT | Extract first match |
| REGEXREPLACE | Replace matches |

### Math

| Function | Description |
|----------|-------------|
| SERIESSUM | Power series sum |
| SQRTPI | Square root of n × π |
| MULTINOMIAL | Multinomial coefficient |

### Matrix

| Function | Description |
|----------|-------------|
| MUNIT | Identity matrix |
| MMULT | Matrix multiplication |
| MDETERM | Matrix determinant (LU decomposition) |
| MINVERSE | Matrix inverse (Gauss-Jordan elimination) |

### Text

| Function | Description |
|----------|-------------|
| ARRAYTOTEXT | Array to text representation |
| VALUETOTEXT | Value to text representation |
| ASC | Full-width to half-width (CJK) |
| DBCS | Half-width to full-width (CJK) |
| BAHTTEXT | Number to Thai Baht text |

**Phase 9 total: 16 functions (2 deferred) → brought library from 384 to 400 functions**

---

## Summary

| Phase | Focus | New | Running Total | Status |
|-------|-------|-----|---------------|--------|
| 1 | High-priority missing | 52 | 95 | Done |
| 2 | Extended essentials | 87 | 182 | Done |
| 3 | Dynamic arrays | 17 | 199 | Done |
| 4 | Financial | 40 | 239 | Done |
| 5 | Advanced statistics | 70 | 309 | Done |
| 6 | Engineering | 54 | 363 | Done |
| 7 | Database | 12 | 375 | Done |
| 8 | Lambda / higher-order | 9 | 384 | Done |
| 9 | Remaining & niche | 16 | 400 | Done |

All 9 phases are complete, covering 400 functions across 14 categories.
