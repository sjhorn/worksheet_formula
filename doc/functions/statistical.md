# Statistical Functions

## COUNT

`COUNT(value1, [value2], ...)`

Counts numeric values.

**Example:**
```
=COUNT(1, 2, "text", 4)    → 3
```

## COUNTA

`COUNTA(value1, [value2], ...)`

Counts non-empty values.

**Example:**
```
=COUNTA(1, "hello", TRUE, "", 5)    → 4
```

## COUNTBLANK

`COUNTBLANK(range)`

Counts empty cells in a range.

**Example:**
```
=COUNTBLANK(A1:A5)    → 2
```

## COUNTIF

`COUNTIF(range, criteria)`

Counts cells matching criteria.

**Example:**
```
=COUNTIF(A1:A5, ">10")    → 3
```

## SUMIF

`SUMIF(range, criteria, [sum_range])`

Sums cells matching criteria.

**Example:**
```
=SUMIF(A1:A5, ">10", B1:B5)    → 75
```

## AVERAGEIF

`AVERAGEIF(range, criteria, [average_range])`

Averages cells matching criteria.

**Example:**
```
=AVERAGEIF(A1:A5, ">10", B1:B5)    → 25
```

## SUMIFS

`SUMIFS(sum_range, criteria_range1, criteria1, ...)`

Sum with multiple criteria.

**Example:**
```
=SUMIFS(C1:C5, A1:A5, ">10", B1:B5, "<100")    → 50
```

## COUNTIFS

`COUNTIFS(criteria_range1, criteria1, ...)`

Count with multiple criteria.

**Example:**
```
=COUNTIFS(A1:A5, ">10", B1:B5, "<100")    → 2
```

## AVERAGEIFS

`AVERAGEIFS(average_range, criteria_range1, criteria1, ...)`

Average with multiple criteria.

**Example:**
```
=AVERAGEIFS(C1:C5, A1:A5, ">10", B1:B5, "<100")    → 25
```

## MEDIAN

`MEDIAN(number1, [number2], ...)`

Returns the middle value.

**Example:**
```
=MEDIAN(1, 5, 3, 7, 9)    → 5
```

## MODE.SNGL

`MODE.SNGL(number1, [number2], ...)`

Returns the most frequent value.

**Example:**
```
=MODE.SNGL(1, 2, 2, 3, 3, 3)    → 3
```

## MODE

`MODE`

Alias for MODE.SNGL.

**Example:**
```
=MODE(1, 2, 2, 3, 3, 3)    → 3
```

## LARGE

`LARGE(array, k)`

Returns the k-th largest value.

**Example:**
```
=LARGE({3,5,1,8,2}, 2)    → 5
```

## SMALL

`SMALL(array, k)`

Returns the k-th smallest value.

**Example:**
```
=SMALL({3,5,1,8,2}, 2)    → 2
```

## RANK.EQ

`RANK.EQ(number, ref, [order])`

Rank of a number in a list.

**Example:**
```
=RANK.EQ(5, {1,3,5,7,9})    → 3
```

## RANK

`RANK`

Alias for RANK.EQ.

**Example:**
```
=RANK(5, {1,3,5,7,9})    → 3
```

## STDEV.S

`STDEV.S(number1, [number2], ...)`

Sample standard deviation.

**Example:**
```
=STDEV.S(2, 4, 4, 4, 5, 5, 7, 9)    → 2.138
```

## STDEV.P

`STDEV.P(number1, [number2], ...)`

Population standard deviation.

**Example:**
```
=STDEV.P(2, 4, 4, 4, 5, 5, 7, 9)    → 2
```

## VAR.S

`VAR.S(number1, [number2], ...)`

Sample variance.

**Example:**
```
=VAR.S(2, 4, 4, 4, 5, 5, 7, 9)    → 4.571
```

## VAR.P

`VAR.P(number1, [number2], ...)`

Population variance.

**Example:**
```
=VAR.P(2, 4, 4, 4, 5, 5, 7, 9)    → 4
```

## PERCENTILE.INC

`PERCENTILE.INC(array, k)`

Returns the k-th percentile (inclusive).

**Example:**
```
=PERCENTILE.INC({1,2,3,4,5}, 0.4)    → 2.6
```

## PERCENTILE.EXC

`PERCENTILE.EXC(array, k)`

Returns the k-th percentile (exclusive).

**Example:**
```
=PERCENTILE.EXC({1,2,3,4,5}, 0.5)    → 3
```

## PERCENTRANK.INC

`PERCENTRANK.INC(array, x, [significance])`

Returns percent rank (inclusive).

**Example:**
```
=PERCENTRANK.INC({1,2,3,4,5}, 3)    → 0.5
```

## PERCENTRANK.EXC

`PERCENTRANK.EXC(array, x, [significance])`

Returns percent rank (exclusive).

**Example:**
```
=PERCENTRANK.EXC({1,2,3,4,5}, 3)    → 0.5
```

## RANK.AVG

`RANK.AVG(number, ref, [order])`

Rank with averaged ties.

**Example:**
```
=RANK.AVG(3, {3,3,5,7,9})    → 4.5
```

## FREQUENCY

`FREQUENCY(data_array, bins_array)`

Returns frequency distribution.

**Example:**
```
=FREQUENCY({1,2,3,4,5,6}, {2,4})    → {2;2;2}
```

## AVEDEV

`AVEDEV(number1, [number2], ...)`

Average absolute deviation.

**Example:**
```
=AVEDEV(4, 5, 6, 7, 8)    → 1.2
```

## AVERAGEA

`AVERAGEA(value1, [value2], ...)`

Average including text and logical values.

**Example:**
```
=AVERAGEA(1, TRUE, "text", 4)    → 1.5
```

## MAXA

`MAXA(value1, [value2], ...)`

Max including text and logical values.

**Example:**
```
=MAXA(1, TRUE, "text", 4)    → 4
```

## MINA

`MINA(value1, [value2], ...)`

Min including text and logical values.

**Example:**
```
=MINA(1, TRUE, "text", 4)    → 0
```

## TRIMMEAN

`TRIMMEAN(array, percent)`

Mean excluding outliers.

**Example:**
```
=TRIMMEAN({1,2,3,4,5,6,7,8}, 0.2)    → 4.5
```

## GEOMEAN

`GEOMEAN(number1, [number2], ...)`

Geometric mean.

**Example:**
```
=GEOMEAN(4, 9)    → 6
```

## HARMEAN

`HARMEAN(number1, [number2], ...)`

Harmonic mean.

**Example:**
```
=HARMEAN(2, 3, 6)    → 3
```

## MAXIFS

`MAXIFS(max_range, criteria_range1, criteria1, ...)`

Max with criteria.

**Example:**
```
=MAXIFS(B1:B5, A1:A5, ">10")    → 90
```

## MINIFS

`MINIFS(min_range, criteria_range1, criteria1, ...)`

Min with criteria.

**Example:**
```
=MINIFS(B1:B5, A1:A5, ">10")    → 20
```
