# Dynamic Array Functions

## SEQUENCE

`SEQUENCE(rows, [cols], [start], [step])`

Generates a sequence of numbers in an array.

**Example:**
```
=SEQUENCE(3)    → {1;2;3}
```

## RANDARRAY

`RANDARRAY([rows], [cols], [min], [max], [whole])`

Generates an array of random numbers.

**Example:**
```
=RANDARRAY(2, 3, 1, 10, TRUE)    → {4,8,2;7,1,9}
```

## TOCOL

`TOCOL(array, [ignore], [scan_by_col])`

Flattens a 2D array into a single column.

**Example:**
```
=TOCOL({1,2;3,4})    → {1;2;3;4}
```

## TOROW

`TOROW(array, [ignore], [scan_by_col])`

Flattens a 2D array into a single row.

**Example:**
```
=TOROW({1,2;3,4})    → {1,2,3,4}
```

## WRAPROWS

`WRAPROWS(vector, wrap_count, [pad])`

Wraps a row or column vector into rows after a specified number of elements.

**Example:**
```
=WRAPROWS({1,2,3,4,5,6}, 3)    → {1,2,3;4,5,6}
```

## WRAPCOLS

`WRAPCOLS(vector, wrap_count, [pad])`

Wraps a row or column vector into columns after a specified number of elements.

**Example:**
```
=WRAPCOLS({1,2,3,4,5,6}, 3)    → {1,4;2,5;3,6}
```

## CHOOSEROWS

`CHOOSEROWS(array, row_num1, ...)`

Returns the specified rows from an array.

**Example:**
```
=CHOOSEROWS({1,2;3,4;5,6}, 1, 3)    → {1,2;5,6}
```

## CHOOSECOLS

`CHOOSECOLS(array, col_num1, ...)`

Returns the specified columns from an array.

**Example:**
```
=CHOOSECOLS({1,2,3;4,5,6}, 1, 3)    → {1,3;4,6}
```

## DROP

`DROP(array, rows, [cols])`

Drops a specified number of rows and/or columns from an array.

**Example:**
```
=DROP({1,2;3,4;5,6}, 1)    → {3,4;5,6}
```

## TAKE

`TAKE(array, rows, [cols])`

Takes a specified number of rows and/or columns from an array.

**Example:**
```
=TAKE({1,2;3,4;5,6}, 2)    → {1,2;3,4}
```

## EXPAND

`EXPAND(array, rows, [cols], [pad])`

Expands an array to the specified dimensions, padding with a value.

**Example:**
```
=EXPAND({1,2;3,4}, 3, 3, 0)    → {1,2,0;3,4,0;0,0,0}
```

## HSTACK

`HSTACK(array1, ...)`

Horizontally stacks arrays side by side.

**Example:**
```
=HSTACK({1;2}, {3;4})    → {1,3;2,4}
```

## VSTACK

`VSTACK(array1, ...)`

Vertically stacks arrays on top of each other.

**Example:**
```
=VSTACK({1,2}, {3,4})    → {1,2;3,4}
```

## FILTER

`FILTER(array, include, [if_empty])`

Filters an array based on a Boolean criteria array.

**Example:**
```
=FILTER({1;2;3;4}, {TRUE;FALSE;TRUE;FALSE})    → {1;3}
```

## UNIQUE

`UNIQUE(array, [by_col], [exactly_once])`

Returns the unique values from a range or array.

**Example:**
```
=UNIQUE({1;2;2;3;3;3})    → {1;2;3}
```

## SORT

`SORT(array, [sort_index], [sort_order], [by_col])`

Sorts the contents of a range or array.

**Example:**
```
=SORT({3;1;2})    → {1;2;3}
```

## SORTBY

`SORTBY(array, by_array1, [sort_order1], ...)`

Sorts the contents of a range or array based on the values in a corresponding array.

**Example:**
```
=SORTBY({"a";"b";"c"}, {3;1;2})    → {"b";"c";"a"}
```

## MUNIT

`MUNIT(dimension)`

Returns the identity matrix of size N x N.

**Example:**
```
=MUNIT(3)    → {1,0,0;0,1,0;0,0,1}
```

## MMULT

`MMULT(array1, array2)`

Matrix multiplication.

**Example:**
```
=MMULT({1,2;3,4}, {5;6})    → {17;39}
```

## MDETERM

`MDETERM(array)`

Returns the determinant of a square matrix.

**Example:**
```
=MDETERM({1,2;3,4})    → -2
```

## MINVERSE

`MINVERSE(array)`

Returns the inverse of a square matrix.

**Example:**
```
=MINVERSE({4,7;2,6})    → {0.6,-0.7;-0.2,0.4}
```
