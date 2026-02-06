# Lookup & Reference Functions

## VLOOKUP

`VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`

Looks for a value in the leftmost column and returns a value in the same row from a specified column.

**Example:**
```
=VLOOKUP("B", A1:C3, 3, FALSE)    → "Result"
```

## INDEX

`INDEX(array, row_num, [col_num])`

Returns a value from a specific row and column within a range.

**Example:**
```
=INDEX(A1:C3, 2, 3)    → 150
```

## MATCH

`MATCH(lookup_value, lookup_array, [match_type])`

Returns the relative position of a value in an array.

**Example:**
```
=MATCH(5, A1:A4, 0)    → 3
```

## HLOOKUP

`HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])`

Looks for a value in the top row and returns a value in the same column from a specified row.

**Example:**
```
=HLOOKUP("Age", A1:D2, 2, FALSE)    → 30
```

## LOOKUP

`LOOKUP(lookup_value, lookup_vector, [result_vector])`

Looks for a value in a one-column or one-row range, then returns a value from the same position in a second range.

**Example:**
```
=LOOKUP(5, A1:A5, B1:B5)    → "Five"
```

## CHOOSE

`CHOOSE(index_num, value1, [value2], ...)`

Returns value at index.

**Example:**
```
=CHOOSE(2, "Red", "Blue", "Green")    → "Blue"
```

## XMATCH

`XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])`

Returns the relative position of a value in an array, with support for exact, approximate, and wildcard matching.

**Example:**
```
=XMATCH("B", A1:A3, 0)    → 2
```

## XLOOKUP

`XLOOKUP(lookup, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])`

Searches a range for a match and returns the corresponding item from a second range.

**Example:**
```
=XLOOKUP("B", A1:A3, B1:B3, "Not found")    → "Beta"
```

## ROW

`ROW([reference])`

Returns the row number of a reference.

**Example:**
```
=ROW(B3)    → 3
```

## COLUMN

`COLUMN([reference])`

Returns the column number of a reference.

**Example:**
```
=COLUMN(C5)    → 3
```

## ROWS

`ROWS(reference)`

Returns the number of rows in a reference.

**Example:**
```
=ROWS(A1:A10)    → 10
```

## COLUMNS

`COLUMNS(reference)`

Returns the number of columns in a reference.

**Example:**
```
=COLUMNS(A1:D1)    → 4
```

## ADDRESS

`ADDRESS(row_num, col_num, [abs_num], [a1], [sheet_text])`

Creates a cell reference as text given row and column numbers.

**Example:**
```
=ADDRESS(2, 3)    → "$C$2"
```

## INDIRECT

`INDIRECT(ref_text, [a1])`

Returns the reference specified by a text string.

**Example:**
```
=INDIRECT("B2")    → 42
```

## OFFSET

`OFFSET(reference, rows, cols, [height], [width])`

Returns a reference offset from a starting reference.

**Example:**
```
=OFFSET(A1, 2, 1)    → B3_value
```

## TRANSPOSE

`TRANSPOSE(array)`

Swaps rows and columns of an array.

**Example:**
```
=TRANSPOSE(A1:C2)    → {A1,A2;B1,B2;C1,C2}
```

## HYPERLINK

`HYPERLINK(link_location, [friendly_name])`

Returns the friendly name or the URL as text.

**Example:**
```
=HYPERLINK("https://example.com", "Click here")    → "Click here"
```

## AREAS

`AREAS(reference)`

Returns the number of areas in a reference.

**Example:**
```
=AREAS(A1:C3)    → 1
```
