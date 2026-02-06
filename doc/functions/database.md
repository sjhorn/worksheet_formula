# Database Functions

## DSUM

`DSUM(database, field, criteria)`

Sum of matching field values.

**Example:**
```
=DSUM(A1:C10, "Sales", E1:E2)    → 4500
```

## DAVERAGE

`DAVERAGE(database, field, criteria)`

Average of matching field values.

**Example:**
```
=DAVERAGE(A1:C10, "Score", E1:E2)    → 82.5
```

## DCOUNT

`DCOUNT(database, field, criteria)`

Count numeric cells in matching rows.

**Example:**
```
=DCOUNT(A1:C10, "Age", E1:E2)    → 3
```

## DCOUNTA

`DCOUNTA(database, field, criteria)`

Count non-empty cells in matching rows.

**Example:**
```
=DCOUNTA(A1:C10, "Name", E1:E2)    → 5
```

## DMAX

`DMAX(database, field, criteria)`

Maximum of matching numeric field values.

**Example:**
```
=DMAX(A1:C10, "Price", E1:E2)    → 299.99
```

## DMIN

`DMIN(database, field, criteria)`

Minimum of matching numeric field values.

**Example:**
```
=DMIN(A1:C10, "Price", E1:E2)    → 9.99
```

## DGET

`DGET(database, field, criteria)`

Return value from single matching row.

**Example:**
```
=DGET(A1:C10, "Email", E1:E2)    → "alice@example.com"
```

## DPRODUCT

`DPRODUCT(database, field, criteria)`

Product of matching numeric values.

**Example:**
```
=DPRODUCT(A1:C10, "Quantity", E1:E2)    → 120
```

## DSTDEV

`DSTDEV(database, field, criteria)`

Sample standard deviation.

**Example:**
```
=DSTDEV(A1:C10, "Score", E1:E2)    → 12.73
```

## DSTDEVP

`DSTDEVP(database, field, criteria)`

Population standard deviation.

**Example:**
```
=DSTDEVP(A1:C10, "Score", E1:E2)    → 11.02
```

## DVAR

`DVAR(database, field, criteria)`

Sample variance.

**Example:**
```
=DVAR(A1:C10, "Score", E1:E2)    → 162.05
```

## DVARP

`DVARP(database, field, criteria)`

Population variance.

**Example:**
```
=DVARP(A1:C10, "Score", E1:E2)    → 121.54
```
