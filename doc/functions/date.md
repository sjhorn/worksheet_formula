# Date/Time Functions

## DATE

`DATE(year, month, day)`

Creates a date serial number.

**Example:**
```
=DATE(2024, 1, 15)    → 45306
```

## TODAY

`TODAY()`

Returns the current date as a serial number.

**Example:**
```
=TODAY()    → 45658
```

## NOW

`NOW()`

Returns the current date and time as a serial number.

**Example:**
```
=NOW()    → 45658.75
```

## YEAR

`YEAR(serial_number)`

Extracts the year from a date serial number.

**Example:**
```
=YEAR(45306)    → 2024
```

## MONTH

`MONTH(serial_number)`

Extracts the month from a date serial number.

**Example:**
```
=MONTH(45306)    → 1
```

## DAY

`DAY(serial_number)`

Extracts the day from a date serial number.

**Example:**
```
=DAY(45306)    → 15
```

## DAYS

`DAYS(end_date, start_date)`

Days between two dates.

**Example:**
```
=DAYS(45366, 45306)    → 60
```

## DATEDIF

`DATEDIF(start_date, end_date, unit)`

Difference in various units.

**Example:**
```
=DATEDIF(DATE(2020,1,1), DATE(2024,6,15), "Y")    → 4
```

## DATEVALUE

`DATEVALUE(date_text)`

Converts date text to serial number.

**Example:**
```
=DATEVALUE("2024-01-15")    → 45306
```

## WEEKDAY

`WEEKDAY(serial_number, [return_type])`

Day of the week.

**Example:**
```
=WEEKDAY(DATE(2024,1,15))    → 2
```

## HOUR

`HOUR(serial_number)`

Extract hour from time.

**Example:**
```
=HOUR(0.75)    → 18
```

## MINUTE

`MINUTE(serial_number)`

Extract minute from time.

**Example:**
```
=MINUTE(0.75)    → 0
```

## SECOND

`SECOND(serial_number)`

Extract second from time.

**Example:**
```
=SECOND(0.5007)    → 1
```

## TIME

`TIME(hour, minute, second)`

Construct time as fraction of day.

**Example:**
```
=TIME(12, 0, 0)    → 0.5
```

## EDATE

`EDATE(start_date, months)`

Add months to a date.

**Example:**
```
=EDATE(DATE(2024,1,15), 3)    → 45367
```

## EOMONTH

`EOMONTH(start_date, months)`

End of month after adding months.

**Example:**
```
=EOMONTH(DATE(2024,1,15), 0)    → 45322
```

## TIMEVALUE

`TIMEVALUE(time_text)`

Converts a time text string to a fractional day number.

**Example:**
```
=TIMEVALUE("12:00:00")    → 0.5
```

## WEEKNUM

`WEEKNUM(serial_number, [return_type])`

Returns the week number of a date.

**Example:**
```
=WEEKNUM(DATE(2024,3,15))    → 11
```

## ISOWEEKNUM

`ISOWEEKNUM(date)`

Returns the ISO 8601 week number of a date.

**Example:**
```
=ISOWEEKNUM(DATE(2024,1,1))    → 1
```

## NETWORKDAYS

`NETWORKDAYS(start_date, end_date, [holidays])`

Returns the number of working days.

**Example:**
```
=NETWORKDAYS(DATE(2024,1,1), DATE(2024,1,31))    → 23
```

## NETWORKDAYS.INTL

`NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])`

Returns the number of working days with custom weekend settings.

**Example:**
```
=NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,31), 11)    → 26
```

## WORKDAY

`WORKDAY(start_date, days, [holidays])`

Returns the date after N working days.

**Example:**
```
=WORKDAY(DATE(2024,1,1), 10)    → 45313
```

## WORKDAY.INTL

`WORKDAY.INTL(start_date, days, [weekend], [holidays])`

Returns the date after N working days with custom weekend settings.

**Example:**
```
=WORKDAY.INTL(DATE(2024,1,1), 10, 11)    → 45311
```

## DAYS360

`DAYS360(start_date, end_date, [method])`

Days between dates on 30/360 basis.

**Example:**
```
=DAYS360(DATE(2024,1,1), DATE(2024,7,1))    → 180
```

## YEARFRAC

`YEARFRAC(start_date, end_date, [basis])`

Fraction of a year between two dates.

**Example:**
```
=YEARFRAC(DATE(2024,1,1), DATE(2024,7,1))    → 0.5
```
