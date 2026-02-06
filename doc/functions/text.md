# Text Functions

## CONCAT

`CONCAT(text1, [text2], ...)`

Joins text strings.

**Example:**
```
=CONCAT("Hello", " ", "World")    → "Hello World"
```

## CONCATENATE

`CONCATENATE(text1, [text2], ...)`

Legacy version of CONCAT.

**Example:**
```
=CONCATENATE("First", " ", "Last")    → "First Last"
```

## LEFT

`LEFT(text, [num_chars])`

Returns leftmost characters.

**Example:**
```
=LEFT("Apple", 3)    → "App"
```

## RIGHT

`RIGHT(text, [num_chars])`

Returns rightmost characters.

**Example:**
```
=RIGHT("Apple", 3)    → "ple"
```

## MID

`MID(text, start_num, num_chars)`

Returns characters from middle.

**Example:**
```
=MID("Fluid Flow", 1, 5)    → "Fluid"
```

## LEN

`LEN(text)`

Returns length of text.

**Example:**
```
=LEN("Hello")    → 5
```

## LOWER

`LOWER(text)`

Converts to lowercase.

**Example:**
```
=LOWER("HELLO")    → "hello"
```

## UPPER

`UPPER(text)`

Converts to uppercase.

**Example:**
```
=UPPER("hello")    → "HELLO"
```

## TRIM

`TRIM(text)`

Removes leading/trailing spaces and collapses internal spaces.

**Example:**
```
=TRIM("  Hello   World  ")    → "Hello World"
```

## TEXT

`TEXT(value, format_text)`

Formats a number as text.

**Example:**
```
=TEXT(1234.567, "#,##0.00")    → "1,234.57"
```

## FIND

`FIND(find_text, within_text, [start_num])`

Case-sensitive search.

**Example:**
```
=FIND("M", "Miriam McGovern")    → 1
```

## SEARCH

`SEARCH(find_text, within_text, [start_num])`

Case-insensitive search with wildcards.

**Example:**
```
=SEARCH("e", "Statements", 6)    → 7
```

## SUBSTITUTE

`SUBSTITUTE(text, old_text, new_text, [instance_num])`

Replace text.

**Example:**
```
=SUBSTITUTE("Quarter 1, 2024", "1", "2", 1)    → "Quarter 2, 2024"
```

## REPLACE

`REPLACE(old_text, start_num, num_chars, new_text)`

Replace by position.

**Example:**
```
=REPLACE("abcdefg", 3, 2, "XY")    → "abXYefg"
```

## VALUE

`VALUE(text)`

Convert text to number.

**Example:**
```
=VALUE("123.45")    → 123.45
```

## TEXTJOIN

`TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)`

Join with delimiter.

**Example:**
```
=TEXTJOIN(", ", TRUE, "Sun", "", "Moon")    → "Sun, Moon"
```

## PROPER

`PROPER(text)`

Capitalize first letter of each word.

**Example:**
```
=PROPER("this is a TITLE")    → "This Is A Title"
```

## EXACT

`EXACT(text1, text2)`

Case-sensitive comparison.

**Example:**
```
=EXACT("Hello", "hello")    → FALSE
```

## REPT

`REPT(text, number_times)`

Repeats text a given number of times.

**Example:**
```
=REPT("*-", 3)    → "*-*-*-"
```

## CHAR

`CHAR(number)`

Returns the character specified by the code number.

**Example:**
```
=CHAR(65)    → "A"
```

## CODE

`CODE(text)`

Returns a numeric code for the first character.

**Example:**
```
=CODE("A")    → 65
```

## CLEAN

`CLEAN(text)`

Removes non-printable characters (codes 0-31).

**Example:**
```
=CLEAN("Line1" & CHAR(9) & "Line2")    → "Line1Line2"
```

## DOLLAR

`DOLLAR(number, [decimals])`

Formats a number as currency.

**Example:**
```
=DOLLAR(1234.567, 2)    → "$1,234.57"
```

## FIXED

`FIXED(number, decimals, [no_commas])`

Formats a number with fixed decimals.

**Example:**
```
=FIXED(1234.567, 2)    → "1,234.57"
```

## T

`T(value)`

Returns the text if value is text, otherwise empty string.

**Example:**
```
=T("Hello")    → "Hello"
```

## NUMBERVALUE

`NUMBERVALUE(text, [decimal_separator], [group_separator])`

Converts text to number.

**Example:**
```
=NUMBERVALUE("1.234,56", ",", ".")    → 1234.56
```

## UNICHAR

`UNICHAR(number)`

Returns the Unicode character for a code point.

**Example:**
```
=UNICHAR(9731)    → "☃"
```

## UNICODE

`UNICODE(text)`

Returns the Unicode code point for the first character.

**Example:**
```
=UNICODE("A")    → 65
```

## TEXTBEFORE

`TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])`

Returns text before a delimiter.

**Example:**
```
=TEXTBEFORE("Red-Blue-Green", "-")    → "Red"
```

## TEXTAFTER

`TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])`

Returns text after a delimiter.

**Example:**
```
=TEXTAFTER("Red-Blue-Green", "-")    → "Blue-Green"
```

## TEXTSPLIT

`TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])`

Splits text by delimiters into an array.

**Example:**
```
=TEXTSPLIT("Jan,Feb,Mar", ",")    → {"Jan", "Feb", "Mar"}
```

## ARRAYTOTEXT

`ARRAYTOTEXT(array, [format])`

Converts an array to text.

**Example:**
```
=ARRAYTOTEXT({1, "hello", 3}, 1)    → "{1,"hello",3}"
```

## VALUETOTEXT

`VALUETOTEXT(value, [format])`

Converts a value to text.

**Example:**
```
=VALUETOTEXT("Hello", 1)    → ""Hello""
```

## ASC

`ASC(text)`

Converts full-width (CJK) characters to half-width.

**Example:**
```
=ASC("ＡＢＣ")    → "ABC"
```

## DBCS

`DBCS(text)`

Converts half-width characters to full-width (CJK).

**Example:**
```
=DBCS("ABC")    → "ＡＢＣ"
```

## BAHTTEXT

`BAHTTEXT(number)`

Converts a number to Thai Baht text.

**Example:**
```
=BAHTTEXT(100)    → "หนึ่งร้อยบาทถ้วน"
```
