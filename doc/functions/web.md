# Web & Regex Functions

## ENCODEURL

`ENCODEURL(text)`

URI percent-encodes a string.

**Example:**
```
=ENCODEURL("Hello World")    → "Hello%20World"
```

## REGEXMATCH

`REGEXMATCH(text, regular_expression)`

Returns TRUE if text matches the regular expression pattern.

**Example:**
```
=REGEXMATCH("Hello123", "\d+")    → TRUE
```

## REGEXEXTRACT

`REGEXEXTRACT(text, regular_expression)`

Returns the first substring that matches the regular expression pattern.

**Example:**
```
=REGEXEXTRACT("Hello123World", "\d+")    → "123"
```

## REGEXREPLACE

`REGEXREPLACE(text, regular_expression, replacement)`

Replaces all matches of a regular expression pattern with a replacement string.

**Example:**
```
=REGEXREPLACE("Hello 123 World 456", "\d+", "#")    → "Hello # World #"
```
