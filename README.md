# PStrings

eng | [rus](README_ru.md)

PStrings is a lightweight VBA library that provides various string manipulation functions to simplify common tasks in Visual Basic for Applications (VBA) programming.

## Table of Contents

- [Installation](#installation)
- [Usage](#usage)
- [Functions](#functions)
  - [FromatBytes](#formatbytes)
  - [Wrap](#wrap)
  - [IndexOf](#indexof)
  - [LastIndexOf](#lastindexof)
  - [CharCodeAt](#charcodeat)
  - [CharAt](#charat)
  - [Slice](#slice)
  - [StartsWith](#startswith)
  - [EndsWith](#endswith)
  - [Join](#join)
  - [Trim](#trim)
  - [JoinNonEmpty](#joinnonempty)
  - [FString](#fstring)
  - [IsNullString](#isnullstring)
  - [FormatString](#formatstring)
  - [InString](#instring)
  - [IsEqual](#isequals)

## Installation

To use PStrings in your VBA project, follow these steps:

1. Download the PStrings module file from the [releases](https://github.com/artemdorozhkin/PStrings/releases/) section of this repository.
2. Import the PStrings module into your VBA project:
   - Open your Excel workbook or VBA project.
   - Press `Alt + F11` to open the VBA editor.
   - Go to `File > Import File` and select the downloaded PStrings module file.
3. Once imported, you can start using the PStrings functions in your VBA code.

## Usage

Here's how you can use PStrings in your VBA code:

```vb
' Example usage of the PStrings functions
Sub ExampleUsage()
    Dim text As String
    text = "Hello, {0}!"

    ' Example of using FString function
    Debug.Print FString(text, "world") ' Output: Hello, world!
End Sub
```

## Functions

### FormatBytes

```vb
Public Function FormatBytes(ByVal Amount As Long) As String
```

#### Description:

Formats a string representing the amount, specifying the output in kilobytes, megabytes, or gigabytes.

#### Parameters:

- `Amount`: The amount in bytes.

#### Example:

```vb
Debug.Print FormatBytes(1023)       ' 1023 bytes
Debug.Print FormatBytes(1024)       ' 1.00 KB
Debug.Print FormatBytes(1048576)    ' 1.00 MB
Debug.Print FormatBytes(1073741824) ' 1.00 GB
```

#### Returns:

A formatted string representing the amount with appropriate unit of measurement (bytes, KB, MB, GB).

### Wrap

```vb
Public Function Wrap(ByVal Text As String, ByVal Wrapper As String) As String
```

#### Description:

The `Wrap()` function wraps the input text in the specified wrapper and returns the new string.

#### Parameters:

- `Text`: The input text.
- `Wrapper`: The wrapper string.

#### Example:

```vb
Debug.Print Wrap("ABC", Chr(34)) ' "ABC"
Debug.Print Wrap("ABC", "ABC")   ' ABCABCABC
```

#### Returns:

A string representing the input text wrapped in the specified wrapper.

### IndexOf

```vb
Public Function IndexOf(ByVal Text As String, ByVal Char As String) As Integer
```

#### Description:

Returns the index of the first occurrence of a specified character within a given text string.

#### Parameters:

- `Text`: The text string to search within.
- `Char`: The character to search for within the text string.

#### Returns:

The index of the first occurrence of the specified character within the text string. If the character is not found, returns -1.

#### Remarks:

This function iterates through each character in the text string and returns the index of the first occurrence of the specified character.

#### Example:

```vb
Debug.Print IndexOf("hello", "e")   ' Output: 2
Debug.Print IndexOf("hello", "z")   ' Output: -1
```

### LastIndexOf

```vb
Public Function LastIndexOf(ByVal Text As String, ByVal Char As String) As Integer
```

#### Description:

Returns the index of the last occurrence of a specified character within a given text string.

#### Parameters:

- `Text`: The text string to search within.
- `Char`: The character to search for within the text string.

#### Returns:

The index of the last occurrence of the specified character within the text string. If the character is not found, returns -1.

#### Remarks:

This function iterates through each character in the text string in reverse order and returns the index of the last occurrence of the specified character.

#### Example:

```vb
Debug.Print LastIndexOf("hello", "l")   ' Output: 4
Debug.Print LastIndexOf("hello", "z")   ' Output: -1
```

### CharCodeAt

```vb
Public Function CharCodeAt(ByVal Text As String, Optional ByVal Index As Integer = 0) As Integer
```

#### Description:

The `CharCodeAt()` function returns the Unicode value from 0 to 65535, representing the ASCII character code based on the specified index.

#### Parameters:

- `Text`: The input string.
- `Index`: The index from 0 to Len(Str) - 1. If the index is not provided, defaults to the first character.

#### Returns:

The Unicode value representing the ASCII character code at the specified index. If an invalid string is provided, the function returns -1.

#### Example:

```vb
Debug.Print CharCodeAt("ABC", 0) ' Output: 65
```

### CharAt

```vb
Public Function CharAt(ByVal Text As String, Optional ByVal Index As Integer = 0) As String
```

#### Description:

The `CharAt()` function returns a new string consisting of a single character extracted from a specified index position within the input string.

#### Parameters:

- `Text`: The input string.
- `Index`: The index from 0 to Len(Str) - 1. If the index is not provided, defaults to the first character.

#### Example:

In the following example, characters are accessed in various positions within the string "Brave new world":

```vb
Dim AnyString As String: AnyString = "Brave new world"
Debug.Print "Character at index 0 is " & CharAt(AnyString)
' Without specifying the index defaults to 0.

Debug.Print "Character at index 0 is " & CharAt(AnyString, 0)
Debug.Print "Character at index 1 is " & CharAt(AnyString, 1)
Debug.Print "Character at index 2 is " & CharAt(AnyString, 2)
Debug.Print "Character at index 3 is " & CharAt(AnyString, 3)
Debug.Print "Character at index 4 is " & CharAt(AnyString, 4)
Debug.Print "Character at index 999 is " & CharAt(AnyString, 999)
```

These lines display the following:

```vb
Character at index 0 is 'B'
Character at index 0 is 'B'
Character at index 1 is 'r'
Character at index 2 is 'a'
Character at index 3 is 'v'
Character at index 4 is 'e'
Character at index 999 is ''
```

### Slice

```vb
Public Function Slice(ByVal Text As String, ByVal StartIndex As Integer, Optional ByVal EndIndex As Integer = -1) As String
```

#### Description:

Returns a substring of the given text and concatenates it into a new string, excluding the specified end index.

#### Parameters:

- `Text`: The input string.
- `StartIndex`: The index of the first character to include in the resulting substring.
- `EndIndex`: The index of the character immediately following the end of the desired substring. Defaults to -1, indicating the end of the string.

#### Example:

The following example demonstrates the `Slice()` function for creating a new substring:

```vb
Dim Text1 As String: Text1 = "The morning is upon us." ' The length of Text1 is 23.
Dim Text2 As String: Text2 = Slice(Text1, 1, 8)
Dim Text3 As String: Text3 = Slice(Text1, 4, -2)
Dim Text4 As String: Text4 = Slice(Text1, 12)
Dim Text5 As String: Text5 = Slice(Text1, 30)
Debug.Print Text2   ' "he morn"
Debug.Print Text3   ' "morning is upon u"
Debug.Print Text4   ' "is upon us."
Debug.Print Text5   ' ""
```

The following example demonstrates the `Slice()` function with default start index:

```vb
Dim Text1 As String: Text1 = "The morning is upon us."
Slice(Text1, -3)        ' "us."
Slice(Text1, -3, -1)    ' "us"
Slice(Text1, 0, -1)     ' "The morning is upon us"
Slice(Text1, 4, -1)     ' "morning is upon us"
```

In this example, slicing starts from the 11th character and ends at the 16th character:

```vb
Slice(Text1, -11, 16)   ' "is u"
```

These examples slice from the 5th character to the 1st character:

```vb
Slice(Text1, -5, -1)   ' "n us"
```

### StartsWith

```vb
Public Function StartsWith(ByVal Text As String, ByVal Expression As String, Optional ByVal Compare As VbCompareMethod = VbCompareMethod.vbBinaryCompare) As Boolean
```

#### Description:

Checks if the beginning of the text matches the specified `Expression`.

#### Parameters:

- `Text`: The text to be checked.
- `Expression`: The value to be searched for.
- `Compare`: The comparison method. Enumeration `VbCompareMethod`. Defaults to `vbBinaryCompare`.

#### Returns:

- `True` if the beginning of the text matches the specified expression; otherwise, `False`.

#### Example:

The following example returns `True` because the word "Check" starts with "che" and the comparison method `vbTextCompare` is selected:

```vb
Debug.Print StartsWith("Check", "che", vbTextCompare)
```

The following example returns `False` because the default comparison method `vbBinaryCompare` is selected:

```vb
Debug.Print StartsWith("Check", "che")
```

### EndsWith

```vb
Public Function EndsWith(ByVal Text As String, ByVal Expression As String, Optional ByVal Compare As VbCompareMethod = VbCompareMethod.vbBinaryCompare) As Boolean
```

#### Description:

Checks if the end of the text matches the specified `Expression`.

#### Parameters:

- `Text`: The text to be checked.
- `Expression`: The value to be searched for.
- `Compare`: The comparison method. Enumeration `VbCompareMethod`. Defaults to `vbBinaryCompare`.

#### Returns:

- `True` if the end of the text matches the specified expression; otherwise, `False`.

#### Example:

The following example returns `True` because the word "Check" ends with "ECK" and the comparison method `vbTextCompare` is selected:

```vb
Debug.Print EndsWith("Check", "ECK", vbTextCompare)
```

The following example returns `False` because the default comparison method `vbBinaryCompare` is selected:

```vb
Debug.Print EndsWith("Check", "ECK")
```

### Join

```vb
Public Function Join(ByVal Delimiter As String, ByVal ParamArray Values() As String) As String
```

#### Description:

Concatenates multiple strings with a specified delimiter.

#### Parameters:

- `Delimiter`: The delimiter.
- `Values`: The strings to concatenate.

#### Returns:

The concatenated string with the specified delimiter.

#### Example:

```vb
Dim Value1 As String: Value1 = "Joined string"
Dim Value2 As String: Value2 = "from two strings"
Debug.Print Join(" ", Value1, Value2) ' Joined string from two strings
```

### Trim

```vb
Public Function Trim(ByVal Text As String) As String
```

#### Description:

Removes leading and trailing spaces from a string.

#### Parameters:

- `Text`: The string with leading and trailing spaces.

#### Example:

```vb
Dim Text As String: Text = "  String  with     whitespaces    "
Debug.Print ">" & Trim(Text) & "<" ' >String with whitespaces<
```

### JoinNonEmpty

```vb
Public Function JoinNonEmpty(ByVal Data As Variant, Optional ByVal Delimiter As String = ", ") As String
```

#### Description:

Concatenates elements of the given array, excluding empty values (`vbNullString` or `Empty`), with a specified delimiter.

#### Parameters:

- `Data`: The array of data.
- `Delimiter`: The delimiter. Defaults to a comma and space.

#### Returns:

The concatenated string with non-empty values separated by the specified delimiter.

#### Example:

```vb
Dim DataWithEmpty As Variant
DataWithEmpty = Array("Value1", Empty, "Value2", "")
Debug.Print JoinNonEmpty(DataWithEmpty, ", ") ' Value1, Value2
```

### FString

```vb
Public Function FString(ByVal Text As String, ParamArray Values() As Variant) As String
```

#### Description:

Replaces placeholders in the text with corresponding values from the array `Values`.

#### Remarks:

For interpolation, use placeholders in the form of {0}, {1}, etc. corresponding to the index of the value in the array.
The function does not handle escape sequences such as: `\n`, `\t`, `\r`.

#### Example:

```vb
Dim Text as String
Text = "Example usage of function {0}!"
Dim FuncName as String
FuncName = "FString"
Debug.Print FString(Text, FuncName) ' Example usage of function FString!
```

### IsEmptyString

```vb
Public Function IsEmptyString(ByVal Expression As String) As Boolean
```

#### Description:

Checks if the given `String` value is empty.

#### Parameters:

- `Expression`: The value to check.

#### Returns:

`True` if the value is empty.

#### Example:

```vb
Dim str As String
str = ""
Debug.Print IsEmptyString(str) ' True

str = "Hello"
Debug.Print IsEmptyString(str) ' False
```

### FormatString

```vb
Public Function FormatString(ByVal Text As String, ByVal ParamArray Values() As Variant) As String
```

#### Description:

Formats placeholders in the text with corresponding values from the array `Values`.

#### Remarks:

For interpolation, use placeholders in the following format:

- `%s` for String values
- `%t` for Date values
- `%d` for numeric values

#### Example:

```vb
Dim Text as String
Text = "Example usage of function %s!"
Dim FuncName as String
FuncName = "FormatString"
Debug.Print FormatString(Text, FuncName) ' Result: "Example usage of function FormatString!"
```

#### Parameters:

- `Text`: The text with placeholders.

#### Returns:

The formatted text.

### InString

```vb
Public Function InString(ByVal Text As String, ByVal ParamArray Values() As Variant) As Boolean
```

#### Description:

Checks if any of the `Values` are present in the `Text` string.

#### Parameters:

- `Text`: The text to search within.
- `Values`: The values to search for.

#### Returns:

Returns `True` if any of the specified values are found in the text.

### IsEqual

```vb
Public Function IsEqual(ByVal Text1 As String, ByVal Text2 As String, Optional ByVal Compare As VbCompareMethod = VbCompareMethod.vbTextCompare) As Boolean
```

#### Description:

Compares the equality of `Text1` and `Text2`.

#### Parameters:

- `Text1`: The first string.
- `Text2`: The second string.
- `Compare`: Comparison method. Defaults to `vbTextCompare`.

#### Returns:

Returns `True` if the strings are equal.
