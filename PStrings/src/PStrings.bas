Attribute VB_Name = "PStrings"
'@Folder "PStrings.src"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function StrFormatByteSize _
      Lib "shlwapi" Alias "StrFormatByteSizeA" _
      (ByVal dw As Long, ByVal pszBuf As String, _
      ByVal cchbuf As Long) As Long
#Else
Private Declare Function StrFormatByteSize _
      Lib "shlwapi" Alias "StrFormatByteSizeA" _
      (ByVal dw As Long, ByVal pszBuf As String, _
      ByVal cchbuf As Long) As Long
#End If

''' <summary>
''' Formats a string representing the amount, specifying the output in kilobytes, megabytes, or gigabytes.
''' </summary>
''' <param name="Amount">
''' The amount in bytes.
''' </param>
''' <example>
''' The following examples demonstrate formatting bytes into different units of measurement:
''' <code>
''' Debug.Print FormatBytes(1023)       ' 1023 bytes
''' Debug.Print FormatBytes(1024)       ' 1.00 KB
''' Debug.Print FormatBytes(1048576)    ' 1.00 MB
''' Debug.Print FormatBytes(1073741824) ' 1.00 GB
''' </code>
''' </example>
Public Function FormatBytes(ByVal Amount As Long) As String
    Dim Buffer As String: Buffer = Space(255)

    Call StrFormatByteSize(Amount, Buffer, Len(Buffer))
    Dim NullCharIndex As Long: NullCharIndex = InStr(Buffer, vbNullChar)

    If NullCharIndex > 1 Then FormatBytes = Left(Buffer, NullCharIndex - 1)
End Function

''' <summary>
''' The Wrap() function wraps the input text in the specified wrapper and returns the new string.
''' </summary>
''' <param name="Text">
''' The input text.
''' </param>
''' <param name="Wrapper">
''' The wrapper string.
''' </param>
''' <example>
''' The following examples demonstrate wrapping the input text in different wrappers:
''' <code>
''' Debug.Print Wrap("ABC", Chr(34)) ' "ABC"
''' </code>
''' The following example demonstrates wrapping the input text in the "ABC" wrapper:
''' <code>
''' Debug.Print Wrap("ABC", "ABC") ' ABCABCABC
''' </code>
''' </example>
Public Function Wrap(ByVal Text As String, ByVal Wrapper As String) As String
    Wrap = Wrapper & Text & Wrapper
End Function

''' <summary>
''' Returns the index of the first occurrence of a specified character within a given text string.
''' </summary>
''' <param name="Text">
''' The text string to search within.
''' </param>
''' <param name="Char">
''' The character to search for within the text string.
''' </param>
''' <returns>
''' The index of the first occurrence of the specified character within the text string. If the character is not found, returns -1.
''' </returns>
''' <remarks>
''' This function iterates through each character in the text string and returns the index of the first occurrence of the specified character.
''' </remarks>
''' <example>
''' The following example demonstrates the usage of the IndexOf function:
''' <code>
''' Debug.Print IndexOf("hello", "e")   ' Output: 2
''' Debug.Print IndexOf("hello", "z")   ' Output: -1
''' </code>
''' </example>
Public Function IndexOf(ByVal Text As String, ByVal Char As String) As Integer
    Dim i As Integer
    For i = 1 To Strings.Len(Text)
        If CharAt(Text, i) = Char Then
            IndexOf = i
            Exit Function
        End If
    Next

    IndexOf = -1
End Function

''' <summary>
''' Returns the index of the last occurrence of a specified character within a given text string.
''' </summary>
''' <param name="Text">
''' The text string to search within.
''' </param>
''' <param name="Char">
''' The character to search for within the text string.
''' </param>
''' <returns>
''' The index of the last occurrence of the specified character within the text string. If the character is not found, returns -1.
''' </returns>
''' <remarks>
''' This function iterates through each character in the text string in reverse order and returns the index of the last occurrence of the specified character.
''' </remarks>
''' <example>
''' The following example demonstrates the usage of the LastIndexOf function:
''' <code>
''' Debug.Print LastIndexOf("hello", "l")   ' Output: 4
''' Debug.Print LastIndexOf("hello", "z")   ' Output: -1
''' </code>
''' </example>
Public Function LastIndexOf(ByVal Text As String, ByVal Char As String) As Integer
    Dim i As Integer
    For i = Strings.Len(Text) To 1 Step -1
        If CharAt(Text, i) = Char Then
            LastIndexOf = i
            Exit Function
        End If
    Next

    LastIndexOf = -1
End Function

''' <summary>
''' The CharCodeAt() function returns the Unicode value from 0 to 65535, representing the ASCII character code based on the specified index.
''' </summary>
''' <remarks>
''' If an invalid string is provided, the function returns -1.
''' </remarks>
''' <param name="Text">
''' The input string.
''' </param>
''' <param name="Index">
''' The index from 0 to Len(Str) - 1. If the index is not provided, defaults to the first character.
''' </param>
''' <example>
''' The following example demonstrates the CharCodeAt function returning 65, the ASCII value for A:
''' <code>
''' Debug.Print CharCodeAt("ABC", 0) ' 65
''' </code>
''' </example>
Public Function CharCodeAt(ByVal Text As String, Optional ByVal Index As Integer = 0) As Integer
    Dim Char As String: Char = PStrings.CharAt(Text, Index)
    If PStrings.IsNullString(Char) Then
        CharCodeAt = -1
        Exit Function
    End If

    CharCodeAt = Strings.Asc(Char)
End Function

''' <summary>
''' The CharAt() function returns a new string consisting of a single character extracted from a specified index position within the input string.
''' </summary>
''' <param name="Text">
''' The input string.
''' </param>
''' <param name="Index">
''' The index from 0 to Len(Str) - 1. If the index is not provided, defaults to the first character.
''' </param>
''' <example>
''' In the following example, characters are accessed in various positions within the string "Brave new world":
''' <code>
''' Dim AnyString As String: AnyString = "Brave new world"
''' Debug.Print "Character at index 0 is " & CharAt(AnyString)
''' ' Without specifying the index defaults to 0.
'''
''' Debug.Print "Character at index 0 is " & CharAt(AnyString, 0)
''' Debug.Print "Character at index 1 is " & CharAt(AnyString, 1)
''' Debug.Print "Character at index 2 is " & CharAt(AnyString, 2)
''' Debug.Print "Character at index 3 is " & CharAt(AnyString, 3)
''' Debug.Print "Character at index 4 is " & CharAt(AnyString, 4)
''' Debug.Print "Character at index 999 is " & CharAt(AnyString, 999)
''' </code>
''' These lines display the following:
''' <code>
''' Character at index 0 is 'B'
''' Character at index 0 is 'B'
''' Character at index 1 is 'r'
''' Character at index 2 is 'a'
''' Character at index 3 is 'v'
''' Character at index 4 is 'e'
''' Character at index 999 is ''
''' </code>
''' </example>
Public Function CharAt(ByVal Text As String, Optional ByVal Index As Integer = 0) As String
    If Index < 0 Or Index >= Strings.Len(Text) Then Exit Function
    CharAt = Strings.Mid(Text, Index + 1, 1)
End Function

''' <summary>
''' Returns a substring of the given text and concatenates it into a new string, excluding the specified end index.
''' </summary>
''' <param name="Text">
''' The input string.
''' </param>
''' <param name="StartIndex">
''' The index of the first character to include in the resulting substring.
''' </param>
''' <param name="EndIndex">
''' The index of the character immediately following the end of the desired substring.
''' </param>
''' <example>
''' The following example demonstrates the slice() function for creating a new substring.
''' <code>
''' Dim Text1 As String: Text1 = "The morning is upon us." ' The length of Text1 is 23.
''' Dim Text2 As String: Text2 = Slice(Text1, 1, 8)
''' Dim Text3 As String: Text3 = Slice(Text1, 4, -2)
''' Dim Text4 As String: Text4 = Slice(Text1, 12)
''' Dim Text5 As String: Text5 = Slice(Text1, 30)
''' Debug.Print Text2   ' "he morn"
''' Debug.Print Text3   ' "morning is upon u"
''' Debug.Print Text4   ' "is upon us."
''' Debug.Print Text5   ' ""
''' </code>
''' The following example demonstrates the slice() function with default start index.
''' <code>
''' Dim Text1 As String: Text1 = "The morning is upon us."
''' Slice(Text1, -3)        ' "us."
''' Slice(Text1, -3, -1)    ' "us"
''' Slice(Text1, 0, -1)     ' "The morning is upon us"
''' Slice(Text1, 4, -1)     ' "morning is upon us"
''' </code>
''' In this example, slicing starts from the 11th character and ends at the 16th character.
''' <code>
''' Slice(Text1, -11, 16)   ' "is u"
''' </code>
''' These examples slice from the 5th character to the 1st character.
''' <code>
''' Slice(Text1, -5, -1)   ' "n us"
''' </code>
''' </example>
Public Function Slice(ByVal Text As String, ByVal StartIndex As Integer, Optional ByVal EndIndex As Variant) As String
    Dim Length As Integer: Length = Strings.Len(Text)
    If StartIndex >= Length Then Exit Function

    If Information.IsMissing(EndIndex) Then EndIndex = 0
    If EndIndex >= Length Then EndIndex = Length

    Dim i As Integer
    If StartIndex < 0 Then
        Dim StartPart As String
        StartPart = Strings.Right(Text, -StartIndex)
    Else
        StartPart = Strings.Mid(Text, StartIndex + 1, Length)
    End If

    If EndIndex < 0 Then
        Dim EndPart As String
        EndPart = Strings.Right(Text, -EndIndex)
    Else
        EndPart = Strings.Mid(Text, EndIndex + 1, Length)
    End If

    Dim SliceIndex As Integer: SliceIndex = Strings.Len(StartPart) - Strings.Len(EndPart)
    If SliceIndex <= 0 Then Exit Function
    Slice = Strings.Mid(StartPart, 1, SliceIndex)
End Function

''' <summary>
''' Checks if the beginning of the text matches the specified <c>Expression</c>.
''' </summary>
''' <param name="Text">
''' The text to be checked.
''' </param>
''' <param name="Expression">
''' The value to be searched for.
''' </param>
''' <param name="Compare">
''' The comparison method. Enumeration <c>VbCompareMethod</c>. Defaults to <c>vbBinaryCompare</c>.
''' </param>
''' <example>
''' The following example returns <c>True</c> because the word "Check" starts with "che" and the comparison method <c>vbTextCompare</c> is selected:
''' <code>
''' Debug.Print StartsWith("Check", "che", vbTextCompare)
''' </code>
''' The following example returns <c>False</c> because the default comparison method <c>vbBinaryCompare</c> is selected:
''' <code>
''' Debug.Print StartsWith("Check", "che")
''' </code>
''' </example>
''' </summary>
Public Function StartsWith(ByVal Text As String, ByVal Expression As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Boolean
    StartsWith = Strings.StrComp(Left(Text, Len(Expression)), Expression, Compare) = 0
End Function

''' <summary>
''' Checks if the end of the text matches the specified <c>Expression</c>.
''' </summary>
''' <param name="Text">
''' The text to be checked.
''' </param>
''' <param name="Expression">
''' The value to be searched for.
''' </param>
''' <param name="Compare">
''' The comparison method. Enumeration <c>VbCompareMethod</c>. Defaults to <c>vbBinaryCompare</c>.
''' </param>
''' <example>
''' The following example returns <c>True</c> because the word "Check" ends with "ECK" and the comparison method <c>vbTextCompare</c> is selected:
''' <code>
''' Debug.Print EndsWith("Check", "ECK", vbTextCompare)
''' </code>
''' The following example returns <c>False</c> because the default comparison method <c>vbBinaryCompare</c> is selected:
''' <code>
''' Debug.Print EndsWith("Check", "ECK")
''' </code>
''' </example>
''' </summary>
Public Function EndsWith(ByVal Text As String, ByVal Expression As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Boolean
    EndsWith = Strings.StrComp(Right(Text, Len(Expression)), Expression, Compare) = 0
End Function

''' <summary>
''' Concatenates multiple strings with a specified delimiter.
''' </summary>
''' <example>
''' <code>
'''     Dim Value1 As String: Value1 = "Joined string"
'''     Dim Value2 As String: Value2 = "from two strings"
'''     Debug.Print JoinStrings(" ", Value1, Value2) ' Joined string from two strings
''' </code>
''' </example>
''' <param name="Delimiter">The delimiter.</param>
''' <param name="Values">The strings to concatenate.</param>
''' <returns>The concatenated string with the specified delimiter.</returns>
''' </summary>
Public Function Join(ByVal Delimiter As String, ParamArray Values() As Variant) As String
    Join = Strings.Join(Values, Delimiter)
End Function

''' <summary>
''' Removes leading and trailing spaces from a string.
''' </summary>
''' <example>
''' <code>
'''     Dim Text As String: Text = "  String  with     whitespaces    "
'''     Debug.Print ">" & Trim(Text) & "<" ' >String with whitespaces<
''' </code>
''' </example>
''' <param name="Text">The string with leading and trailing spaces.</param>
''' </summary>
Public Function Trim(ByVal Text As String) As String
    Trim = WorksheetFunction.Trim(Text)
End Function

''' <summary>
''' Concatenates elements of the given array, excluding empty values (vbNullString or Empty), with a specified delimiter.
''' </summary>
''' <example>
''' <code>
'''     Dim DataWithEmpty As Variant
'''     DataWithEmpty = Array("Value1", Empty, "Value2", "")
'''     Debug.Print JoinNonEmpty(DataWithEmpty, ", ") ' Value1, Value2
''' </code>
''' </example>
''' <param name="Data">The array of data.</param>
''' <param name="Delimiter">The delimiter. Defaults to a comma and space.</param>
''' <returns>The concatenated string with non-empty values separated by the specified delimiter.</returns>
''' </summary>
Public Function JoinNonEmpty(ByVal Data As Variant, Optional ByVal Delimiter As String = " ") As String
    Dim NonEmptyValues As Object: Set NonEmptyValues = CreateObject("System.Collections.ArrayList")
    Dim Value As Variant
    For Each Value In Data
        If Len(Value) <> 0 Then NonEmptyValues.Add Value
    Next

    JoinNonEmpty = Strings.Join(NonEmptyValues.ToArray(), Delimiter)
End Function

''' <summary>
''' Replaces placeholders in the text with corresponding values from the array <c>Values</c>.
''' </summary>
''' <remarks>
''' For interpolation, use placeholders in the form of {0}, {1}, etc. corresponding to the index of the value in the array.
''' The function does not handle escape sequences such as: <c>\n</c>, <c>\t</c>, <c>\r</c>.
''' </remarks>
''' <example>
''' <code>
'''     Dim Text as String
'''     Text = "Example usage of function {0}!"
'''     Dim FuncName as String
'''     FuncName = "FString"
'''     Debug.Print FString(Text, FuncName) ' Example usage of function FString!
''' </code>
''' </example>
''' <param name="Text">The text with placeholders for substitution.</param>
''' <returns>The interpolated text.</returns>
''' </summary>
Public Function FString(ByVal Text As String, ParamArray Values() As Variant) As String
    Dim FormatedString As String: FormatedString = Text

    Dim i As Long
    For i = LBound(Values) To UBound(Values)
        Dim Plug As String: Plug = "{" & i & "}"
        Dim Value As Variant: Value = Values(i)
        If Information.IsMissing(Value) Then Value = Empty
        FormatedString = Strings.Replace(FormatedString, Plug, Value)
    Next

    FormatedString = Strings.Replace(FormatedString, "{CrLf}", vbCrLf, Compare:=vbTextCompare)

    FormatedString = Strings.Replace(FormatedString, "{Cr}", vbCr, Compare:=vbTextCompare)
    FormatedString = Strings.Replace(FormatedString, "\\r", vbCr, Compare:=vbTextCompare)

    FormatedString = Strings.Replace(FormatedString, "{Lf}", vbLf, Compare:=vbTextCompare)

    FormatedString = Strings.Replace(FormatedString, "{NewLine}", vbNewLine, Compare:=vbTextCompare)
    FormatedString = Strings.Replace(FormatedString, "\\n", vbNewLine, Compare:=vbTextCompare)

    FormatedString = Strings.Replace(FormatedString, "\\t", vbTab, Compare:=vbTextCompare)

    FString = FormatedString
End Function

''' <summary>
''' Checks if the given <c>String</c> value is empty.
''' </summary>
''' <param name="Expression">The value to check.</param>
''' <returns><c>True</c> if the value is empty.</returns>
''' </summary>
Public Function IsNullString(ByVal Expression As String) As Boolean
    IsNullString = Expression = vbNullString Or Information.IsEmpty(Expression)
End Function

''' <summary>
''' Formats placeholders in the text with corresponding values from the array <c>Values</c>.
''' </summary>
''' <remarks>
''' For interpolation, use placeholders in the following format:
'''     %s - for String values
'''     %t - for Date values
'''     %d - for numeric values
''' </remarks>
''' <example>
''' <code>
'''     Dim Text as String
'''     Text = "Example usage of function %s!"
'''     Dim FuncName as String
'''     FuncName = "FormatString"
'''     Debug.Print FormatString(Text, FuncName) ' Result: "Example usage of function FormatString!"
''' </code>
''' </example>
''' <param name="Text">The text with placeholders.</param>
''' <returns>The formatted text.</returns>
''' </summary>
Public Function FormatString(ByVal Text As String, ParamArray Values() As Variant) As String
    Const StringPlug As String = "%s": Const NumericPlug As String = "%d": Const DateTimePlug As String = "%t"
    Dim FormatedString As String: FormatedString = Text

ReplaceValues:
    Dim Value As Variant
    For Each Value In Values
        If Information.VarType(Value) = vbString Then
            FormatedString = Strings.Replace(FormatedString, StringPlug, Value, Count:=1, Compare:=vbTextCompare)
            GoTo NextValue
        End If
        If Information.VarType(Value) = vbDate Then
            FormatedString = Strings.Replace(FormatedString, DateTimePlug, Value, Count:=1, Compare:=vbTextCompare)
            GoTo NextValue
        End If
        FormatedString = Strings.Replace(FormatedString, NumericPlug, Value, Count:=1, Compare:=vbTextCompare)
NextValue:
    Next

    If PStrings.InString(FormatedString, StringPlug, NumericPlug, DateTimePlug) Then GoTo ReplaceValues

    FormatString = FormatedString
End Function

''' <summary>
''' Checks if any of the <c>Values</c> are present in the <c>Text</c> string.
''' </summary>
''' <param name="Text">The text to search within.</param>
''' <param name="Values">The values to search for.</param>
''' <returns>Returns <c>True</c> if any of the specified values are found in the text.</returns>
''' </summary>
Public Function InString(ByVal Text As String, ByVal Compare As VbCompareMethod, ParamArray Values() As Variant) As Boolean
    Dim Value As Variant
    For Each Value In Values
        If InStr(1, Text, Value, Compare) > 0 Then InString = True: Exit Function
    Next
End Function

''' <summary>
''' Compares the equality of <c>Text1</c> and <c>Text2</c>.
''' </summary>
''' <param name="Text1">The first string.</param>
''' <param name="Text2">The second string.</param>
''' <param name="Compare">Comparison method. Defaults to <c>vbTextCompare</c>.</param>
''' <returns>Returns <c>True</c> if the strings are equal.</returns>
''' </summary>
Public Function IsEqual(ByVal Text1 As String, ByVal Text2 As String, Optional ByVal Compare As VbCompareMethod = vbTextCompare) As Boolean
    IsEqual = StrComp(Text1, Text2, Compare) = 0
End Function


