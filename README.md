<div align="center">

## CReplaceString


</div>

### Description

CReplaceString is a function that I decided to code when I was bored. Had no idea what to code so I decided to code an advanced ReplaceString function. The concept is simple, it will replace a specified text in the string with something else.
 
### More Info
 
Public Function CReplaceString(ToFind As String, ReplaceWith As String, ToSearch As String, CaseSensitive As Boolean) As String

All explinations are in the function, but here is the usage:

Say you have a string, lets say UR_String and within that string you have the word or phrase, "replace_Me", and you want to replace that with "im_replaced". This is all you need to do

UR_String = CReplaceString("replace_Me", "im_replaced", UR_String, True)

The case sensitive is the advanced section of this. A normal replace string function would not have this. Could come in handy :)

The fixed string the replaced text.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[PowerElite](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/powerelite.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/powerelite-creplacestring__1-33435/archive/master.zip)





### Source Code

```
Public Function CReplaceString(ToFind As String, ReplaceWith As String, ToSearch As String, CaseSensitive As Boolean) As String
'Alias: Crazy
'CReplaceString Function
'This is a advanced replacestring.
'There are obviously the basic replace
'string options with the now included
'case sensitive.
'Casing being Ex: aKa, AkA, aka
'Dim our variables
Dim FoundLeft      As String
Dim FoundRight     As String
Dim Found        As Integer
'If in the string, get it's first location
If CaseSensitive = True Then
    Found = InStr(1, ToSearch$, ToFind$)
ElseIf CaseSensitive = False Then
    Found = InStr(1, LCase(ToSearch$), LCase(ToFind$))
End If
'If the string you want to replace
'is not there, Found will = 0, and this
'If/Then statement will be skipped.
'If Found is not equal to 0 (<>) then
'it will enter the If/Then statement
'and follow the rest of the function
If Found <> 0& Then
  Do
    FoundLeft = Left(ToSearch$, Found - 1)
    FoundRight = Mid(ToSearch$, Found + Len(ToFind$), Len(ToSearch$) - Found + Len(ToFind$))
    ToSearch$ = FoundLeft & ReplaceWith$ & FoundRight
    'Gets next location of string
    'that you want to replace
    If CaseSensitive = True Then
      Found = InStr(Found + 1, ToSearch$, ToFind$)
    ElseIf CaseSensitive = False Then
      Found = InStr(Found + 1, LCase(ToSearch$), LCase(ToFind$))
    End If
  Loop Until Found = 0& 'Will exit loop if no longer found
  'Set the new string
  CReplaceString = ToSearch$
ElseIf Found = 0& Then 'If what you are looking for
            'is not in the string then
            'it will just keep it the
            'same as it was when the
            'function was initiated
  CReplaceString = ToSearch$
End If
End Function
```

