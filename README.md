<div align="center">

## A more compact storage of bytes in string than using Hex


</div>

### Description

Storing data (anything) in a string can be done more compactly using this method than using Hex strings. *Updated - speedier main functions. I've also added functions for more compact storage of dates/times in strings than a simple byte translation.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[jeremyxtz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeremyxtz.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jeremyxtz-a-more-compact-storage-of-bytes-in-string-than-using-hex__1-63991/archive/master.zip)

### API Declarations

Copymemory


### Source Code

```
'**** A more-compact alternative to storing bytes
'in strings than using Hex ******
'Bytes can't be stored successfully in character strings
'because of problems with certain characters
'eg carriage return,linefeed,", nullchar etc
'This method avoids those characters by storing
'bit 128 of each byte in a header character
'and adding 128 (but could be any value above
'34 (chr 34 = ") to the byte so string characters
'will all be above the problem range
'The header is set at 128 initially so it too will be
'above the range and the remaining bits of the
'header 2^0,2^1... 2^6 are set depending
'on whether any of the next 7 bytes has bit 128
'Examples are for long and date variables but any
'data converted to a byte array can be stored for
'8 character per 7 bytes compared with 14 when using
'a predictable-length Hex string
'which is 2 characters per byte
'There's an obvious function overhead - you'd use
'this if you wanted to do something like a store amount of
'data in a constant (conversion to a string is the only way)
'Any compression to the data must be carried out before
'conversion using these functions so as not to undo the
'conversion
'***********
'no problem with characters CRLF or " when storing data in a constant
Const longtostring_Minus4597545 = "×Ø¹ÿ"
Const timeAdjust = 160 ' clear problem character
Const dateOffset = 38728
Dim powers(6) As Integer
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Sub Form_Load()
For i = 0 To 6
powers(i) = 2 ^ i 'for speed - store powers in lookup table
Next
Me.AutoRedraw = True
Me.Caption = "-4597545 " & longtostring(-4597545) & " " & stringtolong(longtostring(-4597545))
'no problem with null terminator (Chr 0) when saving to clipboard
Clipboard.SetText longtostring(-4597545), 1
Dim dd As Date
dd = Date + Time
Me.Print dd & "    " & DatetoString(dd) & "    " & stringToDate(DatetoString(dd))
Me.Print dd & "    " & DatetoString6(dd) & "    " & stringToDate6(DatetoString6(dd))
End Sub
'convert whatever data into bytes first
'long =4 bytes + 1 header byte = 5 character string result (HEX = 8)
Function longtostring(no As Long) As String
Dim b(3) As Byte
CopyMemory b(0), no, 4
longtostring = AnyToString(b)
End Function
Function stringtolong(st As String) As Long
If Len(st) <> 5 Then Exit Function
Dim b() As Byte
b = stringToAny(st)
Dim a As Long
CopyMemory stringtolong, b(0), 4
End Function
'date > 7 (8) bytes so process first 7 then last byte
'8 bytes + 2 header bytes = 10 characters (HEX = 16)
Function DatetoString(d As Date) As String
Dim b() As Byte, c(0) As Byte
ReDim b(7)
CopyMemory b(0), d, 8
c(0) = b(7)
ReDim Preserve b(6)
DatetoString = AnyToString(b) & AnyToString(c)
End Function
Function stringToDate(st As String) As Date
If Len(st) <> 10 Then Exit Function
Dim b() As Byte, c() As Byte
b = stringToAny(Left(st, 8))
c = stringToAny(Right(st, 2))
ReDim Preserve b(7)
b(7) = c(0)
Dim d As Date
CopyMemory stringToDate, b(0), 8
End Function
'*************main functions
'max 7 bytes for these functions
'for larger numbers eg date,user type, array process in up
'to 7 byte chunks
Function stringToAny(st As String) As Byte()
Dim b() As Byte, i As Long, c As Integer, header As Byte
b = st
header = b(0)
For i = 2 To UBound(b) - 1 Step 2
b(c) = b(i) - 128
If (header And powers(c)) Then b(c) = b(c) Or 128
c = c + 1
Next
ReDim Preserve b(Len(st) - 2)
stringToAny = b()
End Function
Function AnyToString(bb() As Byte) As String
Dim i As Long, header As Byte, d As Integer, b() As Byte
header = 128
ReDim Preserve b((UBound(bb) * 2) + 3)
For i = 0 To UBound(bb)
d = d + 2
If bb(i) And 128 Then header = header Or powers(i)
b(d) = bb(i) Or 128
Next
b(0) = header
AnyToString = b()
End Function
'*********** more compact storage of dates or times in 3 characters
'or 6 characters for a full date
'assumes date will be spread over integer range
'-32767 to 32768 = ~ 180 year range or 90 years either side of today
'if we use an offset of 38727 (today's date)
'which will suit many applications
'if needed we could increase this date range by using the 5
'unused bits of the header byte and still only need 3 characters
'but I haven't coded that.
Function DatetoString6(d As Date) As String
DatetoString6 = DatetoString3(d) & TimetoString3(d)
End Function
Function stringToDate6(st As String) As Date
stringToDate6 = stringtoDate3(Left(st, 3)) + stringtoTime3(Right(st, 3))
End Function
Function DatetoString3(d As Date) As String
Dim b(1) As Byte, t As Integer, s As Single
s = d
If s < 0 Then s = Fix(s) Else s = Int(s)
t = s - dateOffset
CopyMemory b(0), t, 2
DatetoString3 = AnyToString(b)
End Function
Function stringtoDate3(st As String) As Date
Dim b() As Byte
b = stringToAny(st)
Dim a As Integer
CopyMemory a, b(0), 2
stringtoDate3 = a + dateOffset
End Function
Function TimetoString3(d As Date) As String
TimetoString3 = Chr(Hour(d) + timeAdjust) & Chr(Minute(d) + timeAdjust) & Chr(Second(d) + timeAdjust)
End Function
Function stringtoTime3(st As String) As Date
stringtoTime3 = TimeSerial(Asc(st) - timeAdjust, Asc(Mid(st, 2, 1)) - timeAdjust, Asc(Mid(st, 3, 1)) - timeAdjust)
End Function
```

