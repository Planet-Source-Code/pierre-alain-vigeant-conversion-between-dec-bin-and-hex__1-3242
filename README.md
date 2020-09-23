<div align="center">

## Conversion between Dec, Bin and Hex


</div>

### Description

This module contain function that are used to convert between decimal, binary and hexadecimal.
 
### More Info
 
Depend on the function

Each function are 'stand-alone'. This mean that u can copy one of them without needing another one.

The conversion function are written in this way: <from>2<to>

Example: The function 'Dec2Bin' will convert from decimal to binary


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Pierre\-Alain Vigeant](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/pierre-alain-vigeant.md)
**Level**          |Unknown
**User Rating**    |4.9 (34 globes from 7 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/pierre-alain-vigeant-conversion-between-dec-bin-and-hex__1-3242/archive/master.zip)





### Source Code

```
'**************************************************************
'*             Best Tools             *
'*             Conversion             *
'*         v2.1 (Improved performance)        *
'*              for VB              *
'*                              *
'*This module contain a lot of subs and functions for basic  *
'*conversion between Hexadecimal, Binary and decimal.     *
'**************************************************************
Option Explicit
Public Function Bin2Dec(ByVal sBin As String) As Long
 Dim i As Integer
 For i = 1 To Len(sBin)
  Bin2Dec = Bin2Dec + CLng(CInt(Mid(sBin, Len(sBin) - i + 1, 1)) * 2 ^ (i - 1))
 Next i
End Function
Public Function Bin2Hex(ByVal sBin As String) As String
 Dim i As Integer
 Dim nDec As Long
 sBin = String(4 - Len(sBin) Mod 4, "0") & sBin 'Add zero to complete Byte
 For i = 1 To Len(sBin)
  nDec = nDec + CInt(Mid(sBin, Len(sBin) - i + 1, 1)) * 2 ^ (i - 1)
 Next i
 Bin2Hex = Hex(nDec)
 If Len(Bin2Hex) Mod 2 = 1 Then Bin2Hex = "0" & Bin2Hex
End Function
Public Function Dec2Bin(ByVal nDec As Integer) As String
 'This function is the same then Hex2Bin, but it has been copied to speed up process
 Dim i As Integer
 Dim j As Integer
 Dim sHex As String
 Const HexChar As String = "0123456789ABCDEF"
 sHex = Hex(nDec) 'That the only part that is different
 For i = 1 To Len(sHex)
  nDec = InStr(1, HexChar, Mid(sHex, i, 1)) - 1
  For j = 3 To 0 Step -1
   Dec2Bin = Dec2Bin & nDec \ 2 ^ j
   nDec = nDec Mod 2 ^ j
  Next j
 Next i
 'Remove the first unused 0
 i = InStr(1, Dec2Bin, "1")
 If i <> 0 Then Dec2Bin = Mid(Dec2Bin, i)
End Function
Public Function Hex2Bin(ByVal sHex As String) As String
 Dim i As Integer
 Dim j As Integer
 Dim nDec As Long
 Const HexChar As String = "0123456789ABCDEF"
 For i = 1 To Len(sHex)
  nDec = InStr(1, HexChar, Mid(sHex, i, 1)) - 1
  For j = 3 To 0 Step -1
   Hex2Bin = Hex2Bin & nDec \ 2 ^ j
   nDec = nDec Mod 2 ^ j
  Next j
 Next i
 'Remove the first unused 0
 i = InStr(1, Hex2Bin, "1")
 If i <> 0 Then Hex2Bin = Mid(Hex2Bin, i)
End Function
Public Function Hex2Dec(ByVal sHex As String) As Long
 Dim i As Integer
 Dim nDec As Long
 Const HexChar As String = "0123456789ABCDEF"
 For i = Len(sHex) To 1 Step -1
  nDec = nDec + (InStr(1, HexChar, Mid(sHex, i, 1)) - 1) * 16 ^ (Len(sHex) - i)
 Next i
 Hex2Dec = CStr(nDec)
End Function
Public Function HiWord(ByVal DWord As Long) As Long
 HiWord = (DWord \ 65536) And &HFFFF
End Function
Public Function LoWord(ByVal DWord As Long) As Long
 LoWord = DWord And &HFFFF
End Function
Public Function DWord(ByVal HiWord As Long, ByVal LoWord As Long) As Long
 DWord = ((LoWord And 65536) Or ((HiWord And 65536) * 65536))
End Function
```

