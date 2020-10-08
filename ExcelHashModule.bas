Attribute VB_Name = "Module1"
Option Explicit
'mdlExcelHash
Public Function getExcelPasswordHash(pass As String)
Dim PassBytes() As Byte
PassBytes = StrConv(pass, vbFromUnicode)
Dim cchPassword As Long
cchPassword = UBound(PassBytes) + 1
Dim wPasswordHash As Long
If cchPassword = 0 Then
getExcelPasswordHash = wPasswordHash
Exit Function
End If
Dim pch As Long
pch = cchPassword - 1
While pch >= 0
wPasswordHash = wPasswordHash Xor PassBytes(pch)
wPasswordHash = RotateLeft_15bit(wPasswordHash, 1)
pch = pch - 1
Wend
wPasswordHash = wPasswordHash Xor cchPassword
wPasswordHash = wPasswordHash Xor &HCE4B&
getExcelPasswordHash = wPasswordHash
End Function
Private Function RotateLeft_15bit(num As Long, Count As Long) As Long
Dim outLong As Long
Dim i As Long
outLong = num
For i = 0 To Count - 1
outLong = ((outLong \ 2 ^ 14) And &H1) Or ((outLong * 2) And &H7FFF) 'Rotates left around 15 bits, kind of a signed rotateleft
Next
RotateLeft_15bit = outLong
End Function
