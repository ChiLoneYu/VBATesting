VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCRC32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private CRCTable(0 To 255) As Long
 
Public Function CalcCRC32(FilePath As String) As Long
Dim ByteArray() As Byte
Dim Limit As Long
Dim CRC As Long
Dim Temp1 As Long
Dim Temp2 As Long
Dim I As Long
Dim intFF As Integer
  
  intFF = FreeFile
  Open FilePath For Binary Access Read As #intFF
    Limit = LOF(intFF)
    ReDim ByteArray(Limit - 1)
    Get #intFF, , ByteArray
  Close #intFF
  
  Limit = Limit - 1
  CRC = -1
  For I = 0 To Limit
    If CRC < 0 Then
      Temp1 = CRC And &H7FFFFFFF
      Temp1 = Temp1 \ 256
      Temp1 = (Temp1 Or &H800000) And &HFFFFFF
    Else
      Temp1 = (CRC \ 256) And &HFFFFFF
    End If
    Temp2 = ByteArray(I)   ' get the byte
    Temp2 = CRCTable((CRC Xor Temp2) And &HFF)
    CRC = Temp1 Xor Temp2
  Next I
  CRC = CRC Xor &HFFFFFFFF
  CalcCRC32 = CRC
End Function
 
Private Sub Class_Initialize()
Dim I As Integer
Dim J As Integer
Dim Limit As Long
Dim CRC As Long
Dim Temp1 As Long
  Limit = &HEDB88320
  For I = 0 To 255
    CRC = I
    For J = 8 To 1 Step -1
      If CRC < 0 Then
        Temp1 = CRC And &H7FFFFFFF
        Temp1 = Temp1 \ 2
        Temp1 = Temp1 Or &H40000000
      Else
        Temp1 = CRC \ 2
      End If
      If CRC And 1 Then
        CRC = Temp1 Xor Limit
      Else
        CRC = Temp1
      End If
    Next J
    CRCTable(I) = CRC
  Next I
End Sub

