Option Strict Off
Option Explicit On
Friend Class clsCRC32
	
	Private CRCTable(255) As Integer
	
	Public Function CalcCRC32(ByRef FilePath As String) As Integer
		Dim ByteArray() As Byte
		Dim Limit As Integer
		Dim CRC As Integer
		Dim Temp1 As Integer
		Dim Temp2 As Integer
		Dim i As Integer
		Dim intFF As Short
		
		intFF = FreeFile
		FileOpen(intFF, FilePath, OpenMode.Binary, OpenAccess.Read)
		Limit = LOF(intFF)
		ReDim ByteArray(Limit - 1)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(intFF, ByteArray)
		FileClose(intFF)
		
		Limit = Limit - 1
		CRC = -1
		For i = 0 To Limit
			If CRC < 0 Then
				Temp1 = CRC And &H7FFFFFFF
				Temp1 = Temp1 \ 256
				Temp1 = (Temp1 Or &H800000) And &HFFFFFF
			Else
				Temp1 = (CRC \ 256) And &HFFFFFF
			End If
			Temp2 = ByteArray(i) ' get the byte
			Temp2 = CRCTable((CRC Xor Temp2) And &HFF)
			CRC = Temp1 Xor Temp2
		Next i
		CRC = CRC Xor &HFFFFFFFF
		CalcCRC32 = CRC
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		Dim i As Short
		Dim J As Short
		Dim Limit As Integer
		Dim CRC As Integer
		Dim Temp1 As Integer
		Limit = &HEDB88320
		For i = 0 To 255
			CRC = i
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
			CRCTable(i) = CRC
		Next i
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class