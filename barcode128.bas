' --------------------------------------
' Barcode 128
' For LibreOffice Calc or Excel
' Convert text for use with a barcode 128 font
' BdR - bdr1976@gmail.com
' --------------------------------------
Function BARCODE128_ENCODED(strinput) as Variant

	Dim i, j, checksum, mini, dummy, tableB as Integer
	Dim Code128 as String
	Dim Bytes(255) as Integer
	Dim idx as Integer
	
	' Barcode 128 character encoding 00=Â 01=! 02=" 03=# etc. 
	Dim C128CHARS as String
	C128CHARS = "Â!""#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~ÃÄÅÆÇÈÉÊËÌÍÎ"
	'           ---^  escape the " character
	'Â!"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~ÃÄÅÆÇÈÉÊËÌÍÎ

	Code128 = ""
	strinput = Trim(strinput)

	' not an empty string
	If Len(strinput) > 0 Then
	
		' Verify if all characters are valid
		For i = 1 To Len(strinput)
			j = Asc(Mid(strinput, i, 1))
			If (j >= 32 And j <= 126) Or (j = 203) Then
			Else
				i = 0
				Exit For
			End If
		Next

		' Calculation of the code string with optimized use of tables B and C
		Code128 = ""
		tableB = True
		If (i > 0) Then
			' prepare byte array, set bytes array all to -1
			For idx = 0 to 255
				Bytes(idx) = 0
			Next 

			' initialise byte array index, and string character index
			idx = 0
			i = 1

			' process all input string characters
			Do While i <= Len(strinput)
				If tableB Then
					' See if interesting to switch to table C
					' It is only worth switching to table C if next 4 characters are digits
					If (i + 4-1 <= Len(strinput) And ONLY_DIGITS(Mid(strinput, i, 4))) Then ' Choice of table C
						If i = 1 Then ' Starting with table C
							Bytes(idx) = 105 ' Start Code C
						Else ' Switch to table C
							Bytes(idx) = 99 ' Code C
						End If
						idx = idx + 1
						tableB = False
					Else
						If i = 1 Then
							Bytes(idx) = 104 ' Starting with table B
							idx = idx + 1
						End If
					End If
				End If
				If Not tableB Then
					' We are on table C, try to process 2 digits
					mini = 2
					If (i+mini-1 <= Len(strinput) And ONLY_DIGITS(Mid(strinput, i, mini))) Then ' OK for 2 digits, process it
						dummy = CInt(Mid(strinput, i, 2)) ' take next 2 digits
						i = i + 2
					Else ' We haven't 2 digits, switch to table B
						dummy = 100 ' Switch to table B
						tableB = True
					End If
					Bytes(idx) = dummy
					idx = idx + 1
				End If
				If tableB Then
					' Process 1 digit with table B
					Bytes(idx) = Asc(Mid(strinput, i, 1)) - 32
					idx = idx + 1
					i = i + 1
				End If
			Loop

			' calculation of the checksum
			For i = 0 To idx-1
				dummy = Bytes(i)
				If i = 0 Then checksum = dummy
				checksum = (checksum + (i * dummy)) Mod 103
			Next
			' add checksum also to output
			Bytes(idx) = checksum
			idx = idx + 1
			' add STOP character
			Bytes(idx) = 106
			idx = idx + 1
			
			' convert bytes to barcode 128 string
			Code128 = ""
			For i = 0 To idx-1
				dummy = Bytes(i)
				Code128 = Code128 & Mid(C128CHARS, dummy+1, 1)
			Next
		End If

	End If

	BARCODE128_ENCODED = Code128

End Function

Function ONLY_DIGITS(str as String) as Boolean
	' check if str contains only digits
	Dim ret as Boolean
	Dim i, a as Integer
	ret = True
	i = 1
	Do While (i <= len(str))
		' check for characters other than digits
		a = Asc(Mid(str, i, 1)) 
		If (a < 48 Or a > 57) Then ' < "0" or > "9"
			ret = False 
			Exit Do
		End If
		i = i + 1
	Loop
	ONLY_DIGITS = ret
End Function
