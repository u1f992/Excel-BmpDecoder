Attribute VB_Name = "Module2"
Function Byte2Hex(data As Byte) As String
    
    Byte2Hex = Right("0" & hex(data), 2)
    
End Function

Function Byte2Dec(ByRef data() As Byte, start As Long, length As Long) As Long
    
    Dim i As Long
    Dim temp As String
    
    For i = start To start + length - 1
        temp = Byte2Hex(data(i)) & temp
    Next i
    
    Byte2Dec = Val("&H" & temp)
    
End Function

Function Byte2RGB(ByRef data() As Byte, start As Long) As String
    
    Dim i As Long
    
    For i = start To start + 3 - 1
        Byte2RGB = Byte2Hex(data(i)) & Byte2RGB
    Next i
    
End Function

Function Byte2RGBA(ByRef data() As Byte, start As Long) As String
    
    Dim i As Long
    
    For i = start + 1 To start + 4 - 1
        Byte2RGBA = Byte2Hex(data(i)) & Byte2RGBA
    Next i
    
    Byte2RGBA = Left(Byte2RGBA, 6) & hex(255 - Right(Byte2RGBA, 2))
    
End Function

Function Hex2Bin(hex As String) As String
    
    For i = 1 To Len(hex)
        Select Case Mid(hex, i, 1)
            Case "0"
                Hex2Bin = Hex2Bin & "0000"
            Case "1"
                Hex2Bin = Hex2Bin & "0001"
            Case "2"
                Hex2Bin = Hex2Bin & "0010"
            Case "3"
                Hex2Bin = Hex2Bin & "0011"
            Case "4"
                Hex2Bin = Hex2Bin & "0100"
            Case "5"
                Hex2Bin = Hex2Bin & "0101"
            Case "6"
                Hex2Bin = Hex2Bin & "0110"
            Case "7"
                Hex2Bin = Hex2Bin & "0111"
            Case "8"
                Hex2Bin = Hex2Bin & "1000"
            Case "9"
                Hex2Bin = Hex2Bin & "1001"
            Case "A"
                Hex2Bin = Hex2Bin & "1010"
            Case "B"
                Hex2Bin = Hex2Bin & "1011"
            Case "C"
                Hex2Bin = Hex2Bin & "1100"
            Case "D"
                Hex2Bin = Hex2Bin & "1101"
            Case "E"
                Hex2Bin = Hex2Bin & "1110"
            Case "F"
                Hex2Bin = Hex2Bin & "1111"
        End Select
    Next i
    
End Function

Function Hex2Dec(hex As String) As Long
    
    For i = 1 To Len(hex)
        Select Case Mid(hex, i, 1)
            Case "0"
                Hex2Dec = Hex2Dec + 0 * (16 ^ (Len(hex) - i))
            Case "1"
                Hex2Dec = Hex2Dec + 1 * (16 ^ (Len(hex) - i))
            Case "2"
                Hex2Dec = Hex2Dec + 2 * (16 ^ (Len(hex) - i))
            Case "3"
                Hex2Dec = Hex2Dec + 3 * (16 ^ (Len(hex) - i))
            Case "4"
                Hex2Dec = Hex2Dec + 4 * (16 ^ (Len(hex) - i))
            Case "5"
                Hex2Dec = Hex2Dec + 5 * (16 ^ (Len(hex) - i))
            Case "6"
                Hex2Dec = Hex2Dec + 6 * (16 ^ (Len(hex) - i))
            Case "7"
                Hex2Dec = Hex2Dec + 7 * (16 ^ (Len(hex) - i))
            Case "8"
                Hex2Dec = Hex2Dec + 8 * (16 ^ (Len(hex) - i))
            Case "9"
                Hex2Dec = Hex2Dec + 9 * (16 ^ (Len(hex) - i))
            Case "A"
                Hex2Dec = Hex2Dec + 10 * (16 ^ (Len(hex) - i))
            Case "B"
                Hex2Dec = Hex2Dec + 11 * (16 ^ (Len(hex) - i))
            Case "C"
                Hex2Dec = Hex2Dec + 12 * (16 ^ (Len(hex) - i))
            Case "D"
                Hex2Dec = Hex2Dec + 13 * (16 ^ (Len(hex) - i))
            Case "E"
                Hex2Dec = Hex2Dec + 14 * (16 ^ (Len(hex) - i))
            Case "F"
                Hex2Dec = Hex2Dec + 15 * (16 ^ (Len(hex) - i))
        End Select
    Next i
    
End Function

Function Byte2Bin(data As Byte) As String

    Byte2Bin = Hex2Bin(Byte2Hex(data))
    
End Function

Function Bin2Dec(bin As String) As Long
    
    For i = 1 To Len(bin)
        Bin2Dec = Bin2Dec + (Mid(bin, i, 1) * (2 ^ (Len(bin) - 1 - (i - 1))))
    Next i

End Function
