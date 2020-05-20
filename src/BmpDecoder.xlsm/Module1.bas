Attribute VB_Name = "Module1"
Sub BmpDecoder()

    ActiveSheet.Copy After:=Worksheets(Worksheets.Count)
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    
    Dim file_path As Variant
    Dim file_len As Long
    Dim file_id As Integer
    Dim file_data() As Byte
    
    'ファイルヘッダ
    Dim bfType As String
    Dim bfSize As Long
'    Dim bfReserved1 As Integer
'    Dim bfReserved2 As Integer
    Dim bfOffBits As Long
    'CORE(bc) / V5(bV5) 情報ヘッダ
    Dim Size As Integer
    Dim width As Long
    Dim height As Long
    Dim Planes As Integer
    Dim BitCount As Integer
    'V5情報ヘッダ
    Dim bV5Compression As Integer
'    Dim bV5SizeImage As Long
'    Dim bV5XPelsPerMeter As Long
'    Dim bV5YPelsPerMeter As Long
    Dim bV5ClrUsed As Integer
'    Dim bV5ClrImportant
    Dim bV5RedMask As String
    Dim bV5GreenMask As String
    Dim bV5BlueMask As String
    Dim bV5AlphaMask As String
'    Dim bV5CSType
'    Dim bV5Endpoints
'    Dim bV5GammaRed
'    Dim bV5GammaGreen
'    Dim bV5GammaBlue
'    Dim bV5Intent
'    Dim bV5ProfileData
'    Dim bV5ProfileSize
'    Dim bV5Reserved
    'TRIPLE(rgbt) / QUAD(rgb) カラーパレット
    Dim Blue As String
    Dim Green As String
    Dim Red As String
    'QUADカラーパレット
    Dim rgbReserved As String
    
    Dim offset As Long
    offset = 0
    
    Dim num_color As Long
    Dim palette() As String
    Dim line_size As Long
    Dim spacer As Integer
    Dim loading As String
    Dim flag As Boolean
    Dim num_mask As Integer
    
    Dim temp_mask() As String
    '(1,1) : R
    '(1,2) : R始点
    '(1,3) : R終点
    Dim temp_bin As String
    
    Dim temp_R As String
    Dim temp_G As String
    Dim temp_B As String
    Dim temp_A As String
    
    Dim drawing() As String
    
    'ファイルの読み込み
    file_path = Application.GetOpenFilename(",*.bmp")
    'キャンセル
    If file_path = False Then
        Exit Sub
    End If
    'ファイルサイズ0
    file_len = filelen(file_path)
    If file_len = 0 Then
        Exit Sub
    End If
    
    file_id = FreeFile
    Open file_path For Binary As #file_id
    
    ReDim file_data(0 To file_len - 1)
    
    Get #file_id, , file_data
    Close #file_id
    
    
    'ファイルヘッダの検証
    bfType = Chr(file_data(offset)) & Chr(file_data(offset + 1))
    bfSize = Byte2Dec(file_data, offset + 2, 4)
'    bfReserved1 = Byte2Dec(file_data, offset + 6, 2)
'    bfReserved2 = Byte2Dec(file_data, offset + 8, 2)
    bfOffBits = Byte2Dec(file_data, offset + 10, 2)
    
    If bfType <> "BM" Then
        MsgBox "不正なファイルヘッダ : BMPファイルではありません。"
        Exit Sub
    End If
    offset = offset + 14
    
    
    '情報ヘッダの検証
    Size = Byte2Dec(file_data, offset, 4)
    
    If Size = 12 Then 'COREタイプ
        width = Byte2Dec(file_data, offset + 4, 2)
        height = Byte2Dec(file_data, offset + 6, 2)
        Planes = Byte2Dec(file_data, offset + 8, 2)
        BitCount = Byte2Dec(file_data, offset + 10, 2)
        
        If width < 1 Or height < 1 Then
            MsgBox "不正な情報ヘッダ : 幅 / 高さ"
            Exit Sub
        ElseIf Planes <> 1 Then
            MsgBox "不正な情報ヘッダ : プレーン数"
            Exit Sub
        ElseIf BitCount <> 1 And BitCount <> 4 And BitCount <> 8 And BitCount <> 24 Then
            MsgBox "不正な情報ヘッダ : ピクセル毎のビット数"
            Exit Sub
        End If
        
        num_color = 2 ^ BitCount
        
    ElseIf Size = 40 Or Size = 52 Or Size = 56 Or Size = 60 Or Size = 96 Or Size = 108 Or Size = 112 Or Size = 120 Or Size = 124 Then 'V5タイプ
        width = Byte2Dec(file_data, offset + 4, 4)
        height = Byte2Dec(file_data, offset + 8, 4)
        Planes = Byte2Dec(file_data, offset + 12, 2)
        BitCount = Byte2Dec(file_data, offset + 14, 2)
        
        bV5Compression = Byte2Dec(file_data, offset + 16, 4)
'        bV5SizeImage = Byte2Dec(file_data, offset + 20, 4)
'        bV5XPelsPerMeter = Byte2Dec(file_data, offset + 24, 4)
'        bV5YPelsPerMeter = Byte2Dec(file_data, offset + 28, 4)
        bV5ClrUsed = Byte2Dec(file_data, offset + 32, 4)
'        bV5ClrImportant = Byte2Dec(file_data, offset + 36, 4)
        
        If Planes <> 1 Then
            MsgBox "不正な情報ヘッダ : プレーン数"
            Exit Sub
        ElseIf width <= 0 Or height = 0 Then
            MsgBox "不正な情報ヘッダ : 幅 / 高さ"
            Exit Sub
        ElseIf bV5Compression <> 0 And bV5Compression <> 3 Then
            MsgBox "圧縮されたBMPファイルには対応していません。"
            Exit Sub
        ElseIf BitCount <> 0 And BitCount <> 1 And BitCount <> 4 And BitCount <> 8 And BitCount <> 16 And BitCount <> 24 And BitCount <> 32 Then
            MsgBox "不正な情報ヘッダ : ピクセル毎のビット数"
            Exit Sub
        End If
        
        If bV5ClrUsed = 0 And BitCount <= 8 Then
            num_color = 2 ^ BitCount
        Else
            num_color = bV5ClrUsed
        End If
        
        If Size >= 52 And bV5Compression = 3 Then
'            bV5RedMask = Byte2Dec(file_data, offset + 40, 4)
            bV5RedMask = Byte2Hex(file_data(offset + 43)) & Byte2Hex(file_data(offset + 42)) & Byte2Hex(file_data(offset + 41)) & Byte2Hex(file_data(offset + 40))
            bV5GreenMask = Byte2Hex(file_data(offset + 47)) & Byte2Hex(file_data(offset + 46)) & Byte2Hex(file_data(offset + 45)) & Byte2Hex(file_data(offset + 44))
            bV5BlueMask = Byte2Hex(file_data(offset + 51)) & Byte2Hex(file_data(offset + 50)) & Byte2Hex(file_data(offset + 49)) & Byte2Hex(file_data(offset + 48))
            
            num_mask = 3
            
            If Size >= 56 Then
                bV5AlphaMask = Byte2Hex(file_data(offset + 55)) & Byte2Hex(file_data(offset + 54)) & Byte2Hex(file_data(offset + 53)) & Byte2Hex(file_data(offset + 52))
                
                num_mask = 4
                
'                If Size >= 60 Then
'                    bV5CSType = Byte2Dec(file_data, offset + 56, 4)
'
'                    If Size >= 96 Then
'                        bV5Endpoints = Byte2Dec(file_data, offset + 60, 36)
'
'                        If Size >= 108 Then
'                            bV5GammaRed = Byte2Dec(file_data, offset + 96, 4)
'                            bV5GammaGreen = Byte2Dec(file_data, offset + 100, 4)
'                            bV5GammaBlue = Byte2Dec(file_data, offset + 104, 4)
'
'                            If Size >= 112 Then
'                                bV5Intent = Byte2Dec(file_data, offset + 108, 4)
'
'                                If Size >= 120 Then
'                                    bV5ProfileData = Byte2Dec(file_data, offset + 112, 4)
'                                    bV5ProfileSize = Byte2Dec(file_data, offset + 116, 4)
'
'                                    If Size >= 124 Then
'                                        bV5Reserved = Byte2Dec(file_data, offset + 120, 4)
'
'                                    End If
'                                End If
'                            End If
'                        End If
'                    End If
'                End If
            End If
        End If
        
    Else
        MsgBox "不正な情報ヘッダ : ヘッダサイズ"
        Exit Sub
        
    End If
    
    offset = offset + Size
    
    'ビットフィールド
    If Size = 40 And (BitCount = 16 Or BitCount = 32) And bV5Compression = 3 Then
        bV5RedMask = Byte2Hex(file_data(offset + 3)) & Byte2Hex(file_data(offset + 2)) & Byte2Hex(file_data(offset + 1)) & Byte2Hex(file_data(offset))
        bV5GreenMask = Byte2Hex(file_data(offset + 7)) & Byte2Hex(file_data(offset + 6)) & Byte2Hex(file_data(offset + 5)) & Byte2Hex(file_data(offset + 4))
        bV5BlueMask = Byte2Hex(file_data(offset + 11)) & Byte2Hex(file_data(offset + 10)) & Byte2Hex(file_data(offset + 9)) & Byte2Hex(file_data(offset + 8))
        
        num_mask = 3
        
        offset = offset + 12
    End If
    
    'カラーパレット
    If (BitCount = 1 Or BitCount = 4 Or BitCount = 8) Or bV5ClrUsed >= 1 Then
        If Size = 12 Then 'TRIPLEタイプ
            ReDim palette(0 To num_color - 1)
            For i = 0 To num_color - 1
                palette(i) = Byte2RGB(file_data, offset + (i * 3)) & "FF"
            Next i
            offset = offset + (num_color * 3)
        Else 'QUADタイプ
            ReDim palette(0 To num_color - 1)
            For i = 0 To num_color - 1
                palette(i) = Byte2RGB(file_data, offset + (i * 4)) & "FF"
            Next i
            offset = offset + (num_color * 4)
        End If
    End If
    
    
    '描画情報をdrawingに格納
    
    If bfOffBits <> 0 And offset < bfOffBits Then
        offset = bfOffBits
    End If
    
    ReDim drawing(height, width)
    
    line_size = Int((width * BitCount + 31) / 32) * 4
    
    If BitCount = 1 Or BitCount = 4 Or BitCount = 8 Then
        'パレットタイプ
        Select Case BitCount
            Case 1
                For i = offset To file_len - 1
                    loading = loading & Hex2Bin(Byte2Hex(file_data(i)))
                Next i
                
                For i = height To 1 Step -1
                    For j = 1 To width
                        drawing(i, j) = palette(Mid(loading, j, 1))
                    Next j

                    loading = Mid(loading, (line_size * (8 / BitCount)) + 1)

                Next i
                
            Case 4
                For i = height To 1 Step -1
                    Select Case width Mod 4
                        Case 0, 2
                            For j = 1 To width
                                drawing(i, j) = palette(Hex2Dec(Left(Byte2Hex(file_data(offset)), 1)))
                                j = j + 1
                                drawing(i, j) = palette(Hex2Dec(Right(Byte2Hex(file_data(offset)), 1)))
                                offset = offset + 1
                            Next
                            offset = offset + (1.5 * (width Mod 4))
                        Case 1, 3
                            For j = 1 To width - 1
                                drawing(i, j) = palette(Hex2Dec(Left(Byte2Hex(file_data(offset)), 1)))
                                j = j + 1
                                drawing(i, j) = palette(Hex2Dec(Right(Byte2Hex(file_data(offset)), 1)))
                                offset = offset + 1
                            Next j
                            drawing(i, j) = palette(Hex2Dec(Left(Byte2Hex(file_data(offset)), 1)))
                            offset = offset + (-0.5 * (width Mod 4) + 4.5)

                    End Select
                Next i
                
            Case 8
                For i = height To 1 Step -1
                    For j = 1 To width
                        drawing(i, j) = palette(Byte2Dec(file_data, offset, 1))
                        offset = offset + 1
                    Next j

                    offset = offset + 4 - (width Mod 4)

                Next i
                
        End Select
        
    ElseIf BitCount = 16 Or BitCount = 32 Then
        'ビットフィールドタイプ
        
        '規定ビットフィールド
        If bV5Compression = 0 Then
            Select Case BitCount
                Case 16
                    bV5RedMask = "00007C00"
                    bV5GreenMask = "000003E0"
                    bV5BlueMask = "0000001F"
                    
                    num_mask = 3
'                    If Size >= 56 Then
'                        bV5AlphaMask = "00008000"
'                    End If
                Case 32
                    bV5RedMask = "00FF0000"
                    bV5GreenMask = "0000FF00"
                    bV5BlueMask = "000000FF"
                    
                    num_mask = 3
'                    If Size >= 56 Then
'                        bV5AlphaMask = "FF000000"
'                    End If
            End Select
        End If
        
        ReDim temp_mask(num_mask, 3)
        temp_mask(1, 1) = bV5RedMask
        temp_mask(2, 1) = bV5GreenMask
        temp_mask(3, 1) = bV5BlueMask
        If num_mask = 4 Then
            temp_mask(4, 1) = bV5AlphaMask
        End If
        For i = 1 To num_mask
            temp_mask(i, 2) = 1
            temp_mask(i, 3) = 1
        Next i
        
        Select Case BitCount
            Case 16
                For i = 1 To num_mask
                    If InStr(17, Hex2Bin(temp_mask(i, 1)), "1") <> 0 Then
                        Do While Mid(Right(Hex2Bin(temp_mask(i, 1)), 16), temp_mask(i, 2), 1) = "0"
                            temp_mask(i, 2) = CInt(temp_mask(i, 2)) + 1
                        Loop
                        Do While Mid(Right(Hex2Bin(temp_mask(i, 1)), 16), CInt(temp_mask(i, 2)) + CInt(temp_mask(i, 3)), 1) = "1"
                            temp_mask(i, 3) = CInt(temp_mask(i, 3)) + 1
                        Loop
                        temp_mask(i, 3) = CInt(temp_mask(i, 2)) + CInt(temp_mask(i, 3)) - 1
                        If InStr(CInt(temp_mask(i, 3)) + 1, Right(Hex2Bin(temp_mask(i, 1)), 16), "1") <> 0 Then
                            MsgBox "不正なビットフィールドです。"
                            Exit Sub
                        End If
                    End If
                Next i
                
                
                For i = height To 1 Step -1
                    For j = 1 To width
                        temp_bin = Hex2Bin(Byte2Hex(file_data(offset + 1)) & Byte2Hex(file_data(offset)))
                        
                        For k = 1 To num_mask
                            drawing(i, j) = drawing(i, j) & Right("0" & hex(Application.WorksheetFunction.Round(Bin2Dec(Mid(temp_bin, temp_mask(k, 2), CInt(temp_mask(k, 3)) - CInt(temp_mask(k, 2)) + 1)) / (2 ^ (CInt(temp_mask(k, 3)) - CInt(temp_mask(k, 2)) + 1) - 1) * 255, 0)), 2)
                        Next k
                        
                        If num_mask = 3 Then
                            drawing(i, j) = drawing(i, j) & "FF"
                        End If
                        
                        offset = offset + 2
                    Next j
                    
                    If width Mod 4 <> 0 Then
                        offset = offset + 2
                    End If
                    
                Next i
                
            Case 32
                For i = 1 To num_mask
                    If InStr(1, Hex2Bin(temp_mask(i, 1)), "1") <> 0 Then
                        Do While Mid(Hex2Bin(temp_mask(i, 1)), temp_mask(i, 2), 1) = "0"
                            temp_mask(i, 2) = CInt(temp_mask(i, 2)) + 1
                        Loop
                        Do While Mid(Hex2Bin(temp_mask(i, 1)), CInt(temp_mask(i, 2)) + CInt(temp_mask(i, 3)), 1) = "1"
                            temp_mask(i, 3) = CInt(temp_mask(i, 3)) + 1
                        Loop
                        temp_mask(i, 3) = CInt(temp_mask(i, 2)) + CInt(temp_mask(i, 3)) - 1
                        If InStr(CInt(temp_mask(i, 3)) + 1, Right(Hex2Bin(temp_mask(i, 1)), 16), "1") <> 0 Then
                            MsgBox "不正なビットフィールドです。"
                            Exit Sub
                        End If
                    End If
                Next i
                
                For i = height To 1 Step -1
                    For j = 1 To width
                        temp_bin = Hex2Bin(Byte2Hex(file_data(offset + 3)) & Byte2Hex(file_data(offset + 2)) & Byte2Hex(file_data(offset + 1)) & Byte2Hex(file_data(offset)))
                        
                        For k = 1 To num_mask
                            drawing(i, j) = drawing(i, j) & Right("0" & hex(Application.WorksheetFunction.Round(Bin2Dec(Mid(temp_bin, temp_mask(k, 2), CInt(temp_mask(k, 3)) - CInt(temp_mask(k, 2)) + 1)) / (2 ^ (CInt(temp_mask(k, 3)) - CInt(temp_mask(k, 2)) + 1) - 1) * 255, 0)), 2)
                        Next k
                        
                        If num_mask = 3 Then
                            drawing(i, j) = drawing(i, j) & "FF"
                        End If
                        
                        offset = offset + 4
                    Next j
                    
                Next i
                
        End Select
        
    Else
        'カラータイプ
        For i = height To 1 Step -1
            For j = 1 To width
                drawing(i, j) = Byte2RGB(file_data, offset) & "FF"
                offset = offset + 3
            Next j
            
            offset = offset + (width Mod 4)
            
        Next i
    End If
    
    '描画
    Application.ScreenUpdating = False

    Cells.clear

    Cells.RowHeight = 1 * 0.75      '3      '5  '8
    Cells.ColumnWidth = 0.4 * 0.118   '1.5    '3  '5

    For i = 1 To height
        For j = 1 To width
            
            temp_R = Left(drawing(i, j), 2)
            temp_G = Mid(drawing(i, j), 3, 2)
            temp_B = Mid(drawing(i, j), 5, 2)
            temp_A = Right(drawing(i, j), 2)
            
            If temp_A <> "00" Then
                Cells(i, j).Interior.Color = Val("&H" & temp_A & temp_B & temp_G & temp_R)
                Cells(i, j).Value = drawing(i, j)
            End If
            
            limit = limit + 1
            
        Next j
    Next i
    
    Cells.NumberFormatLocal = ";;;"
    
    Application.ScreenUpdating = True
    
End Sub

Sub clear()
    Cells.RowHeight = 1 * 0.75        '3      '5  '8
    Cells.ColumnWidth = 0.4 * 0.118   '1.5    '3  '5
    Cells.clear
End Sub

Sub delete_style()
'デバッグ用
    On Error Resume Next

    Dim M()

    j = ActiveWorkbook.Styles.Count
    ReDim M(j)
    
    For i = 1 To j
        M(i) = ActiveWorkbook.Styles(i).Name
    Next
    For i = 1 To j
        If InStr("Hyperlink,Normal,Followed Hyperlink", _
                    M(i)) = 0 Then
            ActiveWorkbook.Styles(M(i)).Delete
        End If
    Next

End Sub
