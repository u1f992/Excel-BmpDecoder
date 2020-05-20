Attribute VB_Name = "Module3"
Sub test_sort()
    
    ActiveSheet.Copy After:=Worksheets(Worksheets.Count)
    
    Dim start_x As Long
    Dim start_y As Long
    
    Dim width As Long
    Dim height As Long
    
    start_x = Selection(1).Column
    start_y = Selection(1).Row
    width = Selection.Columns.Count
    height = Selection.Rows.Count
    
    Application.ScreenUpdating = False
    
    For j = start_x To start_x + width
        Range(Cells(start_y, j), Cells(start_y + height, j)).Sort Key1:=Cells(start_y, j), order1:=xlAscending, Orientation:=xlTopToBottom
    Next j
        
    For j = start_y To start_y + height
        Range(Cells(j, start_x), Cells(j, start_x + width)).Sort Key1:=Cells(j, start_x), order1:=xlAscending, Orientation:=xlLeftToRight
    Next j
    
    Application.ScreenUpdating = True
    
End Sub

Sub test_HSV()

    ActiveSheet.Copy After:=Worksheets(Worksheets.Count)
    
    Dim start_x As Long
    Dim start_y As Long
    
    Dim width As Long
    Dim height As Long
    
    Dim temp As String
    
    Dim target_H As Long
    Dim target_S As Double
    Dim target_V As Double
    
    Dim judge As Variant
    
    Dim R As Integer
    Dim G As Integer
    Dim B As Integer
    Dim A As String
    
    Dim max As Integer
    Dim min As Integer
    
    Dim H As Integer
    Dim S As Integer
    Dim V As Integer
    
    start_x = Selection(1).Column
    start_y = Selection(1).Row
    width = Selection.Columns.Count
    height = Selection.Rows.Count
    
    Do
        temp = InputBox("色相の回転角度を入力してください。(-180~180)", Default:="0")
        'キャンセルはStrPtr関数
        If StrPtr(temp) = 0 Then
            Exit Sub
        ElseIf temp = "" Then
            judge = MsgBox("入力欄が空欄です。", Buttons:=5)
            If judge = 2 Then
                Exit Sub
            End If
        ElseIf Not IsNumeric(temp) Then
            judge = MsgBox("入力欄が不正です。", Buttons:=5)
            If judge = 2 Then
                Exit Sub
            End If
        Else
            target_H = CLng(temp)
            If target_H < -180 Or target_H > 180 Then
                judge = MsgBox("入力欄が不正です。", Buttons:=5)
                If judge = 2 Then
                    Exit Sub
                End If
            Else
                Exit Do
            End If
        End If
    Loop
    
    Do
        temp = InputBox("彩度の係数を入力してください。(0~10)", Default:="1")
        'キャンセルはStrPtr関数
        If StrPtr(temp) = 0 Then
            Exit Sub
        ElseIf temp = "" Then
            judge = MsgBox("入力欄が空欄です。", Buttons:=5)
            If judge = 2 Then
                Exit Sub
            End If
        ElseIf Not IsNumeric(temp) Then
            judge = MsgBox("入力欄が不正です。", Buttons:=5)
            If judge = 2 Then
                Exit Sub
            End If
        Else
            target_S = CDbl(temp)
            If target_S < 0 Or target_S > 10 Then
                judge = MsgBox("入力欄が不正です。", Buttons:=5)
                If judge = 2 Then
                    Exit Sub
                End If
            Else
                Exit Do
            End If
        End If
    Loop
    
    Do
        temp = InputBox("明度の係数を入力してください。(0~10)", Default:="1")
        'キャンセルはStrPtr関数
        If StrPtr(temp) = 0 Then
            Exit Sub
        ElseIf temp = "" Then
            judge = MsgBox("入力欄が空欄です。", Buttons:=5)
            If judge = 2 Then
                Exit Sub
            End If
        ElseIf Not IsNumeric(temp) Then
            judge = MsgBox("入力欄が不正です。", Buttons:=5)
            If judge = 2 Then
                Exit Sub
            End If
        Else
            target_V = CDbl(temp)
            If target_V < 0 Or target_S > 10 Then
                judge = MsgBox("入力欄が不正です。", Buttons:=5)
                If judge = 2 Then
                    Exit Sub
                End If
            Else
                Exit Do
            End If
        End If
    Loop
    
    Application.ScreenUpdating = False
    
    For i = start_y To start_y + height
        For j = start_x To start_x + width
            If Cells(i, j).Value <> "" Then
                R = Hex2Dec(Left(Cells(i, j).Value, 2))
                G = Hex2Dec(Mid(Cells(i, j).Value, 3, 2))
                B = Hex2Dec(Mid(Cells(i, j).Value, 5, 2))
                A = Right(Cells(i, j).Value, 2)
                
                max = Application.WorksheetFunction.max(R, G, B)
                min = Application.WorksheetFunction.min(R, G, B)
                
                'V、Sを求める
                V = max
                S = Application.WorksheetFunction.Round((max - min) / max * 255, 0)
                
                If max = min Then
                    H = 0
                Else
                    'Hを求める
                    If R = G = B Then
                        H = 0
                    ElseIf R >= G And R >= B Then
                        'Rが最大
                        H = 60 * ((G - B) / (max - min))
                    ElseIf G >= R And G >= B Then
                        'Gが最大
                        H = 60 * ((B - R) / (max - min)) + 120
                    ElseIf B >= R And B >= G Then
                        'Bが最大
                        H = 60 * ((R - G) / (max - min)) + 240
                    End If
        
                    If H < 0 Then
                        H = H + 360
                    ElseIf H > 360 Then
                        H = H - 360
                    End If
                End If
                
                '変化値の適用
                H = H + target_H
                If H < 0 Then
                    H = H + 360
                ElseIf H > 360 Then
                        H = H - 360
                End If
                
                
                If target_S < 1 Then
                    S = Application.WorksheetFunction.Round(S * target_S, 0)
                ElseIf target_S > 1 Then
                    S = S + Application.WorksheetFunction.Round((255 - S) / 10 * target_S, 0)
                End If
                
                If target_V < 1 Then
                    V = Application.WorksheetFunction.Round(V * target_V, 0)
                ElseIf target_V > 1 Then
                    V = V + Application.WorksheetFunction.Round((255 - V) / 10 * target_V, 0)
                End If
                
                'RGBに戻す
                max = V
                min = Application.WorksheetFunction.Round(max - ((S / 255) * max), 0)
                
                If H >= 0 And H < 60 Then
                    R = max
                    G = Application.WorksheetFunction.Round((H / 60) * (max - min), 0) + min
                    B = min
                ElseIf H >= 60 And H < 120 Then
                    R = Application.WorksheetFunction.Round(((120 - H) / 60) * (max - min), 0) + min
                    G = max
                    B = min
                ElseIf H >= 120 And H < 180 Then
                    R = min
                    G = max
                    B = Application.WorksheetFunction.Round(((H - 120) / 60) * (max - min), 0) + min
                ElseIf H >= 180 And H < 240 Then
                    R = min
                    G = Application.WorksheetFunction.Round(((240 - H) / 60) * (max - min), 0) + min
                    B = max
                ElseIf H >= 240 And H < 300 Then
                    R = Application.WorksheetFunction.Round(((H - 240) / 60) * (max - min), 0) + min
                    G = min
                    B = max
                ElseIf H >= 300 And H <= 360 Then
                    R = max
                    G = min
                    B = Application.WorksheetFunction.Round(((360 - H) / 60) * (max - min), 0) + min
                End If
                
                Cells(i, j).Interior.Color = Val("&H" & A & Right("0" & hex(B), 2) & Right("0" & hex(G), 2) & Right("0" & hex(R), 2))
                Cells(i, j).Value = Right("0" & hex(R), 2) & Right("0" & hex(G), 2) & Right("0" & hex(B), 2) & A
            End If
        Next j
    Next i
    
    Application.ScreenUpdating = True
    
End Sub

Sub test_rotation()
'飽きたので未完成
    ActiveSheet.Copy After:=Worksheets(Worksheets.Count)
    
    Dim i As Long
    Dim j As Long
    
    Dim start_x As Long
    Dim start_y As Long
    
    Dim width As Long
    Dim height As Long
    
    start_x = Selection(1).Column
    start_y = Selection(1).Row
    width = Selection.Columns.Count
    height = Selection.Rows.Count
    
    Dim buffer() As String
    ReDim buffer(width, height)
    
    'ここがおかしい
    For i = 1 To height
        For j = 1 To width
            buffer(j, i) = Cells(i, j).Value
        Next j
    Next i
    
    Application.ScreenUpdating = False
        
    Cells.clear
    
    For i = 1 To width
        For j = 1 To height
            
            temp_R = Left(buffer(i, j), 2)
            temp_G = Mid(buffer(i, j), 3, 2)
            temp_B = Mid(buffer(i, j), 5, 2)
            temp_A = Right(buffer(i, j), 2)
            
            If temp_A <> "00" Then
                Cells(i, j).Interior.Color = Val("&H" & temp_A & temp_B & temp_G & temp_R)
                Cells(i, j).Value = buffer(i, j)
            End If
            
            If Cells(i, j).Value = "" Then
                Cells(i, j).clear
            End If
        Next j
    Next i
    
    Cells.NumberFormatLocal = ";;;"
    
    Application.ScreenUpdating = True
    
End Sub
