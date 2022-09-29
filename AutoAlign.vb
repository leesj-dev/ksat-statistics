Private Sub Worksheet_Activate()
    With Application
        .StatusBar = "In Progress..."
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .ErrorCheckingOptions.OmittedCells = False
    End With
    
    sh1 = "Input"
    sh2 = "Intermediate"
    sh3 = "Output"
    
    Sheets(sh2).UsedRange.Clear
    Sheets(sh3).UsedRange.Clear
    
    Dim rng As Range
    Dim rngStart As Range
    Dim rngEnd As Range
    Dim T_score As Range
    Dim cut As Range
    Dim ncut As Range
    
    With Sheets(sh1)
        .Range("A1:DB162").Copy Destination:=Sheets(sh2).Range("A1")
    End With
    
    With Sheets(sh2)
        For m = 1 To 3
            For n = 1 To 7
                If m = 3 And n = 4 Then
                    flg = True
                    Exit For
                End If
                
                Set T_score = .Range(.Cells(55 * m - 51, 11 * n + 21), .Cells(55 * m - 3, 11 * n + 21))
                For Each rng In T_score
                    If rng.HasFormula Then
                        rng.Formula = rng.Value
                    End If
                    rng = Round(rng, 0)
                Next rng
                
                For c = 55 * m - 51 To 55 * m - 3
                    missing = .Cells(c, 11 * n + 21).Value
                    If .Range(.Cells(55 * m - 51, 11 * n + 23), .Cells(55 * m - 3, 11 * n + 23)).Find(missing, LookIn:=xlValues) Is Nothing Then
                        s = 0
                        If .Cells(55 * m - 51, 11 * n + 23).Value > missing Then
                            Do Until .Cells(55 * m - 51 + s, 11 * n + 23).Value > missing And missing > .Cells(55 * m - 50 + s, 11 * n + 23).Value
                                s = s + 1
                            Loop
                            s = s + 1
                        End If
                        
                        Application.Calculation = xlCalculationAutomatic
                        .Range(.Cells(55 * m - 51 + s, 11 * n + 23), .Cells(55 * m - 4, 11 * n + 23)).Copy Destination:=.Cells(55 * m - 50 + s, 11 * n + 23)
                        .Range(.Cells(55 * m - 51 + s, 11 * n + 26), .Cells(55 * m - 4, 11 * n + 26)).Copy Destination:=.Cells(55 * m - 50 + s, 11 * n + 26)
                        .Cells(55 * m - 51 + s, 11 * n + 23).Value = missing
                        .Cells(55 * m - 51 + s, 11 * n + 26).Value = 0
                        Application.Calculation = xlCalculationManual
                    End If
                Next c
            Next n
            If flg = True Then
                Exit For
            End If
        Next m
        
        .Range("A7:D26").Copy Destination:=Sheets(sh3).Range("A1")
        .Range("F1:N109").Copy Destination:=Sheets(sh3).Range("F1")
        .Range("O1:AC150").Copy
    End With
    
    With Sheets(sh3)
        With .Range("O1")
            .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        End With
        
        .Range("C4:D20").Interior.ColorIndex = 0
        .Range("O4:AC150").Interior.ColorIndex = 0
        
        For k = 1 To 2
            For d = 4 To 150
                If .Cells(d, 8 * k + 8) <> .Cells(d + 1, 8 * k + 8) Then
                    Set ncut = .Range(.Cells(5, 3 * k + 4), .Cells(13, 3 * k + 4)).Find(.Cells(d, 8 * k + 8).Value, LookIn:=xlValues)
                    .Cells(ncut.Row, ncut.Column - 1).Value = .Cells(d, 8 * k + 7).Value
                    .Range(.Cells(d, 8 * k + 7), .Cells(d, 8 * k + 13)).Interior.ColorIndex = 44
                End If
                
                If .Cells(d, 8 * k + 13).Value = 1 Then
                    .Cells(d, 8 * k + 13).Select
                    .Application.CommandBars.FindControl(ID:=399).Execute
                    .Application.CommandBars.FindControl(ID:=399).Execute
                End If
                
                If IsEmpty(.Cells(d, 8 * k + 7)) Then
                    .Range(.Cells(d, 8 * k + 7), .Cells(d, 8 * k + 13)).Clear
                End If
            Next d
        Next k
    End With
    
    flg = False
    
    For m = 1 To 3
        For n = 1 To 7
            If m = 3 And n = 4 Then
                flg = True
                Exit For
            End If
            
            With Sheets(sh2)
                .Cells(55 * m - 54, 11 * n + 20).Copy Destination:=Sheets(sh3).Cells(55 * m - 54, 9 * n + 22)
                .Range(.Cells(55 * m - 54, 11 * n + 24), .Cells(55 * m - 54, 11 * n + 29)).Copy Destination:=Sheets(sh3).Cells(55 * m - 54, 9 * n + 24)
                .Range(.Cells(55 * m - 52, 11 * n + 20), .Cells(55 * m - 3, 11 * n + 21)).Copy Destination:=Sheets(sh3).Cells(55 * m - 52, 9 * n + 22)
                .Range(.Cells(55 * m - 52, 11 * n + 23), .Cells(55 * m - 52, 11 * n + 29)).Copy Destination:=Sheets(sh3).Cells(55 * m - 52, 9 * n + 23)
            End With
            
            With Sheets(sh3)
                .Cells(55 * m - 54, 9 * n + 25).Formula = .Cells(55 * m - 54, 9 * n + 25).Value
                .Cells(55 * m - 54, 9 * n + 27).Formula = .Cells(55 * m - 54, 9 * n + 27).Value
                .Cells(55 * m - 54, 9 * n + 29).Formula = .Cells(55 * m - 54, 9 * n + 29).Value
                
                Set T_score = .Range(.Cells(55 * m - 51, 9 * n + 23), .Cells(55 * m - 3, 9 * n + 23))
                For Each rng In T_score
                    If rng.HasFormula Then
                        rng.Formula = rng.Value
                    End If
                    rng = Round(rng, 0)
                Next rng
                
                Set rngStart = T_score.Cells(1, 1)
                For i = 1 To T_score.Rows.Count
                    If T_score.Cells(i, 1) <> T_score.Cells(i + 1, 1) Then
                        Set rngEnd = T_score.Cells(i, 1)
                        Application.DisplayAlerts = False
                        Range(rngStart, rngEnd).Merge
                        Application.DisplayAlerts = True
                        Set rngStart = T_score.Cells(i + 1, 1)
                    End If
                Next i
                
                .Range(.Cells(55 * m - 54, 9 * n + 22), Cells(55 * m - 54, 9 * n + 23)).Merge
                b = 0
                
                For a = 55 * m - 51 To 55 * m - 3
                    If .Cells(a, 9 * n + 23) = Empty Then
                        b = b + 1
                        For x = 9 * n + 23 To 9 * n + 29
                            With .Range(.Cells(a - 1, x), .Cells(a, x))
                                .Merge
                            End With
                        Next x
                    Else
                        With Sheets(sh2)
                            .Range(.Cells(a - b, 11 * n + 24), .Cells(a - b, 11 * n + 29)).Copy
                        End With
                        
                        With .Cells(a, 9 * n + 24)
                            .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                            .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                        End With
                    End If
                Next a
                
                .Range(.Cells(55 * m - 51, 9 * n + 22), Cells(55 * m - 3, 9 * n + 29)).Interior.Color = xlNone
                
                For c = 55 * m - 51 To 55 * m - 3
                    If Sheets(sh2).Cells(c, 11 * n + 24) <> Sheets(sh2).Cells(c + 1, 11 * n + 24) Then
                        Set cut = .Range(.Cells(55 * m - 51, 9 * n + 23), .Cells(55 * m - 3, 9 * n + 23)).Find(Sheets(sh2).Cells(c, 11 * n + 23).Value, LookIn:=xlValues)
                        
                        If .Cells(cut.Row, 9 * n + 26).Value = 0 Then
                            u = 1
                            Do While .Cells(cut.Row - u, 9 * n + 26).Value = 0
                                u = u + 1
                            Loop
                        Else
                            u = 0
                        End If
                        
                        .Range(.Cells(cut.Row - u, 9 * n + 23), Cells(cut.Row - u, 9 * n + 29)).Interior.ColorIndex = 44
                        j = 0
                        
                        Do While .Cells(cut.Row + j - u + 1, 9 * n + 23) = Empty And .Cells(cut.Row + j - u + 1, 9 * n + 22) <> Empty
                            j = j + 1
                        Loop
                        
                        If n < 4 Then
                            Set ncut = .Range(.Cells(12 * m + 17, 3 * n + 4), .Cells(12 * m + 25, 3 * n + 4)).Find(.Cells(cut.Row - u, 9 * n + 24).Value, LookIn:=xlValues)
                            .Cells(ncut.Row, ncut.Column - 1).Value = .Cells(cut.Row + j - u, 9 * n + 22).Value
                        Else
                            Set ncut = .Range(.Cells(12 * n + 17, 3 * m + 4), .Cells(12 * n + 25, 3 * m + 4)).Find(.Cells(cut.Row - u, 9 * n + 24).Value, LookIn:=xlValues)
                            .Cells(ncut.Row, ncut.Column - 1).Value = .Cells(cut.Row + j - u, 9 * n + 22).Value
                        End If

                        .Cells(cut.Row + j - u, 9 * n + 22).Interior.ColorIndex = 44
                    End If
                    
                    If .Cells(c, 9 * n + 26).Value = 0 Then
                        .Range(.Cells(c, 9 * n + 23), .Cells(c, 9 * n + 29)).Interior.ColorIndex = 1
                    End If
                    
                    If .Cells(c, 9 * n + 29).Value = 1 Then
                       .Cells(c, 9 * n + 29).Select
                       .Application.CommandBars.FindControl(ID:=399).Execute
                       .Application.CommandBars.FindControl(ID:=399).Execute
                    End If
                Next c
                
                .Range(.Cells(55 * m - 54, 9 * n + 22), Cells(55 * m - 54, 9 * n + 29)).Borders.LineStyle = 1
                .Range(.Cells(55 * m - 52, 9 * n + 22), Cells(55 * m - 3, 9 * n + 29)).Borders.LineStyle = 1
                
            End With
        Next n
        If flg = True Then
            Exit For
        End If
    Next m

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .StatusBar = False
    End With
    
End Sub
