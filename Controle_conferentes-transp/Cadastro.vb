Sub grava_cadastro()
'
' grava_cadastro Macro
'

'
    Sheets("BD").Select
    Range("Tabela1[[#Headers],[DATA]]").Select
    
            If Range("E2").Value <> "" Then
            Selection.End(xlDown).Select
            End If
       
    Sheets("cADASTRO").Select
    Range("B7:N7").Select
    Selection.Copy
    Sheets("BD").Select
    
            ActiveCell.Offset(1, 0).Select
        
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
          
    Sheets("CADASTRO").Select
    Range("E13:I13").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("BD").Select
    
    ActiveCell.Offset(0, 13).Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
     
          
            Sheets("CADASTRO").Select
            Application.CutCopyMode = False
    
                Dim A As Range
                Set A = Planilha2.Range("O14")
    
                Dim B As String
                B = A
    
                If B = True Then

                    Range("C7").Select
                    Selection.ClearContents
                    Range("H7").Select
                    Selection.ClearContents
                    Range("I7").Select
                    Selection.ClearContents
                    Range("J7").Select
                    Selection.ClearContents
                    Range("L7").Select
                    Selection.ClearContents
                    Range("M7").Select
                    Selection.ClearContents
                    Range("C7").Select
                        Else
                Range("C7").Select
                Selection.ClearContents
                Range("D7").Select
                Selection.ClearContents
                Range("G7").Select
                Selection.ClearContents
                Range("H7").Select
                Selection.ClearContents
                Range("I7").Select
                Selection.ClearContents
                Range("J7").Select
                Selection.ClearContents
                Range("L7").Select
                Selection.ClearContents
                Range("M7").Select
                Selection.ClearContents
                Range("N7").Select
                Selection.ClearContents
                Range("E13").Select
                Selection.ClearContents
                Range("F13").Select
                Selection.ClearContents
                Range("G13").Select
                Selection.ClearContents
                Range("H13").Select
                Selection.ClearContents
                Range("I13").Select
                Selection.ClearContents
        
            Range("B7").Select
    End If
    
    
End Sub