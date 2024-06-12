Sub ConvertCSVsAndInsertDataFormulasWithFSO()
    On Error GoTo ErrorHandler
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim csvPath As String
    Dim newFolder As Object
    Dim ws As Worksheet
    Dim fd As FileDialog

    ' フォルダ選択ダイアログを表示
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "CSVファイルが格納されているフォルダを選択してください"
        If .Show = -1 Then
            csvPath = .SelectedItems(1) & "\"
        Else
            MsgBox "フォルダが選択されませんでした。"
            Exit Sub
        End If
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(csvPath)

    Dim filesProcessed As Integer
    filesProcessed = 0

    For Each file In folder.Files
        If Right(file.Name, 4) = ".csv" Then
            ' CSVファイルを開く
            Workbooks.Open Filename:=csvPath & file.Name
            Set ws = ActiveWorkbook.Sheets(1)
            ' 指定されたセルにデータと式を挿入
            With ws
                .Range("A19").Value = "V_max"
                .Range("B19").Formula = "=MAX(E:E)"
                .Range("A20").Value = "V_min"
                .Range("B20").Formula = "=MIN(E:E)"
                .Range("A21").Value = "I_max"
                .Range("A22").Value = "I_min"
                .Range("A23").Value = "V_phase"
                .Range("B23").Value = 4.5
                .Range("A24").Value = "I_phase"
                .Range("B24").Value = 3.5
                .Range("A25").Value = "H_v"
                .Range("A26").Value = "tanθ"
                .Range("B26").Formula = "=TAN(B23-B24)"
                .Range("A27").Value = "I/V"
                .Range("B27").Formula = "=B21/B19"
                .Range("A28").Value = "ω"
                .Range("B28").Formula = "=2*PI()*T1"
                .Range("A29").Value = "freq.[kHz]"
                .Range("B29").Formula = "=T1/1000"
                .Range("A30").Value = "(V_max+V_min)/2"
                .Range("B30").Formula = "=(B19+B20)/2"
                .Range("A31").Value = "V_adj"
                .Range("B31").Formula = "=-B30"
                .Range("A32").Value = "(I_max+I_min)/2"
                .Range("B32").Formula = "=(B21+B22)/2"
                .Range("A33").Value = "I_adj"
                .Range("B33").Formula = "=-B32"
                .Range("A35").Value = "I_d_max"
                .Range("B35").Formula = "=MAX(N:N)"
                .Range("A36").Value = "I_d_min"
                .Range("B36").Formula = "=MIN(N:N)"
                .Range("A37").Value = "(I_d_max+I_d_min)/2"
                .Range("B37").Formula = "=(B35+B36)/2"
                .Range("A38").Value = "I_d_adj"
                .Range("B38").Formula = "=-B37"
                .Range("A40").Value = "R"
                .Range("B40").Value = 1000
                .Range("O1").Value = 0
                .Range("Q1").Formula = "=(1/(2*PI()*T1*R1))*B26"
                .Range("R1").Formula = "=SQRT(((B27^2)*(B26^2))/((B28^2)*(1+(B26^2))))"
                .Range("S1").Formula = "=B25"
                .Range("T1").Formula = "=1/S1"

                ' M列に式を自動入力
                For i = 1 To 10000
                    .Cells(i, "M").Formula = "=$B$21*SIN(2*PI()*J" & i & "/$S$1-$B$24)"
                Next i
                
                ' N列に式を自動入力
                For i = 1 To 10000
                    .Cells(i, "N").Formula = "=$B$21*SIN(2*PI()*D" & i & "/$S$1-$B$23+(PI()/2))"
                Next i
                
                ' P列に式を自動入力
                For i = 1 To 10000
                    .Cells(i, "P").Formula = "=L" & i & "-N" & i
                Next i

                ' F列に式を自動入力（先ほどの要件から）
                For i = 1 To 10000
                    .Cells(i, "F").Formula = "=$B$19*SIN(2*PI()*D" & i & "/$S$1-B$23)"
                Next i

                ' L列にK列の値を1000で割った結果を入力
                For i = 1 To 10000
                    .Cells(i, "L").Formula = "=K" & i & "/$B$36"
                Next i
            End With
            ' 指定列の書式を指数表示の8桁に設定
            Dim expCols As Variant
            expCols = Array("D", "E", "F", "J", "K", "L", "M", "N", "P", "Q", "R", "S", "T")
            
            Dim col As Variant
            For Each col In expCols
                ws.Columns(col).NumberFormat = "0.00000000E+00"
            Next col
            
            ' AとB列の書式を標準に設定
            ws.Columns("A:B").NumberFormat = "General"

            ' 新しいフォルダの作成
            Dim folderName As String
            folderName = csvPath & Replace(file.Name, ".csv", "")
            If Not fso.FolderExists(folderName) Then
                Set newFolder = fso.CreateFolder(folderName)
            End If

            ' .xlsxファイルを新しいフォルダに保存
            Dim savePath As String
            savePath = folderName & "\" & Replace(file.Name, ".csv", ".xlsx")
            ActiveWorkbook.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
            
            ' 再変換した.csvファイルを同じフォルダに保存
            Dim csvSavePath As String
            csvSavePath = folderName & "\" & Replace(file.Name, ".csv", "") & ".csv"
            ActiveWorkbook.SaveAs Filename:=csvSavePath, FileFormat:=xlCSV
            
            ActiveWorkbook.Close SaveChanges:=False
            filesProcessed = filesProcessed + 1
        End If
    Next file

    If filesProcessed > 0 Then
        MsgBox filesProcessed & " 個のファイルを処理しました。", vbInformation
    Else
        MsgBox "処理するCSVファイルが見つかりませんでした。", vbExclamation
    End If

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub
