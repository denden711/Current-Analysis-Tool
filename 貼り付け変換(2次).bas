Sub 貼り付け変換_2次()
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

    ' フォルダを取得
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(csvPath)

    Dim filesProcessed As Integer
    filesProcessed = 0

    ' フォルダ内の各ファイルを処理
    For Each file In folder.Files
        If Right(file.Name, 4) = ".csv" Then
            ProcessCSVFile csvPath, file, fso
            filesProcessed = filesProcessed + 1
        End If
    Next file

    ' 処理結果のメッセージを表示
    If filesProcessed > 0 Then
        MsgBox filesProcessed & " 個のファイルを処理しました。", vbInformation
    Else
        MsgBox "処理するCSVファイルが見つかりませんでした。", vbExclamation
    End If

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました (貼り付け変換_2次): " & Err.Description, vbCritical
End Sub

' CSVファイルの処理
Sub ProcessCSVFile(csvPath As String, file As Object, fso As Object)
    On Error GoTo ErrorHandler
    Dim ws As Worksheet

    ' CSVファイルを開く
    Workbooks.Open Filename:=csvPath & file.Name
    Set ws = ActiveWorkbook.Sheets(1)
    
    ' "L"列の右側に3列を挿入する
    InsertColumns ws, 3

    ' データラベルと数式を挿入
    InsertDataLabelsAndFormulas ws

    ' 各列に数式を範囲指定で挿入
    InsertFormulas ws

    ' 指定列の書式を指数表示の8桁に設定
    SetColumnNumberFormat ws

    ' 新しいフォルダの作成と保存
    SaveWorkbook ws, csvPath, file, fso

    ' ファイルを閉じる
    ActiveWorkbook.Close SaveChanges:=False

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました (ProcessCSVFile): " & Err.Description, vbCritical
    If Not ActiveWorkbook Is Nothing Then ActiveWorkbook.Close SaveChanges:=False
End Sub

' 指定された列の右側に複数列を挿入するサブルーチン
Sub InsertColumns(ws As Worksheet, numColumns As Integer)
    On Error GoTo ErrorHandler
    Dim j As Integer
    For j = 1 To numColumns
        ws.Columns("M:M").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Next j
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました (InsertColumns): " & Err.Description, vbCritical
End Sub

' データラベルと数式を挿入するサブルーチン
Sub InsertDataLabelsAndFormulas(ws As Worksheet)
    On Error GoTo ErrorHandler
    With ws
        ' A列にデータラベルを挿入
        .Range("A19").Value = "V_amp"
        .Range("A20").Value = "V_max"
        .Range("A21").Value = "V_min"
        .Range("A22").Value = "V_adj"
        .Range("A24").Value = "I_1_amp"
        .Range("A25").Value = "I_1_max"
        .Range("A26").Value = "I_1_min"
        .Range("A27").Value = "I_1_adj"
        .Range("A29").Value = "I_2_amp"
        .Range("A30").Value = "I_2_max"
        .Range("A31").Value = "I_2_min"
        .Range("A32").Value = "I_2_adj"
        .Range("A34").Value = "V_phase"
        .Range("A35").Value = "I_1_phase"
        .Range("A36").Value = "I_2_phase"
        .Range("A38").Value = "T [s]"
        .Range("A39").Value = "f [Hz]"
        .Range("A40").Value = "f [kHz]"
        .Range("A41").Value = "ω"
        .Range("A42").Value = "R_1 [Ω]"
        .Range("A43").Value = "R_2 [Ω]"
        .Range("A45").Value = "tanθ_1"
        .Range("A46").Value = "tanθ_2"
        .Range("A48").Value = "I_d_1_max"
        .Range("A49").Value = "I_d_2_max"

        ' B列にデータおよび数式を挿入
        .Range("B20").Formula = "=MAX(E:E)"
        .Range("B21").Formula = "=MIN(E:E)"
        .Range("B22").Formula = "=-(B20+B21)/2"
        .Range("B24").Formula = "=(B25-B26)/2"
        .Range("B27").Formula = "=-(B25+B26)/2"
        .Range("B29").Formula = "=(B30-B31)/2"
        .Range("B32").Formula = "=-(B30+B31)/2"
        .Range("B34").Value = 4.5
        .Range("B35").Value = 3.5
        .Range("B36").Formula = 3.5
        .Range("B39").Formula = "=1/B38"
        .Range("B40").Formula = "=B39/1000"
        .Range("B41").Formula = "=2*PI()*B39"
        .Range("B42").Value = 1000
        .Range("B43").Value = 1000
        .Range("B45").Formula = "=TAN(B34-B35)"
        .Range("B46").Formula = "=TAN(B34-B36)"
        .Range("B48").Formula = "=MAX(N:N)"
        .Range("B49").Formula = "=MAX(W:W)"
    End With
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました (InsertDataLabelsAndFormulas): " & Err.Description, vbCritical
End Sub

' 各列に数式を範囲指定で挿入するサブルーチン
Sub InsertFormulas(ws As Worksheet)
    On Error GoTo ErrorHandler
    Dim lastRow As Long
    lastRow = 10000 ' 必要に応じて調整

    With ws
        .Range("F1:F" & lastRow).Formula = "=$B$19*SIN(2*PI()*D1/$B$38-$B$34)"
        .Range("L1:L" & lastRow).Formula = "=K1/$B$42"
        .Range("M1:M" & lastRow).Formula = "=$B$24*SIN(2*PI()*D1/$B$38-$B$35)"
        .Range("N1:N" & lastRow).Formula = "=$B$24*SIN(2*PI()*D1/$B$38-$B$34+(PI()/2))"
        .Range("O1:O" & lastRow).Formula = "=L1-N1"
        .Range("U1:U" & lastRow).Formula = "=T1/$B$43"
        .Range("V1:V" & lastRow).Formula = "=$B$29*SIN(2*PI()*D1/$B$38-$B$36)"
        .Range("W1:W" & lastRow).Formula = "=$B$29*SIN(2*PI()*D1/$B$38-$B$34+(PI()/2))"
        .Range("X1:X" & lastRow).Formula = "=U1-W1"
        .Range("Z1:Z" & lastRow).Formula = "=D1"
        .Range("AA1:AA" & lastRow).Formula = "=$B$22+E1"
        .Range("AB1:AB" & lastRow).Formula = "=F1"
        .Range("AD1:AD" & lastRow).Formula = "=L1+$B$27"
        .Range("AE1:AE" & lastRow).Formula = "=M1"
        .Range("AF1:AF" & lastRow).Formula = "=N1"
        .Range("AG1:AG" & lastRow).Formula = "=O1+$B$27"
        .Range("AI1:AI" & lastRow).Formula = "=U1+$B$32"
        .Range("AJ1:AJ" & lastRow).Formula = "=V1"
        .Range("AK1:AK" & lastRow).Formula = "=W1"
        .Range("AL1:AL" & lastRow).Formula = "=X1+$B$32"
    End With
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました (InsertFormulas): " & Err.Description, vbCritical
End Sub

' 指定列の書式を指数表示の8桁に設定するサブルーチン
Sub SetColumnNumberFormat(ws As Worksheet)
    On Error GoTo ErrorHandler
    Dim expCols As Variant
    expCols = Array("D", "E", "F", "J", "K", "L", "M", "N", "O", "S", "T", "U", "V", "W", "X", "Z", "AA", "AB", "AD", "AE", "AF", "AG", "AI", "AJ", "AK", "AL")

    Dim col As Variant
    For Each col In expCols
        ws.Columns(col).NumberFormat = "0.00000000E+00"
    Next col
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました (SetColumnNumberFormat): " & Err.Description, vbCritical
End Sub

' ワークブックを保存するサブルーチン
Sub SaveWorkbook(ws As Worksheet, csvPath As String, file As Object, fso As Object)
    On Error GoTo ErrorHandler
    Dim folderName As String
    folderName = csvPath & Replace(file.Name, ".csv", "")
    
    ' フォルダが存在しない場合に作成
    If Not fso.FolderExists(folderName) Then
        fso.CreateFolder(folderName)
    End If

    ' .xlsxファイルを新しいフォルダに保存
    Dim savePath As String
    savePath = folderName & "\" & Replace(file.Name, ".csv", ".xlsx")
    ActiveWorkbook.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook

    ' 再変換した.csvファイルを同じフォルダに保存
    Dim csvSavePath As String
    csvSavePath = folderName & "\" & Replace(file.Name, ".csv", "") & ".csv"
    ActiveWorkbook.SaveAs Filename:=csvSavePath, FileFormat:=xlCSV
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました (SaveWorkbook): " & Err.Description, vbCritical
End Sub
