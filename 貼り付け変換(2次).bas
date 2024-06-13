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
            
            ' "L"列の右側に6列を挿入する
            Dim j As Integer
            For j = 1 To 6
                Columns("M:M").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Next j

            ' 指定されたセルにデータと式を挿入
            With ws
                .Range("A20").Value = "I_1_max"
                .Range("A21").Value = "I_1_min"
                .Range("A22").Value = "I_2_max"
                .Range("A23").Value = "I_2_min"
                .Range("A25").Value = "V_1_phase"
                .Range("A26").Value = "I_1_phase"
                .Range("A27").Value = "V_2_phase"
                .Range("A28").Value = "I_2_phase"
                .Range("A29").Value = "T"
                .Range("A30").Value = "f"
                .Range("A31").Value = "R_1"
                .Range("A32").Value = "R_2"
                .Range("A34").Value = "V_max"
                .Range("A35").Value = "V_min"
                .Range("A36").Value = "tanθ_1"
                .Range("A37").Value = "tanθ_2"
                .Range("A39").Value = "I_1/V"
                .Range("A40").Value = "I_2/V"
                .Range("A42").Value = "ω"
                .Range("A44").Value = "(V_max+V_min)/2"
                .Range("A45").Value = "V_adj"
                .Range("A46").Value = "(I_1_max+I_1_min)/2"
                .Range("A47").Value = "I_1_adj"
                .Range("A48").Value = "(I_2_max+I_2_min)/2"
                .Range("A49").Value = "I_1_adj"
                .Range("A51").Value = "Id_1_max"
                .Range("A52").Value = "Id_1_min"
                .Range("A53").Value = "(I_d_1_max+I_d_1_min)/2"
                .Range("A54").Value = "I_d_1_adj"
                .Range("A55").Value = "Id_2_max"
                .Range("A56").Value = "Id_2_min"
                .Range("A57").Value = "(I_d_2_max+I_d_2_min)/2"
                .Range("A58").Value = "I_d_2_adj"

                .Range("B25").Value = 4.5
                .Range("B26").Value = 3.5
                .Range("B27").Formula = "=B25"
                .Range("B28").Value = 3.5
                .Range("B30").Formula = "=1/B29"
                .Range("B31").Value = 1000
                .Range("B32").Value = 1000
                .Range("B34").Formula = "=MAX(E:E)"
                .Range("B35").Formula = "=MIN(E:E)"
                .Range("B36").Formula = "=TAN(B25-B26)"
                .Range("B37").Formula = "=TAN(B27-B28)"
                .Range("B39").Formula = "=B20/B34"
                .Range("B40").Formula = "=B22/B34"
                .Range("B42").Formula = "=2*PI()*B30"
                .Range("B44").Formula = "=(B34+B35)/2"
                .Range("B45").Formula = "=-B44"
                .Range("B46").Formula = "=(B20+B21)/2"
                .Range("B47").Formula = "=-B46"
                .Range("B48").Formula = "=(B22+B23)/2"
                .Range("B49").Formula = "=-B48"
                .Range("B51").Formula = "=MAX(N:N)"
                .Range("B52").Formula = "=MIN(N:N)"
                .Range("B53").Formula = "=(B51+B52)/2"
                .Range("B54").Formula = "=-B53"
                .Range("B55").Formula = "=MAX(Z:Z)"
                .Range("B56").Formula = "=MIN(Z:Z)"
                .Range("B57").Formula = "=(B52+B53)/2"
                .Range("B58").Formula = "=-B57"

                .Range("O1").Value = 0
                .Range("Q1").Formula = "=(1/(2*PI()*$B$30*$R$1))*$B$36"
                .Range("R1").Formula = "=SQRT((($B$39^2)*($B$36^2))/(($B$42^2)*(1+($B$36^2))))"

                .Range("AA1").Value = 0
                .Range("AC1").Formula = "=(1/(2*PI()*$B$30*$AD$1))*$B$37"
                .Range("AD1").Formula = "=SQRT((($B$40^2)*($B$37^2))/(($B$42^2)*(1+($B$37^2))))"

                ' F列に式を自動入力
                For i = 1 To 10000
                    .Cells(i, "F").Formula = "=$B$34*SIN(2*PI()*D" & i & "/$B$29-$B$25)"
                Next i

                ' L列に式を自動入力
                For i = 1 To 10000
                    .Cells(i, "L").Formula = "=K" & i & "/$B$31"
                Next i

                ' M列に式を自動入力
                For i = 1 To 10000
                    .Cells(i, "M").Formula = "=$B$20*SIN(2*PI()*J" & i & "/$B$29-$B$26)"
                Next i

                ' N列に式を自動入力
                For i = 1 To 10000
                    .Cells(i, "N").Formula = "=$B$20*SIN(2*PI()*J" & i & "/$B$29-$B$25+(PI()/2))"
                Next i

                ' P列に式を自動入力
                For i = 1 To 10000
                    .Cells(i, "P").Formula = "=L" & i & "-N" & i
                Next i

                ' X列に式を自動入力
                For i = 1 To 10000
                    .Cells(i, "X").Formula = "=W" & i & "/$B$32"
                Next i

                ' Y列に式を自動入力
                For i = 1 To 10000
                    .Cells(i, "Y").Formula = "=$B$22*SIN(2*PI()*V" & i & "/$B$29-$B$28)"
                Next i
                
                ' Z列に式を自動入力
                For i = 1 To 10000
                    .Cells(i, "Z").Formula = "=$B$22*SIN(2*PI()*V" & i & "/$B$29-$B$27+(PI()/2))"
                Next i

                ' AB列に式を自動入力
                For i = 1 To 10000
                    .Cells(i, "AB").Formula = "=X" & i & "-Z" & i
                Next i

            End With
            ' 指定列の書式を指数表示の8桁に設定
            Dim expCols As Variant
            expCols = Array("D", "E", "F", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD")
            
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
