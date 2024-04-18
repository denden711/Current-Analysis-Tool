Sub ConvertCSVsAndInsertDataFormulasWithFSO()
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim csvPath As String
    Dim xlsxPath As String
    Dim ws As Worksheet

    csvPath = "C:\Users\User\OneDrive - Chiba Institute of Technology\研究室\研究活動\202402\ワイヤー\y=4\解析\csv\"
    xlsxPath = "C:\Users\User\OneDrive - Chiba Institute of Technology\研究室\研究活動\202402\ワイヤー\y=4\解析\xlsx\"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(csvPath)

    For Each file In folder.Files
        If Right(file.Name, 4) = ".csv" Then
            ' CSVファイルを開く
            Workbooks.Open Filename:=csvPath & file.Name
            Set ws = ActiveWorkbook.Sheets(1)

        ' 指定されたセルにデータと式を挿入
        With ws
            ' 文字とMAX/MIN関数の挿入
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
            .Range("B26").Value = "=TAN(B23-B24)"
            .Range("A27").Value = "I/V"
            .Range("B27").Formula = "=B21/B19"
            .Range("A28").Value = "ω"
            .Range("B28").Formula = "=2*PI()*T1"
            .Range("A29").Value = "(|V_max|/|V_min|)/2"
            .Range("B29").Formula = "=(ABS(B19)+ABS(B20))/2"
            .Range("B30").Formula = "=B29-B19"
            .Range("A31").Value = "(|I_max|/|I_min|)/2"
            .Range("B31").Formula = "=(ABS(B21)+ABS(B22))/2"
            .Range("B32").Formula = "=B21-B31"
            .Range("A34").Value = "I_d_max"
            .Range("B34").Formula = "=MAX(N:N)"
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
                .Cells(i, "N").Formula = "=2*PI()*$T$1*$R$1*$B$19*SIN(2*PI()*D" & i & "/$S$1-$B$23+(PI()/2))"
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
                .Cells(i, "L").Formula = "=K" & i & "/1000"
            Next i
        End With


            ' Excel形式で保存
            ActiveWorkbook.SaveAs Filename:=xlsxPath & Replace(file.Name, ".csv", ".xlsx"), FileFormat:=xlOpenXMLWorkbook
            
            ' ファイルを閉じる
            ActiveWorkbook.Close SaveChanges:=False
        End If
    Next file
End Sub


