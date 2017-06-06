### Accessに存在するすべてのテーブルをCSVファイルに出力する

```
Sub ExportCsvPerTable()
    Set mydb = CurrentDb
    
    For Each mytbl In mydb.TableDefs
        
        If Left(mytbl.name, 4) <> "MSys" Then 'システムテーブルは除外
            DoCmd.TransferText _
                TransferType:=acExportDelim, _
                TableName:=mytbl.name, _
                FileName:="C:\temp\" & mytbl.name & ".csv"
        End If
    Next
End Sub

```
