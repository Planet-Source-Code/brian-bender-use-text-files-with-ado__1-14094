<div align="center">

## Use Text Files with ADO


</div>

### Description

Connect to text file(s) and perform advanced queries using ADO. You can even return recordsets on CSV file without a header.
 
### More Info
 


Save following Data in the app.path and name file Data.txt

ID,Name,Price

1,"Chairs",$40.00

2,"Table",$75.00

3,"Fork",$1.50

4,"Lamp",$15.00

5,"Rug",$35.00

6,"Desk",&150.00


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brian Bender](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian-bender.md)
**Level**          |Beginner
**User Rating**    |4.7 (33 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brian-bender-use-text-files-with-ado__1-14094/archive/master.zip)





### Source Code

```
Option Explicit
Dim oConn As New ADODB.Connection
Dim oRS As New ADODB.Recordset
Private Sub Form_Load()
    oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
      & "Data Source=" & App.Path & ";" _
      & "Extended Properties='text;FMT=Delimited'"
  '-- Use Following connection string if text file doesn't have a header for field names
  'oConn.Open "Provider=Microsoft.Jet" _
      & ".OLEDB.4.0;Data Source=" & App.Path _
      & ";Extended Properties='text;HDR=NO;" _
      & "FMT=Delimited'"
  Set oRS = oConn.Execute("Select * from Data.txt ")
  Dim ofield As ADODB.Field
  Do Until oRS.EOF
    For Each ofield In oRS.Fields
      Debug.Print "Field Name = " & ofield.Name & " Field Value = " & ofield.Value
    Next ofield
    oRS.MoveNext
  Loop
End Sub
```

