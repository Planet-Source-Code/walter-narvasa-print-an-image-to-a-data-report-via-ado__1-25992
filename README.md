<div align="center">

## PRINT AN IMAGE TO A DATA REPORT VIA ADO


</div>

### Description

Im sick and tired of using a third party report generator such as Crystal Reports, etc. so what i did was i mastered Data Report of Visual Basic 6 and here is sample of how to print an image to Data Report using ADO Connections. I hope you like this sample of mine.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Walter Narvasa](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/walter-narvasa.md)
**Level**          |Advanced
**User Rating**    |4.5 (36 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/walter-narvasa-print-an-image-to-a-data-report-via-ado__1-25992/archive/master.zip)





### Source Code

```
dim rs as ADODB.Recordset
dim db as ADODB.Connection
dim xDataPath as String
dim xPeoplePicture as String
xDataPath = App.Path & "\Database\Test.MDB"
db.CursorLocation = adUseClient
db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & xDataPath
Set rs = New Recordset
rs.Open "SELECT People_Name, People_Picture FROM People ORDER BY People_Name;", db, adOpenStatic, adLockOptimistic
rs.Movefirst
xPeoplePicture = rs!People_Picture ' "\SAMPLE.JPG" <-Sample Contents of rs recordset
Set DataReport1.DataSource = rs
Set DataReport1.Sections(3).Controls("Image1").Picture = LoadPicture( App. Path & xPeoplePicture)
DataReport1.Show
```

