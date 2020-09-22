<div align="center">

## Create Access Table in VB


</div>

### Description

Allows the programmer to create an MS Access table in Visual Basic where the primary key’s field data type is set to AutoNumber. It is not like creating the primary key field in Access where you can select AutoNumber date type. In VB 5 and 6 you can’t request AutoNumber for a field type in the SQL Create Table string, it does not exist. So to create a table in VB where primary key’s numeric field type will be AutoNumber, you have to do it the way the included source code shows. Hope this helps.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Advanced
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/create-access-table-in-vb__1-8729/archive/master.zip)





### Source Code

```
Private Sub Creat_Table()
 Dim stSQLstr As String
 Dim dbs As Database
 stSQLstr = "CREATE TABLE NameTbl (NameID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, FirstName Text (15), LastName Text (20));"
  Set dbs = OpenDatabase("c:\test\Demo.mdb")
  dbs.Execute stSQLstr
  dbs.Close
End Sub
```

