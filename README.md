<div align="center">

## Returning the Autonumber Value when adding records


</div>

### Description

When using Autonumber fields in a database for a UID, you might need this value after you add the record, for that record. This is my example of how to get that value back from the database after it is added using DAO or ADO.
 
### More Info
 
This is not a complete application, just code snippets. I expect that anyone using this knows that they have to set references to DAO and ADO, and know how to connect to a database using these objects.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[James N\. Wink](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/james-n-wink.md)
**Level**          |Beginner
**User Rating**    |4.3 (30 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/james-n-wink-returning-the-autonumber-value-when-adding-records__1-5966/archive/master.zip)





### Source Code

```
'DAO Example
'First Open a updateable recordset
Set rs = db.OpenRecordset("SomeTable")
  With rs
    'Start a New Record
    .AddNew
      !Field2 = "Add your data for this new record"
    'Add the record to the database
    .Update
    'Set the bookmark to Last modified
    .Bookmark = .LastModified
    lngResult = rs!AutoNumberUID
  End With
  rs.Close
'Ado Example
  Set mrsMDB = New ADODB.Recordset
  mrsMDB.CursorType = adOpenKeyset
  mrsMDB.LockType = adLockOptimistic
  mrsMDB.Open "SomeTable", mcnnMDB, , , adCmdTable
  With mrsMDB
    .AddNew
    !Field2 = "Add your Data for this record"
    .Update
    varBkMark = .Bookmark
    .Requery
    .Bookmark = varBkMark
    lngNewUID = !AutoNumberUID
  End With
```

