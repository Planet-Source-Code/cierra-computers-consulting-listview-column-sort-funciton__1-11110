<div align="center">

## ListView Column Sort Funciton


</div>

### Description

Easily Sort Any ListView Column Ascending and Descending.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Cierra Computers & Consulting](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/cierra-computers-consulting.md)
**Level**          |Beginner
**User Rating**    |4.0 (36 globes from 9 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/cierra-computers-consulting-listview-column-sort-funciton__1-11110/archive/master.zip)





### Source Code

```
Public Sub SortListView(ctlListView As ListView, intColulunHeaderIndex As Integer)
ctlListView.Sorted = True
ctlListView.SortKey = intColulunHeaderIndex - 1
If ctlListView.SortOrder = lvwAscending Then
   ctlListView.SortOrder = lvwDescending
Else
   ctlListView.SortOrder = lvwAscending
End If
End Sub
```

