<div align="center">

## MSFlexGrid Column resizer


</div>

### Description

Hey guys out there, I make this function as i was seeking some good and easy way to make my FlexGrid look better.......I have not copied it, but if any other did the same then do tell me i will try to modify it....... and please feel free to use it and kindely also vote for me........Thanks
 
### More Info
 
name of the flex grid while calling function

User must be aware of function in VB..... and coding ways..

Not found yet, but if u find do tell me will try to debug it......


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mohit Agarwal 16](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mohit-agarwal-16.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mohit-agarwal-16-msflexgrid-column-resizer__1-66365/archive/master.zip)





### Source Code

```
Public Function AlignGridData(FlexGrid As MSFlexGrid)
'THIS FUNCTION IS USED FOR MAKING ENOUGH SPACE FOR TEXT SO THAT THE
'TEXT HAVING MAXIMUM WIDHT CAN BE EASILY BE READ IN FLEXGRID
  Dim ObjFrm As Form
  Dim MaxCol As String
  Dim J, I As Integer
  For J = 1 To FlexGrid.Cols - 1
    MaxCol = ""
    For I = 0 To FlexGrid.Rows - 2
     If I = 0 Then
       FlexGrid.Row = 0
       FlexGrid.Col = J
'       FlexGrid.CellFontBold = True
     End If
     If FlexGrid.Parent.TextWidth(FlexGrid.TextMatrix(I, J)) > FlexGrid.Parent.TextWidth(MaxCol) Then MaxCol = FlexGrid.TextMatrix(I, J)
    Next I
    FlexGrid.ColWidth(J) = FlexGrid.Parent.TextWidth(MaxCol) * 1.6
  Next J
 End Function
```

