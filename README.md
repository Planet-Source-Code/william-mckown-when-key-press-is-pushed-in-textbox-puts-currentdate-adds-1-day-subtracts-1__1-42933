<div align="center">

## When Key Press =,\+, \- is pushed in textbox puts currentdate, adds 1 day, subtracts 1 day


</div>

### Description

This code is a input feature that shows how to take a keypress on a textbox and either add, subtract or set current date. + key will add 1 day, - will subtract ond day, and = sets current date. This helps with getting dates from users.
 
### More Info
 
Just open a new project, add a text box and paste code. Enjoy


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[William McKown](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/william-mckown.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/william-mckown-when-key-press-is-pushed-in-textbox-puts-currentdate-adds-1-day-subtracts-1__1-42933/archive/master.zip)





### Source Code

```
Private Sub Form_Load()
 Text1 = Date
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
  Case Is = 61
   KeyAscii = 0
   Text1 = Date
  Case Is = 43
   KeyAscii = 0
    If Text1 = "" Then
    Text1 = DateAdd("d", 1, Date)
   Else
    Text1 = DateAdd("d", 1, Text1)
    End If
  Case Is = 45
   KeyAscii = 0
   If Text1 = "" Then
    Text1 = DateAdd("d", -1, Date)
   Else
    Text1 = DateAdd("d", -1, Text1)
   End If
 End Select
End Sub
```

