<div align="center">

## Listbox Item •move• Multiselect compatable


</div>

### Description

Planning to use (a) listbox in your program? Maybe loading database data or MP3 Playlist. With these two functions (including the normal remove function) you are able to move items (compatable to Multiselect) up and down. I may be wrong but I haven't seen these two function on PSC so that's why I decided to posted these functions. Find it useful??? ***Please Vote***
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[uDmx IoCp©](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/udmx-iocp.md)
**Level**          |Intermediate
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/udmx-iocp-listbox-item-move-multiselect-compatable__1-34449/archive/master.zip)





### Source Code

```
'''ALL THESE FUNCTIONS ARE COMPATABLE TO MULTISELECT'''
'Moving Listbox item down
Public Function LstMoveDown(lst As ListBox)
Dim i
Dim strString As String
Dim strItemData As Long
For i = lst.ListCount - 2 To 0 Step -1
 If (lst.Selected(i) = False) Then GoTo skip
 strString = lst.List(i)
 strItemData = lst.ItemData(i)
 lst.RemoveItem (i)
 If i < lst.ListCount - 1 Then
  lst.AddItem strString, i + 1
  lst.ItemData(i + 1) = strItemData
  lst.Selected(i + 1) = True
 Else
  lst.AddItem strString
  lst.Selected(lst.ListCount - 1) = True
 End If
skip:
Next i
End Function
'Moving Listbox item up
Public Function LstMoveUp(lst As ListBox)
Dim i
Dim strString As String
Dim strItemData As Long
For i = 0 To lst.ListCount - 1
 If (lst.Selected(i) = False) Or i = 0 Then GoTo skip
 strString = lst.List(i)
 strItemData = lst.ItemData(i)
 lst.RemoveItem (i)
 lst.AddItem strString, i - 1
 lst.ItemData(i - 1) = strItemData
 lst.Selected(i - 1) = True
skip:
Next i
End Function
'Removing Listbox items
Public Function LstRemoveItem(lst As ListBox)
Dim i
For i = lst.ListCount - 1 To 0 Step -1
 If (lst.Selected(i) = True) Then
 lst.RemoveItem (i)
 End If
Next i
End Function
```

