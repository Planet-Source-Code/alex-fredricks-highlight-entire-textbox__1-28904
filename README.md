<div align="center">

## Highlight Entire Textbox


</div>

### Description

Highlight textbox/combobox and its entire contents. Simple but cosmetically useful.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Alex Fredricks](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/alex-fredricks.md)
**Level**          |Beginner
**User Rating**    |3.7 (22 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/alex-fredricks-highlight-entire-textbox__1-28904/archive/master.zip)





### Source Code

```
Public Sub Focus(varX As Variant)
'selects entire txtbox
 With varX
  If .Text <> "" Then
   .SelStart = 0
   .SelLength = Len(.Text)
  End If
 End With
End Sub
''''''''''''''''''''''''''''''''''''
call statement
''''''''''''''''''''''''''''''''''''
Private Sub txtStoreNo_GotFocus()
 Focus txtstoreno
End Sub
```

