<div align="center">

## Overlord


</div>

### Description

I have seen alot of code disabling the "X" on the form to prevent uses closing the form, i.e. in the middle of data processing, this is a little "work around" i have been using...no magic..(o: or heavy API calls
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Flemming Sudsgaard](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/flemming-sudsgaard.md)
**Level**          |Beginner
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/flemming-sudsgaard-overlord__1-37269/archive/master.zip)





### Source Code

```
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    MsgBox "You gotta use the button"
    Cancel = 1
  Else
    Unload Form1
    Set Form1 = Nothing
  End If
End Sub
```

