<div align="center">

## Dragging a form by a control


</div>

### Description

This code is reusable and small enough to paste into whatever you're doing and instantly have a form that has no need for a title bar.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VB Pro](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vb-pro.md)
**Level**          |Unknown
**User Rating**    |4.2 (161 globes from 38 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vb-pro-dragging-a-form-by-a-control__1-111/archive/master.zip)

### API Declarations

```
In the general declarations section, insert these lines:
Declare Sub ReleaseCapture Lib "User" ()
Declare Function SendMessage _
Lib "User" (ByVal hWnd As Integer, _
ByVal wMsg As Integer, _
ByVal wParem As Integer, lParem As Any) As Long
```


### Source Code

```

In the Mousedown event of the control, insert:
Sub Command1_MouseDown (Button As Integer, _
Shift As Integer, X As Single, Y As Single)
Dim Ret&
ReleaseCapture
Ret& = SendMessage(Me.hWnd, &H112, &HF012, 0)
End Sub
```

