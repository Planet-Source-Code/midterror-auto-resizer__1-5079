<div align="center">

## Auto Resizer


</div>

### Description

I was sick and tired of seeing all the control resizers in planetsource, so i just made this myself and wanted to post it, cause someone might actually find it useful. All it does is when the form is resized, it changes all the controls (command buttons, lines, text boxes etc) to make the controls still look like they're in the right place. ie: If i had a command button that wa the size of the form, normally when i change the forms size, the command button is either too big for the form, or too little. With this, the command button is automatically resized so its still in the same proportion with the form.
 
### More Info
 
A form


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[MidTerror](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/midterror.md)
**Level**          |Beginner
**User Rating**    |4.9 (44 globes from 9 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/midterror-auto-resizer__1-5079/archive/master.zip)

### API Declarations

```
'Put this code in a module
Option Explicit
Dim PrevResizeX As Long
Dim PrevResizeY As Long
Public Function ResizeAll(FormName As Form)
  Dim tmpControl As Control
  On Error Resume Next
  'Ignores errors in case the control doesn't
  'have a width, height, etc.
  If PrevResizeX = 0 Then
    'If the previous form width was 0
    'Which means that this function wasn't run before
    'then change prevresizex and y and exit function
    PrevResizeX = FormName.ScaleWidth
    PrevResizeY = FormName.ScaleHeight
    Exit Function
  End If
  For Each tmpControl In FormName
    'A loop to make tmpControl equal to every
    'control on the form
    If TypeOf tmpControl Is Line Then
    'Checks the type of control, if its a
    'Line, change its X1, X2, Y1, Y2 values
      tmpControl.X1 = tmpControl.X1 / PrevResizeX * FormName.ScaleWidth
      tmpControl.X2 = tmpControl.X2 / PrevResizeX * FormName.ScaleWidth
      tmpControl.Y1 = tmpControl.Y1 / PrevResizeY * FormName.ScaleHeight
      tmpControl.Y2 = tmpControl.Y2 / PrevResizeY * FormName.ScaleHeight
      'These four lines see the previous ratio
      'Of the control to the form, and change they're
      'current ratios to the same thing
    Else
    'Changes everything elses left, top
    'Width, and height
      tmpControl.Left = tmpControl.Left / PrevResizeX * FormName.ScaleWidth
      tmpControl.Top = tmpControl.Top / PrevResizeY * FormName.ScaleHeight
      tmpControl.Width = tmpControl.Width / PrevResizeX * FormName.ScaleWidth
      tmpControl.Height = tmpControl.Height / PrevResizeY * FormName.ScaleHeight
      'These four lines see the previous ratio
      'Of the control to the form, and change they're
      'current ratios to the same thing
    End If
  Next tmpControl
  PrevResizeX = FormName.ScaleWidth
  PrevResizeY = FormName.ScaleHeight
  'Changes prevresize x and y to current width
  'and height
End Function
```


### Source Code

```
Option Explicit
Private Sub Form_Resize()
 ResizeAll Form1
'Calls for the ResizeAll function to run
'Change Form1 to the form name
End Sub
```

