<div align="center">

## Kill the Windowz Close Button \(X\)


</div>

### Description

This subroutine will disable the Windows 9x/NT

' close button, also known as the little X

' button. It's really useful when you have a

' form that you don't want to set its ControlBox

' property to False. Write this code in a

' standard module (BAS).
 
### More Info
 
Should be familiar with the Windows API.

If you're playing around with the system

' menu in other ways in your program, you

' might have to change the position number

' in your RemoveMenu function calls. Also,

' you could have problems running this

' with a MDI child.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jonathan Smith](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jonathan-smith.md)
**Level**          |Unknown
**User Rating**    |5.0 (35 globes from 7 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jonathan-smith-kill-the-windowz-close-button-x__1-3915/archive/master.zip)

### API Declarations

```
Private Declare Function GetSystemMenu Lib "user32" _
(ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOS = &H400&
```


### Source Code

```
Public Sub KillCloseButton(hWnd As Long)
 Dim hSysMenu As Long
 hSysMenu = GetSystemMenu(hWnd, 0)
 Call RemoveMenu(hSysMenu, 6, MF_BYPOS)
 Call RemoveMenu(hSysMenu, 5, MF_BYPOS)
End Sub
'Call the above function from a form as it's being loaded
Private Sub Form_Load()
 KillCloseButton Me.hWnd
End Sub
```

