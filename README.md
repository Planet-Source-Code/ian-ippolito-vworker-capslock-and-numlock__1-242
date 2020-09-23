<div align="center">

## CapsLock and NumLock


</div>

### Description

How to Activate CapsLock and NumLock from Code
 
### More Info
 
The keyboard APIs for VB4-16 and VB3 do not support the byte data type.

By changing the Windows constant to Public Const VK_NUMLOCK = &H90, you can use the above to activate the NumLock key.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ian Ippolito \(vWorker\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ian-ippolito-vworker.md)
**Level**          |Unknown
**User Rating**    |3.5 (14 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\)
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ian-ippolito-vworker-capslock-and-numlock__1-242/archive/master.zip)

### API Declarations

```
Public Const VK_CAPITAL = &H14
Public Type KeyboardBytes
     kbByte(0 To 255) As Byte
End Type
Public kbArray As KeyboardBytes
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long
Public Declare Function GetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long
Public Declare Function SetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long
```


### Source Code

```
On a form, add a 3 command buttons (cmdToggle, cmdTurnOff, cmdTurnOff) and a label. Add the following code to the form:
Private Function CapsLock() As Integer
	CapsLock = GetKeyState(VK_CAPITAL) And 1 = 1
End Function
Private Sub Form_Load()
	If CapsLock() = 1 Then Label1 = "On" Else Label1 = "Off"
End Sub
Private Sub cmdToggle_Click()
	GetKeyboardState kbArray
	kbArray.kbByte(VK_CAPITAL) = IIf(kbArray.kbByte(VK_CAPITAL) = 1, 0, 1)
	SetKeyboardState kbArray
	Label1 = IIf(CapsLock() = 1, "On", "Off")
End Sub
Private Sub cmdTurnOn_Click()
	GetKeyboardState kbArray
	kbArray.kbByte(VK_CAPITAL) = 1
	SetKeyboardState kbArray
	Label1 = IIf(CapsLock() = 1, "On", "Off")
End Sub
Private Sub cmdTurnOff_Click()
	GetKeyboardState kbArray
	kbArray.kbByte(VK_CAPITAL) = 0
	SetKeyboardState kbArray
	Label1 = IIf(CapsLock() = 1, "On", "Off")
End Sub
```

