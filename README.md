<div align="center">

## Duck Shoot Game


</div>

### Description

Create your own duck shoot style game, you can even kill your best buddy, just get a picture of him!!!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ed Hockaday](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ed-hockaday.md)
**Level**          |Unknown
**User Rating**    |6.0 (605 globes from 101 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Games](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/games__1-38.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ed-hockaday-duck-shoot-game__1-1568/archive/master.zip)





### Source Code

```
Leave Picture1 blank, make Picture2's picture a "kill picture" (so when the target is hit a bullet hole appears in it) and make Picture3's Picture the blank target (ie an "unwounded target")
Copy and paste this into a notepad and save as form1.frm
VERSION 5.00
Begin VB.Form Form1
  BorderStyle   =  4 'Fixed ToolWindow
  ClientHeight  =  3195
  ClientLeft   =  45
  ClientTop    =  285
  ClientWidth   =  4680
  LinkTopic    =  "Form1"
  MaxButton    =  0  'False
  MinButton    =  0  'False
  ScaleHeight   =  3195
  ScaleWidth   =  4680
  ShowInTaskbar  =  0  'False
  StartUpPosition =  3 'Windows Default
  Begin VB.Timer tmrSeconds
   Left      =  600
   Top       =  960
  End
  Begin VB.CommandButton cmdReset
   Caption     =  "&Reset"
   Height     =  375
   Left      =  120
   TabIndex    =  5
   Top       =  2760
   Visible     =  0  'False
   Width      =  855
  End
  Begin VB.PictureBox Picture3
   Appearance   =  0 'Flat
   BackColor    =  &H00C0C0C0&
   BorderStyle   =  0 'None
   ForeColor    =  &H80000008&
   Height     =  495
   Left      =  960
   ScaleHeight   =  495
   ScaleWidth   =  495
   TabIndex    =  4
   Top       =  2040
   Visible     =  0  'False
   Width      =  495
  End
  Begin VB.CommandButton cmdStart
   Caption     =  "&Start"
   Height     =  375
   Left      =  120
   TabIndex    =  3
   Top       =  2760
   Width      =  855
  End
  Begin VB.CommandButton cmdExit
   Caption     =  "&Exit"
   Height     =  375
   Left      =  3720
   TabIndex    =  2
   Top       =  2760
   Width      =  855
  End
  Begin VB.PictureBox Picture2
   Appearance   =  0 'Flat
   BackColor    =  &H80000004&
   BorderStyle   =  0 'None
   ForeColor    =  &H80000008&
   Height     =  495
   Left      =  480
   ScaleHeight   =  495
   ScaleWidth   =  495
   TabIndex    =  1
   Top       =  2040
   Visible     =  0  'False
   Width      =  495
  End
  Begin VB.PictureBox Picture1
   BackColor    =  &H008080FF&
   BorderStyle   =  0 'None
   Height     =  495
   Left      =  1920
   ScaleHeight   =  495
   ScaleWidth   =  495
   TabIndex    =  0
   Top       =  1320
   Width      =  495
  End
  Begin VB.Timer Timer1
   Interval    =  1
   Left      =  480
   Top       =  240
  End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DeltaX
Dim DeltaY
Dim gTimerSpeed
Dim gGameOn As Boolean
Dim gHit As Boolean
Dim gSeconds
Dim gShots
Dim gTime
Private Sub cmdExit_Click()
Unload Form1
End Sub
Private Sub cmdReset_Click()
Picture1.Picture = Picture3.Picture
Screen.MousePointer = vbCrosshair
Timer1.Interval = 0
DeltaX = 100  ' Initialize variables.
DeltaY = 100
cmdStart.Visible = True
cmdExit.Visible = True
cmdReset.Visible = False
End Sub
Private Sub cmdStart_Click()
Picture1.Picture = Picture3.Picture
Screen.MousePointer = vbCrosshair
Timer1.Interval = 1
gTimerSpeed = 1
DeltaX = 100  ' Initialize variables.
DeltaY = 100
cmdReset.Visible = False
cmdStart.Visible = False
cmdExit.Visible = False
gHit = False
gShots = 0
gGameOn = True
gTime = 0
tmrSeconds.Interval = 1000
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
gShots = gShots + 1
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If gGameOn = True Then
  Timer1.Interval = 0
  Picture1.Picture = Picture2.Picture
  Screen.MousePointer = Default
  cmdReset.Visible = True
  cmdExit.Visible = True
  gHit = True
  If gShots = 0 Then
    MsgBox "It took you " & gShots + 1 & " shot and " & gTime & " seconds to kill him!"
  ElseIf gShots > 0 Then
    MsgBox "It took you " & gShots + 1 & " shots and " & gTime & " seconds to kill him!"
  End If
  gGameOn = False
  tmrSeconds.Interval = 0
  Exit Sub
ElseIf gGameOn = False Then
  Exit Sub
End If
End Sub
Private Sub Timer1_Timer()
If gHit = True Then
  Timer1.Interval = 0
  Exit Sub
End If
If gTimerSpeed < 50 Then gTimerSpeed = gTimerSpeed + 1
Timer1.Interval = gTimerSpeed
  Picture1.Move Picture1.Left + DeltaX, Picture1.Top + DeltaY
  If Picture1.Left < ScaleLeft Then DeltaX = 100
  If Picture1.Left + Picture1.Width > ScaleWidth + ScaleLeft Then
    DeltaX = -100
  End If
  If Picture1.Top < ScaleTop Then DeltaY = 100
  If Picture1.Top + Picture1.Height > ScaleHeight + ScaleTop Then
    DeltaY = -100
  End If
End Sub
Private Sub tmrSeconds_Timer()
gTime = gTime + 1
End Sub
```

