VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Render 3d VB"
   ClientHeight    =   6675
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12645
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   12645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   9600
      TabIndex        =   12
      Top             =   4560
      Width           =   2775
      Begin VB.CheckBox Check2 
         Caption         =   "Reverse"
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "A&uto Rotate"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   5880
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Render while rotating"
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   4680
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.PictureBox picDest 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      Height          =   5685
      Left            =   4560
      ScaleHeight     =   256
      ScaleMode       =   0  'User
      ScaleWidth      =   256
      TabIndex        =   10
      Top             =   5760
      Visible         =   0   'False
      Width           =   5805
      Begin VB.Shape shSel 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   4107
         Shape           =   3  'Circle
         Top             =   4021
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdRotLeft 
      Caption         =   "Rotate &Left"
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdRotRight 
      Caption         =   "Rotate &Right"
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "2"
      Top             =   4725
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show A&dvanced controls"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   4080
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vertical scale"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   3975
      Begin VB.HScrollBar VScroll1 
         Height          =   255
         Left            =   600
         Max             =   50
         Min             =   1
         TabIndex        =   5
         Top             =   240
         Value           =   10
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "10"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "R&ender"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox Render1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      DrawWidth       =   3
      FillStyle       =   0  'Solid
      Height          =   3840
      Left            =   4200
      ScaleHeight     =   256
      ScaleMode       =   0  'User
      ScaleWidth      =   536
      TabIndex        =   1
      Top             =   120
      Width           =   8040
   End
   Begin VB.PictureBox Picmap 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      DrawWidth       =   3
      FillStyle       =   0  'Solid
      Height          =   3840
      Left            =   120
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   120
      Width           =   3840
   End
   Begin VB.Menu file1 
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu open1 
         Caption         =   "Open 2d landscape picture"
         Index           =   0
      End
      Begin VB.Menu Save1 
         Caption         =   "Save Rendered image as"
      End
      Begin VB.Menu exit1 
         Caption         =   "Exit"
         Index           =   0
      End
   End
   Begin VB.Menu RenderM 
      Caption         =   "Render"
      Index           =   1
      Begin VB.Menu RenderN 
         Caption         =   "Render Now"
         Index           =   1
      End
      Begin VB.Menu Rotleft 
         Caption         =   "Rotate Left"
      End
      Begin VB.Menu RotRt 
         Caption         =   "Rotate Right"
      End
      Begin VB.Menu Autorot 
         Caption         =   "Auto Rotate"
      End
      Begin VB.Menu Clear1 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu Tools1 
      Caption         =   "Tools"
      Begin VB.Menu Advan1 
         Caption         =   "Show advanced controls"
      End
      Begin VB.Menu Hide1 
         Caption         =   "Hide advanced controls"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Advan1_Click()
Form1.Height = 6000
End Sub


Private Sub Autorot_Click()
If Command3.Caption = "A&uto Rotate" Then
    Command3.Caption = "A&uto Rotate off"
Else
    Command3.Caption = "A&uto Rotate"
End If
End Sub

Private Sub Clear1_Click()
Render1.Cls
End Sub

Private Sub cmdRotLeft_Click()
    
    m_Angle = m_Angle - 5
    If m_Angle < -360 Then m_Angle = 0
    Text1.Text = m_Angle
    picDest.Cls
    RotateDC picDest.hdc, shSel.Left + (shSel.Width / 2), shSel.Top + (shSel.Height / 2), _
    Picmap.hdc, Picmap.Picture.Handle, m_Angle
    If Check1.Value = 1 Then Render
    
End Sub

Private Sub cmdRotRight_Click()

    m_Angle = m_Angle + 5
    If m_Angle > 360 Then m_Angle = 0
    Text1.Text = m_Angle
    picDest.Cls
    RotateDC picDest.hdc, shSel.Left + (shSel.Width / 2), shSel.Top + (shSel.Height / 2), _
    Picmap.hdc, Picmap.Picture.Handle, m_Angle
    If Check1.Value = 1 Then Render

End Sub

Private Sub Command1_Click()

Render

End Sub

Private Sub Command2_Click()

If Command2.Caption = "Show A&dvanced controls" Then
    Form1.Height = 6000
    Command2.Caption = "Hide A&dvanced controls"
Else
    Form1.Height = 5250
    Command2.Caption = "Show A&dvanced controls"
End If

End Sub


Private Sub Command3_Click()

If Command3.Caption = "A&uto Rotate" Then
    Command3.Caption = "A&uto Rotate off"
Else
    Command3.Caption = "A&uto Rotate"
End If

End Sub


Private Sub exit1_Click(Index As Integer)

End

End Sub

Private Sub Form_Load()

Form1.Height = 5250
m_Angle = 45
picDest.Cls
Text1.Text = m_Angle
RotateDC picDest.hdc, shSel.Left + (shSel.Width / 2), shSel.Top + (shSel.Height / 2), _
Picmap.hdc, Picmap.Picture.Handle, m_Angle
Render

End Sub
Sub Render()

' This is the render routine, pretty simple, the original bitmap is copied to
' Picdest (which you'll find dangling at the bottom of the form) which is rotated
' thanks to Raul Fragoso rotating image routine Submitted on: 3/12/2002
' The routine simply reads this image, stuffs colour info into a two dimensional array
' then plots it onto a destination picturebox.
' Play about with it to your hearts content, but be aware that I've standardised scaleheights
' & scalewidths to base 256

Render1.Cls
Command1.Caption = "Rendering"
Command1.Enabled = False
ModYscale = 160 / picDest.ScaleHeight
MyXscale = (Render1.ScaleWidth - 2)
ModXscale = (Render1.ScaleWidth - 2) / Picmap.ScaleWidth
For GenC = 1 To 256 Step 1
   For GenC2 = 1 To 256 Step 1
        PlotX = GenC
        PlotY = GenC2
        ColInfo(GenC, GenC2) = picDest.Point(PlotX, PlotY)
   Next GenC2
Next GenC
Offset = ((Render1.ScaleHeight / 2))

For GenC = 1 To 255 Step 1
   For GenC2 = 1 To 255 Step 1
      PlotX = GenC
      PlotY = GenC2
      PlotY = PlotY + 120
      PlotX = PlotX * ModXscale
      PlotY = PlotY * ModYscale
      
      ' The routine ignores any pixels of background colour, otherwise ruins rendered image
      If ColInfo(GenC, GenC2) <> picDest.BackColor Then
           Render1.Line (PlotX, PlotY)-(PlotX, PlotY - ((ColInfo(GenC, GenC2) And 255)) _
           * (VScroll1 / 100)), ColInfo(GenC, GenC2)
      End If
   Next GenC2
Next GenC
Command1.Enabled = True
Command1.Caption = "R&ender"

End Sub


Private Sub Hide1_Click()
Form1.Height = 5250
End Sub

Private Sub open1_Click(Index As Integer)
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
  
    CommonDialog1.ShowOpen
    Currpath = CommonDialog1.filename
    Set Picmap.Picture = LoadPicture(Currpath)
    picDest.Cls
    RotateDC picDest.hdc, shSel.Left + (shSel.Width / 2), shSel.Top + (shSel.Height / 2), _
    Picmap.hdc, Picmap.Picture.Handle, m_Angle
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub


Private Sub Picmap_DblClick()
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
  
    CommonDialog1.ShowOpen
    Currpath = CommonDialog1.filename
    Set Picmap.Picture = LoadPicture(Currpath)
    
    picDest.Cls
    RotateDC picDest.hdc, shSel.Left + (shSel.Width / 2), shSel.Top + (shSel.Height / 2), _
    Picmap.hdc, Picmap.Picture.Handle, m_Angle
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub RenderN_Click(Index As Integer)
Render
End Sub

Sub Rotate()
    picDest.Cls
    RotateDC picDest.hdc, shSel.Left + (shSel.Width / 2), shSel.Top + (shSel.Height / 2), _
    Picmap.hdc, Picmap.Picture.Handle, m_Angle
End Sub

Private Sub Rotleft_Click()
    m_Angle = m_Angle - 5
    If m_Angle < -360 Then m_Angle = 0
    Text1.Text = m_Angle
    picDest.Cls
    RotateDC picDest.hdc, shSel.Left + (shSel.Width / 2), shSel.Top + (shSel.Height / 2), _
    Picmap.hdc, Picmap.Picture.Handle, m_Angle
    If Check1.Value = 1 Then Render
End Sub

Private Sub RotRt_Click()
    m_Angle = m_Angle + 5
    If m_Angle > 360 Then m_Angle = 0
    Text1.Text = m_Angle
    picDest.Cls
    RotateDC picDest.hdc, shSel.Left + (shSel.Width / 2), shSel.Top + (shSel.Height / 2), _
    Picmap.hdc, Picmap.Picture.Handle, m_Angle
    If Check1.Value = 1 Then Render
End Sub

Private Sub Save1_Click()
' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.Filter = "*.bmp"
    CommonDialog1.ShowSave
    Currpath = CommonDialog1.filename & ".bmp"
        SavePicture Render1.Image, Currpath
        MsgBox "Rendered image Saved", vbOKOnly, "Saved"
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub



Private Sub Timer1_Timer()
If Command3.Caption = "A&uto Rotate off" Then
    If Check2 = 1 Then
        m_Angle = m_Angle - 5
        If m_Angle < -360 Then m_Angle = 0
    Else
        m_Angle = m_Angle + 5
        If m_Angle > 360 Then m_Angle = 0
    End If
    Text1.Text = m_Angle
    Rotate
    Render
End If
End Sub

Private Sub VScroll1_Change()
Label1.Caption = VScroll1.Value
End Sub
