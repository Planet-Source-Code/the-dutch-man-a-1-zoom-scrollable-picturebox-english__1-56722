VERSION 5.00
Begin VB.Form frmPreview 
   AutoRedraw      =   -1  'True
   Caption         =   "Print Preview"
   ClientHeight    =   8385
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12690
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   12690
   Begin VB.PictureBox picControl 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   7815
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.ComboBox cboZoom 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   1
         Text            =   "cboZoom"
         Top             =   60
         Width           =   2175
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         TabIndex        =   2
         Top             =   60
         Width           =   855
      End
      Begin VB.Label lblView 
         Caption         =   "Zoom:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   10
         Top             =   90
         Width           =   495
      End
   End
   Begin VB.PictureBox picScroll 
      Height          =   6735
      Left            =   360
      ScaleHeight     =   6675
      ScaleWidth      =   9435
      TabIndex        =   3
      Top             =   600
      Width           =   9495
      Begin VB.VScrollBar vsPreview 
         Height          =   1215
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   480
         Width           =   255
      End
      Begin VB.HScrollBar hsPreview 
         Height          =   255
         Left            =   480
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   1725
      End
      Begin VB.PictureBox picShow 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   5415
         Left            =   360
         ScaleHeight     =   5415
         ScaleWidth      =   7020
         TabIndex        =   4
         Top             =   360
         Width           =   7020
         Begin VB.PictureBox picNormal 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   9495
            Left            =   -6960
            Picture         =   "frmPreview.frx":0ECA
            ScaleHeight     =   9495
            ScaleWidth      =   12615
            TabIndex        =   5
            Top             =   -5400
            Visible         =   0   'False
            Width           =   12615
         End
         Begin VB.PictureBox picHold 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1815
            Left            =   0
            ScaleHeight     =   1815
            ScaleWidth      =   2175
            TabIndex        =   6
            Top             =   0
            Width           =   2175
            Begin VB.PictureBox picDoc 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   1215
               Left            =   240
               ScaleHeight     =   1215
               ScaleWidth      =   1695
               TabIndex        =   7
               Top             =   240
               Visible         =   0   'False
               Width           =   1695
            End
         End
      End
      Begin VB.PictureBox Grabber 
         BackColor       =   &H00FFFF00&
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   11
         Top             =   120
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bScrollCode As Boolean
Private sZoom As Single
Private lPage As Integer
Private lPageMax As Integer
Private bDisplayPage As Boolean

Private Sub cboZoom_Click()

Dim iEvents As Integer

If Not bScrollCode Then
  If cboZoom.ListIndex >= 0 Then
    iEvents = DoEvents
    If cboZoom.ItemData(cboZoom.ListIndex) <> sZoom Then
      sZoom = cboZoom.ItemData(cboZoom.ListIndex)
      Zoom_Check
    End If
  End If
End If

End Sub

Private Sub cboZoom_KeyPress(KeyAscii As Integer)

Dim sNewZoom As Single

If KeyAscii = 13 Then
sNewZoom = Val(cboZoom.Text)
If sNewZoom > 0 And sNewZoom <= 200 Then
cboZoom.Text = sNewZoom & " %"
If sNewZoom = sZoom Then
Exit Sub
End If
sZoom = sNewZoom
Zoom_Check
Else
If cboZoom.ListIndex >= 0 Then
cboZoom.Text = cboZoom.List(cboZoom.ListIndex)
Else
cboZoom.Text = sZoom & " %"
End If
End If
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Me.Refresh
bDisplayPage = True
Preview_Display 1
End Sub

Private Sub Form_Load()
sZoom = 100
With cboZoom
  .AddItem "100 %"
  .ItemData(.ListCount - 1) = 100
  .AddItem "75 %"
  .ItemData(.ListCount - 1) = 75
  .AddItem "50 %"
  .ItemData(.ListCount - 1) = 50
  .AddItem "25 %"
  .ItemData(.ListCount - 1) = 25
  .AddItem "Full Page"
  .ItemData(.ListCount - 1) = 0
  .AddItem "Full Width"
  .ItemData(.ListCount - 1) = -1
  bScrollCode = True
  .ListIndex = 0
  bScrollCode = False
End With
sZoom = 100
picScroll.Move 0, picControl.Height, Me.ScaleWidth, Me.ScaleHeight - picControl.Height
vsPreview.Move picScroll.ScaleWidth - vsPreview.Width, 0, vsPreview.Width, picScroll.ScaleHeight - hsPreview.Height
hsPreview.Move 0, picScroll.ScaleHeight - hsPreview.Height, picScroll.ScaleWidth - vsPreview.Width
picShow.Move 0, 0, picScroll.ScaleWidth, picScroll.ScaleHeight
picDoc.Move -picDoc.Width, -picDoc.Height
End Sub

Public Sub Preview_Display(ByVal iPage As Integer)

Dim iMin As Integer
Dim iMax As Integer
Screen.MousePointer = vbHourglass
picNormal.Cls
Zoom_Check
Screen.MousePointer = vbDefault
End Sub
Private Sub Zoom_Check()

Dim sSizeX As Single
Dim sSizeY As Single
Dim sRatio As Single
Dim spImage As StdPicture
Dim sWidth As Single
Dim sHeight As Single
Dim bScroll As Byte
Dim bOldScroll As Byte
Screen.MousePointer = vbHourglass

sWidth = picScroll.ScaleWidth
sHeight = picScroll.ScaleHeight
Do
  bOldScroll = bScroll
  If sZoom = 0 Then
    sRatio = (sHeight - 480) / picNormal.Height
  ElseIf sZoom = -1 Then
    sRatio = (sWidth - 480) / picNormal.Width
  Else
    sRatio = sZoom / 100
  End If
  sSizeX = picNormal.Width * sRatio
  sSizeY = picNormal.Height * sRatio
  If sSizeX > sWidth And (bScroll And 1) <> 1 Then
    sHeight = sHeight - hsPreview.Height
    bScroll = bScroll + 1
  End If
  If sSizeY > sHeight And (bScroll And 2) <> 2 Then
    sWidth = sWidth - vsPreview.Width
    bScroll = bScroll + 2
  End If
Loop While bOldScroll <> bScroll

vsPreview.Height = sHeight
hsPreview.Width = sWidth

picShow.Move 0, 0, sWidth, sHeight
picDoc.Move 240, 240, sSizeX, sSizeY
picDoc.Cls
picDoc.PaintPicture picNormal.Image, 0, 0, sSizeX, sSizeY


' Laat scroll bars zien als dat nodig is
bScrollCode = True
picHold.Move 0, 0, sSizeX + 480, sSizeY + 480
If (bScroll And 2) = 2 Then
  vsPreview.Visible = True
  vsPreview.Max = (picHold.ScaleHeight - picShow.ScaleHeight) / 14.4 + 1
  vsPreview.Min = 0
  vsPreview.SmallChange = 14
  vsPreview.LargeChange = picShow.ScaleHeight / 14.4
  vsPreview.Value = vsPreview.Min
Else
  vsPreview.Visible = False
End If

If (bScroll And 1) = 1 Then
  hsPreview.Visible = True
  hsPreview.Max = (picHold.ScaleWidth - picShow.ScaleWidth) / 14.4 + 1
  hsPreview.Min = 0
  hsPreview.SmallChange = 14
  hsPreview.LargeChange = picShow.ScaleWidth / 14.4
  hsPreview.Value = hsPreview.Min
Else
  hsPreview.Visible = False
End If
bScrollCode = False
Screen.MousePointer = vbDefault
If bDisplayPage Then
picDoc.Visible = True
End If
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
Exit Sub
End If
If Me.ScaleHeight > 600 Then
picScroll.Move 0, 500, Me.ScaleWidth, Me.ScaleHeight - 500
End If
End Sub

Private Sub hsPreview_Change()

If Not bScrollCode Then
  picHold.Left = -hsPreview.Value * 14.4
End If
End Sub

Private Sub picScroll_Resize()
vsPreview.Left = picScroll.ScaleWidth - vsPreview.Width
vsPreview.Height = picScroll.ScaleHeight
hsPreview.Top = picScroll.ScaleHeight - hsPreview.Height
hsPreview.Width = picScroll.ScaleWidth
Zoom_Check
End Sub

Private Sub vsPreview_Change()
If Not bScrollCode Then
  picHold.Top = -vsPreview.Value * 14.4
End If
End Sub


