VERSION 5.00
Begin VB.Form fMdiBack 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7635
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "fMdiBack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin VB.Label LbModulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E.G., A VARIABLE TEXT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   1155
      TabIndex        =   0
      Top             =   2235
      Width           =   3435
   End
   Begin VB.Image ImgLogo 
      Height          =   2520
      Left            =   825
      Picture         =   "fMdiBack.frx":000C
      Stretch         =   -1  'True
      Top             =   60
      Width           =   3765
   End
   Begin VB.Image ImgBack 
      Height          =   7200
      Left            =   300
      Picture         =   "fMdiBack.frx":295E
      Stretch         =   -1  'True
      Top             =   -1065
      Width           =   6735
   End
End
Attribute VB_Name = "fMdiBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'this is to avoid or minimize SubResize reentrance flicker
Private aTop As Long
Private aLef As Long
Private aWid As Long
Private aHei As Long
'this is a variable text existence only requested;
'supposed Label is Logo containned:
Private aTopDiff As Long
Private aLeftDiff As Long

'- To avoid click effects
'
Private Sub Form_Click()
  Me.ZOrder 1
End Sub

Private Sub Form_DblClick()
  Me.ZOrder 1
End Sub


Private Sub Form_Load()
' getting initial measures and position
  aTop = Me.Top
  aLef = Me.Left
  aWid = Me.Width
  aHei = Me.Height

' getting relative lbModulo positionning
  aTopDiff = LbModulo.Top - ImgLogo.Top
  aLeftDiff = LbModulo.Left - ImgLogo.Left

End Sub

Private Sub Arrange_Me()
   If aTop = Me.Top And _
      aLef = Me.Left And _
      aWid = Me.Width And _
      aHei = Me.Height Then
         LbModulo.Visible = False
         ImgLogo.Visible = False
         With ImgBack
          .Top = 0
          .Left = 0
          .Width = Me.ScaleWidth
          .Height = Me.ScaleHeight
         End With
         DoEvents
         ImgLogo.Left = (Me.ScaleWidth - ImgLogo.Width) / 2
         ImgLogo.Top = (Me.ScaleHeight - ImgLogo.Height) / 2
         DoEvents
         LbModulo.Left = ImgLogo.Left + aLeftDiff
         LbModulo.Top = ImgLogo.Top + aTopDiff
         DoEvents
         ImgLogo.Visible = True
         LbModulo.Visible = True
   End If
End Sub

Private Sub Form_Resize()
   If Not aTop = Me.Top Then
      aTop = Me.Top
      Arrange_Me
   End If
   If Not aLef = Me.Left Then
      aLef = Me.Left
      Arrange_Me
   End If
   If Not aWid = Me.Width Then
      aWid = Me.Width
      Arrange_Me
   End If
   If Not aHei = Me.Height Then
      aHei = Me.Height
      Arrange_Me
   End If
End Sub

