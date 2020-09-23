VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   4155
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8025
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MnuLoad 
      Caption         =   "&Loading App Forms"
      Begin VB.Menu MnuLoad1 
         Caption         =   "Non MdiChild Form"
      End
      Begin VB.Menu MnuLoad2 
         Caption         =   "Mdi Child 1"
      End
      Begin VB.Menu MnuLoad_etc 
         Caption         =   "Mdi Child 2 _etc"
      End
   End
   Begin VB.Menu MnuArrange 
      Caption         =   "Windows Arrange"
      Begin VB.Menu MnuCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu MnuTileHorizontal 
         Caption         =   "TileHorizontal"
      End
      Begin VB.Menu mnuTileVertical 
         Caption         =   "TileVertical"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''' VERY IMPORTANT:
''''' The ScrollBars Property = False required
''''' (that's all)

' 1. Copy fMdiBack.* to your App and add it;
'
' 2. imgBack.Picture = your background Picture
'    Suggested Bg Picture size is 1920x1440 (ideal jpg
'    compressed) but you may use any sized, any acceptable
'    file type.
'
' 3. The following are "pizzas":
'    ie., if you have a size-fixed logo image, -> igmLogo
'         ideal is a transparent gif to make background
'         accordance (no masks needed, no paints).
'
'         if you have a variable text or anything as, see
'         LbModulo with a logo relative positionning.
'
' 4. So, you can adjust easy code to do what you want.
'
' 5. Sorry there is a little but acceptable flicker when
'    positionning extras (logo, text, etc) ...
'    (if no extras, no flickers)
'
' WHAT ABOUT UPDATES:
'    Some aditional important informations:
'
'    1. Centralize_Back_Form Subroutine:
'       Added .ZOrder 1 to be sure fMdiBack don´t stay
'       over this "brothers" forms.
'
'    2. Added 2 fool child forms to load.
'
'    3. Load fAnyForm uses vbModeless to permits other
'       MdiLoads and commands.
'
'    4. Exemplifies about MdiForm Method .Arrange:
'       You have to call Centralize_Back_Form Subroutine
'       each new Arrange.
'
' jozew@globo.com
'
'- Resizes Aux Form Child as changes in MdiForm measures
'
Private Sub Centralize_Back_Form()
  With fMdiBack
   .Height = Me.ScaleHeight
   .Width = Me.ScaleWidth
   .Left = 0
   .Top = 0
   .ZOrder 1
  End With

End Sub

Private Sub MDIForm_Activate()
'
'
'Sure fMdiBack don't brings front others Forms
'when reassumes focus
   DoEvents
   Centralize_Back_Form
'
'
End Sub

Private Sub MDIForm_Load()
'
'
'Use Aux MdiChild form (dont .show it)
   Load fMdiBack
   'optional variable logo text
   fMdiBack.LbModulo.Caption = "MY_APPLICATION v" & App.Major & "." & App.Minor & "." & App.Revision
'
End Sub

Private Sub MDIForm_Resize()

' nothing for minimized
   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If
   
' optional sure form don't resizes beyond (opt) logo measures
   If Me.WindowState = vbNormal Then
  With fMdiBack
      If Me.ScaleHeight < .ImgLogo.Height Then
         Me.Height = .ImgLogo.Height + ((Me.Height - Me.ScaleHeight) / 2)
      End If
      If Me.ScaleWidth < .ImgLogo.Width Then
         Me.Width = .ImgLogo.Width + ((Me.Width - Me.ScaleWidth) / 2)
      End If
  End With
   End If
   
' so the essential command
   Centralize_Back_Form
End Sub

' some illustring procedures
Private Sub MnuFileExit_Click()
  End
End Sub

Private Sub MnuLoad1_Click()
  fAnyForm.Show vbModeless, Me
End Sub

Private Sub MnuLoad2_Click()
  fMdiChild1.Show
End Sub

Private Sub MnuLoad_etc_Click()
  fMdiChild2.Show
End Sub

'- Case App uses .Arrange for Childs
'
Private Sub MnuTileHorizontal_Click()
   Me.Arrange vbTileHorizontal
   Centralize_Back_Form
End Sub

Private Sub mnuTileVertical_Click()
   Me.Arrange vbTileVertical
   Centralize_Back_Form
End Sub

Private Sub MnuCascade_Click()
   Me.Arrange vbCascade
   Centralize_Back_Form
End Sub



