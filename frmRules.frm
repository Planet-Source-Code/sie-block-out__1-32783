VERSION 5.00
Begin VB.Form frmRules 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmRules.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRules 
      Height          =   2055
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   5655
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblHeader 
      Caption         =   "Rules Of Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
txtRules = "The aim of the game it to remove as many tiles from the grid as possible.  "
txtRules = txtRules + vbCrLf
txtRules = txtRules + "But the restriction is that you may only remove tiles that have adjacent tiles containing the identicle colour. "
txtRules = txtRules + vbCrLf + vbCrLf
txtRules = txtRules + "The more tiles that are removed from the grid with a single clickresults in more points scored.  Any "
txtRules = txtRules + "tiles remaining at the end of the game are deducted from your final total (yes is is possible to get a negative score!) "
txtRules = txtRules + vbCrLf + vbCrLf
txtRules = txtRules + "Although the aim is to remove all the tiles, it is still possible to get the highest score with tiles still remaining. "
txtRules = txtRules + vbCrLf + vbCrLf
txtRules = txtRules + "Note that the score is only stored in the Hall Of Fame when you cannot remove any more tiles from the grid. "
txtRules = txtRules + vbCrLf + vbCrLf
txtRules = txtRules + "Have fun with this and pass it on it IS freeware!"
txtRules = txtRules + vbCrLf
txtRules = txtRules + "Also, visit www.amalgyte.com and pass on your comments."
txtRules = txtRules + vbCrLf
txtRules = txtRules + vbTab + "Yours Sincerly,"
txtRules = txtRules + vbCrLf + vbCrLf
txtRules = txtRules + vbTab + "Simon Johnson"
txtRules = txtRules + vbCrLf
txtRules = txtRules + vbTab + "enquiries@amalgyte.com"






End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub
