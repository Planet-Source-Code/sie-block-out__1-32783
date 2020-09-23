VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmConfig 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Block Out"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   ControlBox      =   0   'False
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraConfig 
      Caption         =   "Configuration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   375
         Left            =   6840
         TabIndex        =   35
         Top             =   6480
         Width           =   855
      End
      Begin VB.TextBox txtBackground 
         Height          =   375
         Left            =   2160
         TabIndex        =   33
         Top             =   6480
         Width           =   4575
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   855
         Left            =   5400
         TabIndex        =   6
         Top             =   5160
         Width           =   1815
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Ok"
         Default         =   -1  'True
         Height          =   855
         Left            =   5400
         TabIndex        =   5
         Top             =   4200
         Width           =   1815
      End
      Begin VB.HScrollBar scrRows 
         Height          =   375
         Left            =   2640
         Max             =   50
         Min             =   5
         TabIndex        =   4
         Top             =   1080
         Value           =   5
         Width           =   4095
      End
      Begin VB.HScrollBar scrCols 
         Height          =   375
         Left            =   2640
         Max             =   50
         Min             =   5
         TabIndex        =   3
         Top             =   600
         Value           =   5
         Width           =   4095
      End
      Begin VB.Frame fraCellColours 
         Caption         =   "Cell Colours"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   360
         TabIndex        =   9
         Top             =   1560
         Width           =   7335
         Begin VB.OptionButton optStyle 
            Caption         =   "Circle"
            Height          =   375
            Index           =   0
            Left            =   4800
            TabIndex        =   31
            Top             =   1560
            Width           =   2175
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "Block"
            Height          =   375
            Index           =   1
            Left            =   4800
            TabIndex        =   30
            Top             =   1920
            Width           =   2055
         End
         Begin VB.PictureBox picBackColour 
            Height          =   255
            Left            =   2520
            ScaleHeight     =   195
            ScaleWidth      =   1515
            TabIndex        =   19
            Top             =   4080
            Width           =   1575
         End
         Begin VB.PictureBox picCellColour 
            Height          =   255
            Index           =   8
            Left            =   2520
            ScaleHeight     =   195
            ScaleWidth      =   1515
            TabIndex        =   18
            Top             =   3600
            Width           =   1575
         End
         Begin VB.PictureBox picCellColour 
            Height          =   255
            Index           =   7
            Left            =   2520
            ScaleHeight     =   195
            ScaleWidth      =   1515
            TabIndex        =   17
            Top             =   3240
            Width           =   1575
         End
         Begin VB.PictureBox picCellColour 
            Height          =   255
            Index           =   6
            Left            =   2520
            ScaleHeight     =   195
            ScaleWidth      =   1515
            TabIndex        =   16
            Top             =   2880
            Width           =   1575
         End
         Begin VB.PictureBox picCellColour 
            Height          =   255
            Index           =   5
            Left            =   2520
            ScaleHeight     =   195
            ScaleWidth      =   1515
            TabIndex        =   15
            Top             =   2520
            Width           =   1575
         End
         Begin VB.PictureBox picCellColour 
            Height          =   255
            Index           =   4
            Left            =   2520
            ScaleHeight     =   195
            ScaleWidth      =   1515
            TabIndex        =   14
            Top             =   2160
            Width           =   1575
         End
         Begin VB.PictureBox picCellColour 
            Height          =   255
            Index           =   3
            Left            =   2520
            ScaleHeight     =   195
            ScaleWidth      =   1515
            TabIndex        =   13
            Top             =   1800
            Width           =   1575
         End
         Begin VB.PictureBox picCellColour 
            Height          =   255
            Index           =   2
            Left            =   2520
            ScaleHeight     =   195
            ScaleWidth      =   1515
            TabIndex        =   12
            Top             =   1440
            Width           =   1575
         End
         Begin VB.PictureBox picCellColour 
            Height          =   255
            Index           =   1
            Left            =   2520
            ScaleHeight     =   195
            ScaleWidth      =   1515
            TabIndex        =   11
            Top             =   1080
            Width           =   1575
         End
         Begin VB.HScrollBar scrColours 
            Height          =   375
            Left            =   2520
            Max             =   8
            Min             =   2
            TabIndex        =   10
            Top             =   360
            Value           =   2
            Width           =   4095
         End
         Begin MSComDlg.CommonDialog dlgCommon 
            Left            =   6360
            Top             =   1320
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label1 
            Caption         =   "Grid Style"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   32
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label lblBackColour 
            Caption         =   "Background Colour"
            Height          =   255
            Index           =   9
            Left            =   360
            TabIndex        =   29
            Top             =   4080
            Width           =   1695
         End
         Begin VB.Label lblCellColours 
            Caption         =   "Cell Colour 8"
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   28
            Top             =   3600
            Width           =   1215
         End
         Begin VB.Label lblCellColours 
            Caption         =   "Cell Colour 7"
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   27
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label lblCellColours 
            Caption         =   "Cell Colour 6"
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   26
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label lblCellColours 
            Caption         =   "Cell Colour 5"
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   25
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label lblCellColours 
            Caption         =   "Cell Colour 4"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   24
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblCellColours 
            Caption         =   "Cell Colour 3"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   23
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label lblCellColours 
            Caption         =   "Cell Colour 2"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   22
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblCellColours 
            Caption         =   "Cell Colour 1"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   21
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lblNoColours 
            Caption         =   "Number Of Colours"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   20
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Label lblBackgroundFile 
         Caption         =   "Background Picture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   6480
         Width           =   1815
      End
      Begin VB.Label lblRows 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   8
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblCols 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblNoRows 
         Caption         =   "Number Of Rows"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblNoCols 
         Caption         =   "Number Of Columns"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
On Error GoTo errorHandler

dlgCommon.CancelError = True
dlgCommon.Filter = "JPeg Files (*.jpg)|*.jpg|Bitmap Files (*.bmp)|*.bmp|All Files (*.*)|*.*"
dlgCommon.ShowOpen
txtBackground.Text = dlgCommon.FileName

Exit Sub
errorHandler:
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim intloop As Integer
For intloop = 1 To 8
    BlockColours(intloop) = picCellColour(intloop).BackColor
Next intloop
colsX = scrCols.Value
colsY = scrRows.Value
intBackGroundColour = picBackColour.BackColor
strBackgroundFile = txtBackground.Text
frmBlockout.picTemp.Picture = LoadPicture(strBackgroundFile, , vbLPDefault)
intMaxColours = scrColours.Value
Unload Me
End Sub

Private Sub Form_Load()
Dim intloop As Integer
scrColours.Value = intMaxColours
scrCols.Value = colsX
scrRows.Value = colsY

For intloop = 1 To 8
    picCellColour(intloop).BackColor = BlockColours(intloop)
Next intloop
picBackColour.BackColor = intBackGroundColour
optStyle(intDrawStyle).Value = True
txtBackground.Text = strBackgroundFile
End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Private Sub optStyle_Click(Index As Integer)
    intDrawStyle = Index
End Sub

Private Sub picBackColour_Click()
dlgCommon.CancelError = True
On Error GoTo ErrorHandle
dlgCommon.ShowColor
picBackColour.BackColor = dlgCommon.Color

Exit Sub

ErrorHandle:

End Sub

Private Sub picCellColour_Click(Index As Integer)
dlgCommon.CancelError = True
On Error GoTo ErrorHandle

dlgCommon.ShowColor
picCellColour(Index).BackColor = dlgCommon.Color

Exit Sub

ErrorHandle:

End Sub

Private Sub scrColours_Change()
Dim intloop As Integer

For intloop = 1 To 8
    If intloop <= scrColours.Value Then
        picCellColour(intloop).Visible = True
        lblCellColours(intloop).Visible = True
    Else
        picCellColour(intloop).Visible = False
        lblCellColours(intloop).Visible = False
    End If
Next intloop

End Sub

Private Sub scrCols_Change()
lblCols.Caption = scrCols.Value
End Sub

Private Sub scrRows_Change()
lblRows.Caption = scrRows.Value
End Sub

Private Sub Text1_Change()

End Sub

