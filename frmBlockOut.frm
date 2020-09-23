VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmBlockout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   Caption         =   "BLOCK OUT"
   ClientHeight    =   5610
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6795
   ClipControls    =   0   'False
   Icon            =   "frmBlockOut.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      DrawStyle       =   5  'Transparent
      FillColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   0
      ScaleHeight     =   4695
      ScaleWidth      =   3735
      TabIndex        =   17
      Top             =   0
      Width           =   3735
      Begin MSFlexGridLib.MSFlexGrid flxTable 
         Height          =   3015
         Left            =   960
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   5318
         _Version        =   393216
         Cols            =   4
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5235
      Left            =   5700
      ScaleHeight     =   5235
      ScaleWidth      =   1095
      TabIndex        =   1
      Top             =   0
      Width           =   1095
      Begin VB.Label lblTiles 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Tiles Left:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblScore 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Score:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Position:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblPosn 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label lblColor 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblColor 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblColor 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblColor 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblColor 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblColor 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblColor 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblColor 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar stbScore 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5235
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   2963
            TextSave        =   "10/01/2002"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "09:16"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuGameHeader 
      Caption         =   "Game"
      Begin VB.Menu mnuGame 
         Caption         =   "&New"
         Index           =   1
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuGame 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuGame 
         Caption         =   "&Configuration"
         Index           =   3
         Shortcut        =   +^{F4}
      End
      Begin VB.Menu mnuGame 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuGame 
         Caption         =   "E&xit"
         Index           =   5
         Shortcut        =   %{BKSP}
      End
   End
   Begin VB.Menu mnuViewHeader 
      Caption         =   "View"
      Begin VB.Menu mnuView 
         Caption         =   "Top Scores"
         Index           =   1
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuHelpHeader 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Rules"
         Index           =   1
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "About"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmBlockout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim varwait As Variant

Dim intGrid() As Integer
Dim intCheck() As Boolean
Dim intscore As Double
Dim intCellsLeft As Integer
Dim intPairs As Integer
Dim boolGameOver As Boolean

'Game Menu
Const bytNewIndex = 1
Const bytConfigurationIndex = 3
Const bytExitIndex = 5
'View Menu
Const bytTopScoresIndex = 1
'Help Menu
Const bytRulesIndex = 1
Const bytAboutIndex = 3
Private Sub Form_Load()
intDrawStyle = 1
strBackgroundFile = ""
Dim tex
colsX = 10
colsY = 10
defineColours
intBackGroundColour = frmBlockout.BackColor
getDefaultSettings
If strBackgroundFile <> "" And Dir(strBackgroundFile) <> "" Then
    frmBlockout.picTemp.Picture = LoadPicture(strBackgroundFile, , vbLPDefault)

    subGetSizePicture
End If
            
frmBlockout.Refresh
strname = "Unknown"
subnewgame
subMakeColourChanges
End Sub
Private Sub subnewgame()
Dim intloopa As Integer

flxTable.Visible = False
mnuView(bytTopScoresIndex).Checked = False
ReDim intGrid(0 To colsX - 1, 0 To colsY - 1) As Integer
ReDim intCheck(0 To colsX - 1, 0 To colsY - 1) As Boolean
newGrid
intscore = 0
intCellsLeft = colsX * colsY
picGrid.Enabled = True
picGrid.AutoRedraw = True

End Sub
Private Sub defineColours()
Dim intloop As Integer
intMaxColours = 4
BlockColours(1) = vbRed
BlockColours(2) = vbBlue
BlockColours(3) = vbGreen
BlockColours(4) = vbCyan
BlockColours(5) = vbYellow
BlockColours(6) = vbMagenta
BlockColours(7) = vbWhite
BlockColours(8) = vbBlack


End Sub

Private Sub newGrid()
Randomize
Dim intloopx As Integer
Dim intloopy As Integer
Dim intcellwidth As Integer
Dim intcellheight As Integer
Dim dblColour As Double
boolGameOver = False
intcellwidth = picGrid.Width / colsX - 1
intcellheight = picGrid.Height / colsY - 1
For intloopx = 0 To colsX - 1
    For intloopy = 0 To colsY - 1
        dblColour = Int(Rnd * intMaxColours) + 1
        intGrid(intloopx, intloopy) = dblColour
        If intGrid(intloopx, intloopy) <> 0 Then plotblock intloopx, intloopy, intcellwidth, intcellheight
    Next intloopy
Next intloopx
subGetHighScores
subGetScore

End Sub

Private Sub plotblock(PosX As Integer, posY As Integer, WidthX As Integer, WidthY As Integer)
Dim dblAspect As Double
Select Case intDrawStyle
    Case bytCircleIndex
        picGrid.FillColor = BlockColours(intGrid(PosX, posY))
        picGrid.FillStyle = 0
        picGrid.Circle ((PosX * WidthX) + (WidthX / 2), (posY * WidthY) + (WidthY / 2)), (WidthX / 2), 1
    Case bytSquareIndex
        picGrid.Line (PosX * WidthX, posY * WidthY)-Step(WidthX, WidthY), BlockColours(intGrid(PosX, posY)), BF
        
End Select
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_Resize
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_Resize
End Sub

Private Sub Form_Paint()
frmBlockout.BackColor = intBackGroundColour
'fraBoard.BackColor = intBackGroundColour
picGrid.BackColor = intBackGroundColour
Form_Resize

End Sub

Private Sub Form_Resize()
Dim intcellwidth As Integer
Dim intcellheight As Integer
Dim intloopx As Integer
Dim intloopy As Integer
Dim intsavewidth As Integer
Dim intsaveheight As Integer
Dim intminwidth As Integer
Dim intminheight As Integer
'frmBlockout.AutoRedraw = False

intminwidth = 640
intminheight = 2000
intsavewidth = frmBlockout.Width
intsaveheight = frmBlockout.Height

If frmBlockout.Width < intminwidth Then frmBlockout.Width = intminwidth
If frmBlockout.Height < intminheight Then frmBlockout.Height = intminheight

'fraBoard.Width = frmBlockout.ScaleWidth - Picture1.ScaleWidth   '(frmBlockout.Width / 100) * 90
'fraBoard.Height = frmBlockout.ScaleHeight '- stbScore.Height '((frmBlockout.Height - stbScore.Height) / 100) * 80
'fraBoard.Top = 0 '(frmBlockout.Height / 2) - (fraBoard.Height / 2)
'fraBoard.Left = 0 '(frmBlockout.Width / 2) - (fraBoard.Width / 2)
picGrid.Width = frmBlockout.ScaleWidth - Picture1.ScaleWidth   '(frmBlockout.Width / 100) * 90
picGrid.Height = frmBlockout.ScaleHeight - stbScore.Height '((frmBlockout.Height - stbScore.Height) / 100) * 80

picGrid.Left = 0
picGrid.Top = 0

subFlexRedrawTable

picGrid.Cls
subGetSizePicture
intcellwidth = picGrid.Width / colsX - 1
intcellheight = picGrid.Height / colsY - 1
For intloopx = 0 To colsX - 1
    For intloopy = 0 To colsY - 1
        If intGrid(intloopx, intloopy) <> 0 Then plotblock intloopx, intloopy, intcellwidth, intcellheight
    Next intloopy
Next intloopx
'frmBlockout.AutoRedraw = True
End Sub

Private Sub Form_Terminate()
Unload Me
End
End Sub


Private Sub mnuGame_Click(Index As Integer)
Dim intloop As Integer
Select Case Index
    Case bytNewIndex
        subnewgame
    Case bytConfigurationIndex
        frmConfig.Show vbModal
        If Not (frmConfig.cmdCancel) Then
            Set frmConfig = Nothing
            saveDefaultSettings
            subnewgame
            If strBackgroundFile <> "" And Dir(strBackgroundFile) <> "" Then
                subGetSizePicture
            End If
            Form_Paint
        End If
    Case bytExitIndex
        Unload Me
        End
End Select
End Sub

Private Sub mnuHelp_Click(Index As Integer)
Select Case Index
    Case bytRulesIndex
        frmRules.Show vbModal
        Set frmRules = Nothing
    Case bytAboutIndex
        frmAbout.Show vbModal
        Set frmAbout = Nothing
End Select

End Sub

Private Sub mnuView_Click(Index As Integer)
Select Case Index
    Case bytTopScoresIndex
        If mnuView(bytTopScoresIndex).Checked Then
            'view game card
            flxTable.Visible = False
            If boolGameOver Then picGrid.Enabled = False
            
            mnuView(bytTopScoresIndex).Checked = False
            Form_Paint
        Else
            'view top score table
            subGetHighScores
            flxTable.Visible = True
            picGrid.Enabled = True
            
            mnuView(bytTopScoresIndex).Checked = True
        End If
End Select
End Sub

Private Sub picGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim intcellwidth As Integer
Dim intcellheight As Integer
Dim intloopx As Integer
Dim intloopy As Integer
Dim boolmoreavail As Boolean
frmBlockout.AutoRedraw = False
intcellwidth = picGrid.Width / colsX - 1
intcellheight = picGrid.Height / colsY - 1

MousePointer = vbHourglass
If intGrid(Int(X / intcellwidth), Int(Y / intcellheight)) <> 0 Then
    
    If funcheckneighbours(Int(X / intcellwidth), Int(Y / intcellheight)) Then
        subGetScore
        subMakeColourChanges
        Form_Paint
    End If
End If

resetCheckGrid
lblScore.Caption = intscore
lblTiles.Caption = intCellsLeft
If funCheckForAnyNeighbours = 0 Then
    
    boolGameOver = True
    picGrid.DrawWidth = 20
    picGrid.AutoRedraw = False
    picGrid.Line (0, 0)-(picGrid.Width, picGrid.Height), vbRed
    picGrid.DrawWidth = 1
    picGrid.AutoRedraw = True
    intscore = intscore - (intCellsLeft ^ 2)
    lblPosn.Caption = funGetPosn
    lblScore.Caption = intscore
    lblTiles.Caption = intCellsLeft
    subSaveHighScores
    subGetHighScores
    flxTable.Visible = True
    mnuView(bytTopScoresIndex).Checked = True
End If

MousePointer = vbDefault
frmBlockout.AutoRedraw = True
End Sub

Private Function funcheckneighbours(X As Integer, Y As Integer) As Boolean


resetCheckGrid
intCheck(X, Y) = True
funcheckneighbours = False
Dim boolchangemade As Boolean
Dim intloopx As Integer
Dim intloopy As Integer
Dim intcount As Integer
Dim intNumCols As Integer

intNumCols = colsX - 1
Do While intGrid(intNumCols, colsY - 1) = 0 And intNumCols > 0
    intNumCols = intNumCols - 1
Loop
intcount = 1
'On Error Resume Next
boolchangemade = True
Do While boolchangemade = True
    boolchangemade = False
    For intloopx = 0 To intNumCols
        For intloopy = 0 To colsY - 1
            If intGrid(intloopx, intloopy) = intGrid(X, Y) And intCheck(intloopx, intloopy) = False And (X <> intloopx Or Y <> intloopy) Then
                If intloopx - 1 >= 0 Then
                    If intCheck(intloopx - 1, intloopy) = True Then
                            intCheck(intloopx, intloopy) = True
                            boolchangemade = True
                            intcount = intcount + 1
                    End If
                End If
                If intloopy - 1 >= 0 Then
                    If intCheck(intloopx, intloopy - 1) = True Then
                            intCheck(intloopx, intloopy) = True
                            boolchangemade = True
                            intcount = intcount + 1
                    End If
                End If
                If intloopx + 1 <= intNumCols Then
                    If intCheck(intloopx + 1, intloopy) = True Then
                            intCheck(intloopx, intloopy) = True
                            boolchangemade = True
                            intcount = intcount + 1
                    End If
                End If
                If intloopy + 1 <= colsY - 1 Then
                    If intCheck(intloopx, intloopy + 1) = True Then
                            intCheck(intloopx, intloopy) = True
                            boolchangemade = True
                            intcount = intcount + 1
                    End If
                End If
            End If

        Next intloopy
    Next intloopx

Loop
If intcount > 1 Then
    funcheckneighbours = True
End If

End Function
Private Sub resetCheckGrid()

Dim intloopx As Integer
Dim intloopy As Integer
intCellsLeft = 0
For intloopx = 0 To colsX - 1
    For intloopy = 0 To colsY - 1
        intCheck(intloopx, intloopy) = False
        If intGrid(intloopx, intloopy) <> 0 Then
            intCellsLeft = intCellsLeft + 1
        End If

    Next intloopy
Next intloopx

End Sub

Private Sub subMakeColourChanges()

Dim intloopx As Integer
Dim intloopy As Integer
Dim intloopz As Integer
Dim intloopa As Integer
Dim intcount As Integer
Dim boolchangemade As Boolean

'Update Legend
For intloopx = 0 To 7
   lblColor(intloopx).BackColor = BlockColours(intloopx + 1)
   lblColor(intloopx).ForeColor = BlockColours(7 - intloopx + 1)
Next intloopx

For intloopx = 0 To colsX - 1
    For intloopy = 0 To colsY - 1
        If intCheck(intloopx, intloopy) Then
            intGrid(intloopx, intloopy) = 0
        End If
    Next intloopy
Next intloopx

'now close up gaps

For intloopx = 0 To colsX - 1
    boolchangemade = True
    Do While boolchangemade = True
        boolchangemade = False
        intloopz = 0
        While intGrid(intloopx, intloopz) = 0 And intloopz < colsY - 2
          intloopz = intloopz + 1
        Wend
        intcount = 0
        For intloopy = intloopz To colsY - 1
            If intGrid(intloopx, intloopy) = 0 Then
                intcount = intcount + 1
            End If
        Next intloopy
        If intcount > 0 Then
            For intloopz = 1 To intcount
                For intloopy = 0 To colsY - 2
                    If intGrid(intloopx, intloopy) <> 0 Then
                        If intGrid(intloopx, intloopy + 1) = 0 Then
                            intGrid(intloopx, intloopy + 1) = intGrid(intloopx, intloopy)
                            intGrid(intloopx, intloopy) = 0
                            boolchangemade = True
                        End If
                    End If
                Next intloopy
            Next intloopz
        End If
    Loop
Next intloopx

Dim intNumCols As Integer
intNumCols = colsX - 1
Do While intGrid(intNumCols, colsY - 1) = 0 And intNumCols > 0
    intNumCols = intNumCols - 1
Loop

'now check bottom of columns - if 0 move left
For intloopa = 0 To intNumCols - 1
For intloopx = 0 To intNumCols - 1
    intloopy = colsY - 1
    If intGrid(intloopx, intloopy) = 0 Then
        For intloopz = 0 To intloopy
            intGrid(intloopx, intloopz) = intGrid(intloopx + 1, intloopz)
            intGrid(intloopx + 1, intloopz) = 0
        Next intloopz
    End If
Next intloopx
Next intloopa


End Sub

Private Sub subGetScore()

Dim intloopx As Integer
Dim intloopy As Integer
Dim intCellCount As Integer


For intloopx = 0 To 7
    lblColor(intloopx).Caption = ""
Next intloopx

intCellCount = 0
intCellsLeft = 0
For intloopx = 0 To colsX - 1
    For intloopy = 0 To colsY - 1
        If intCheck(intloopx, intloopy) Then
            intCellCount = intCellCount + 1
        Else
        Select Case intGrid(intloopx, intloopy)
            Case 1
                lblColor(0).Caption = Val(lblColor(0).Caption) + 1
            Case 2
                lblColor(1).Caption = Val(lblColor(1).Caption) + 1
            Case 3
                lblColor(2).Caption = Val(lblColor(2).Caption) + 1
            Case 4
                lblColor(3).Caption = Val(lblColor(3).Caption) + 1
            Case 5
                lblColor(4).Caption = Val(lblColor(4).Caption) + 1
            Case 6
                lblColor(5).Caption = Val(lblColor(5).Caption) + 1
            Case 7
                lblColor(6).Caption = Val(lblColor(6).Caption) + 1
            Case 8
                lblColor(7).Caption = Val(lblColor(7).Caption) + 1
        End Select
        End If
    Next intloopy
Next intloopx
'now multiply score diff by power of 2
intscore = intscore + ((intCellCount) ^ 2)
'show score
lblPosn.Caption = funGetPosn

End Sub

Private Function funCheckForAnyNeighbours() As Integer
Dim intloopx As Integer
Dim intloopy As Integer
Dim intcount As Integer

intPairs = 0
intcount = 0
For intloopx = 0 To colsX - 1
    For intloopy = 0 To colsY - 1
        If intGrid(intloopx, intloopy) <> 0 Then
            If intloopx - 1 >= 0 Then
                If intGrid(intloopx - 1, intloopy) = intGrid(intloopx, intloopy) Then
                        intPairs = intPairs + 1
                End If
            End If
            If intloopy - 1 >= 0 Then
                If intGrid(intloopx, intloopy - 1) = intGrid(intloopx, intloopy) Then
                        intPairs = intPairs + 1
                End If
            End If
            If intloopx + 1 <= colsX - 1 Then
                If intGrid(intloopx + 1, intloopy) = intGrid(intloopx, intloopy) Then
                        intPairs = intPairs + 1
                End If
            End If
            If intloopy + 1 <= colsY - 1 Then
                If intGrid(intloopx, intloopy + 1) = intGrid(intloopx, intloopy) Then
                        intPairs = intPairs + 1
                End If
            End If
        End If
    Next intloopy
Next intloopx
funCheckForAnyNeighbours = intPairs
End Function

Private Sub subSaveHighScores()
Dim intloop As Integer
Dim intError As Boolean

intError = True

Do While intError = True
    strname = InputBox("Enter Name", "High Score", strname)
    If strname = "" Or strname = "Unknown" Then
        intError = True
    Else
        intError = False
    End If
Loop

For intloop = 1 To Len(strname)
    If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ 1234567890", UCase(Mid$(strname, intloop, 1))) = 0 Then
        Mid$(strname, intloop, 1) = "_"
    End If
Next intloop
Open (App.Path & "\HighScores.blk") For Append As #1
Write #1, intMaxColours, colsX, colsY, Crypt(strname), intscore, Crypt(Format(Now, "DDD DD/MM/YYYY HH:NN:SS")), intCellsLeft
Close #1
End Sub

Private Sub subGetHighScores()
Dim user As String
Dim intColours, xcols, ycols As Integer
Dim Score As Double
Dim intcount As Integer
Dim strDateTime As String
intcount = 1
On Error Resume Next

flxTable.Clear
flxTable.Cols = 5
flxTable.TextMatrix(0, 0) = "POS"
flxTable.TextMatrix(0, 1) = "DATE"
flxTable.TextMatrix(0, 2) = "NAME"
flxTable.TextMatrix(0, 3) = "SCORE"
flxTable.TextMatrix(0, 4) = "CELLS LEFT"

subFlexRedrawTable

Open (App.Path & "\HighScores.blk") For Input As #1
Do While Not EOF(1)
    Input #1, intColours, xcols, ycols, user, Score, strDateTime, intCellsLeft
    If intColours = intMaxColours And xcols = colsX And ycols = colsY Then
        flxTable.Rows = intcount + 1
        flxTable.TextMatrix(intcount, 1) = Crypt(strDateTime)
        flxTable.TextMatrix(intcount, 2) = Crypt(user)
        flxTable.TextMatrix(intcount, 3) = Score
        flxTable.TextMatrix(intcount, 4) = intCellsLeft
        intcount = intcount + 1
    End If
Loop
Close #1
flxTable.Col = flxTable.Cols - 1
flxTable.Sort = flexSortGenericAscending
flxTable.Col = flxTable.Cols - 2
flxTable.Sort = flexSortGenericDescending

For intcount = 1 To flxTable.Rows - 1
    flxTable.TextMatrix(intcount, 0) = intcount
Next intcount
End Sub

Private Sub subFlexRedrawTable()
flxTable.Left = 0
flxTable.Top = 0
flxTable.Width = picGrid.Width
flxTable.Height = picGrid.Height
flxTable.ColWidth(0) = (flxTable.Width / 100) * 10
flxTable.ColWidth(1) = (flxTable.Width / 100) * 30
flxTable.ColWidth(2) = (flxTable.Width / 100) * 35
flxTable.ColWidth(3) = (flxTable.Width / 100) * 15
flxTable.ColWidth(4) = (flxTable.Width / 100) * 20
flxTable.ColAlignment(3) = vbAlignLeft
End Sub

Private Sub saveDefaultSettings()
Dim intloop As Integer
Open App.Path + "blockout.ini" For Output As 1
Write #1, colsX, colsY, intMaxColours, intBackGroundColour, strBackgroundFile, intDrawStyle
For intloop = 1 To 8
    Write #1, BlockColours(intloop)
Next intloop


Close #1
End Sub
Private Sub getDefaultSettings()
Dim intloop As Integer
On Error GoTo ErrorHandle
Open App.Path + "blockout.ini" For Input As 1
Input #1, colsX, colsY, intMaxColours, intBackGroundColour, strBackgroundFile, intDrawStyle
For intloop = 1 To 8
    Input #1, BlockColours(intloop)
Next intloop
Close #1
Exit Sub
ErrorHandle:
End Sub


Private Function funGetPosn() As Integer
Dim intloop As Integer
intloop = 1
While intscore < Val(flxTable.TextMatrix(intloop, 3)) And intloop < flxTable.Rows - 1
    intloop = intloop + 1
Wend
If intloop = flxTable.Rows - 1 Then
    intloop = intloop + 1
End If
funGetPosn = intloop
End Function

Private Sub subGetSizePicture()
Dim tex As Long
'    picGrid.Cls
    tex = StretchBlt(picGrid.hdc, 0, 0, picGrid.Width, picGrid.Height, picTemp.hdc, 0, 0, picTemp.Width, picTemp.Height, vbSrcCopy)
    picGrid.Refresh
    
End Sub
