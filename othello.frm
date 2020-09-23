VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOthello 
   BackColor       =   &H00000000&
   Caption         =   "Othello"
   ClientHeight    =   9390
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12330
   FillColor       =   &H00C0C0C0&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   HasDC           =   0   'False
   Icon            =   "othello.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   9390
   ScaleWidth      =   12330
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Change color"
      Height          =   315
      Left            =   8880
      TabIndex        =   15
      Top             =   3600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   120
      Top             =   2520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   8040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtIp 
      Height          =   375
      Left            =   8760
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdIp 
      Caption         =   "Get Your IP"
      Height          =   375
      Left            =   8760
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox chkClient 
      Caption         =   "Connect"
      Height          =   375
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox chkServer 
      Caption         =   "Host Game"
      Height          =   375
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock WinChat 
      Left            =   240
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WinClear 
      Left            =   360
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   480
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picBlack 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1800
      Picture         =   "othello.frx":020A
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   4440
      Width           =   735
   End
   Begin VB.PictureBox picWhite 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1800
      Picture         =   "othello.frx":069B
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   3
      Top             =   2880
      Width           =   735
   End
   Begin VB.Timer tmAiMove 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   600
      Top             =   2520
   End
   Begin VB.CommandButton cmdMoves 
      BackColor       =   &H00000000&
      DisabledPicture =   "othello.frx":0B4F
      Height          =   2295
      Left            =   120
      MaskColor       =   &H00004000&
      Picture         =   "othello.frx":932A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin MSFlexGridLib.MSFlexGrid FlexGrid 
      Height          =   5415
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   9551
      _Version        =   393216
      Rows            =   8
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
      GridColor       =   0
      AllowBigSelection=   0   'False
      HighLight       =   0
      GridLines       =   2
      BorderStyle     =   0
      MousePointer    =   99
      MouseIcon       =   "othello.frx":106EF
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   7440
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "The button below changes the color of the label box that shows what the other person is saying"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   8760
      TabIndex        =   16
      Top             =   2760
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   5880
      Visible         =   0   'False
      Width           =   2055
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -40
      WindowlessVideo =   0   'False
   End
   Begin VB.Label lbltalk 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1335
      Left            =   2160
      TabIndex        =   13
      Top             =   6000
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label lblip 
      BackStyle       =   0  'Transparent
      Caption         =   "If your not the host enter the host's ip in the below text box"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   8760
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblblack 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   540
      Left            =   2040
      TabIndex        =   6
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label lblwhite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   540
      Left            =   2040
      TabIndex        =   5
      Top             =   2400
      Width           =   255
   End
   Begin VB.Image hand2 
      Height          =   555
      Left            =   480
      Picture         =   "othello.frx":10909
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image hand 
      Height          =   555
      Left            =   480
      Picture         =   "othello.frx":10DE8
      Stretch         =   -1  'True
      Top             =   4560
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&NewGame"
         Begin VB.Menu mnuOnePLayer 
            Caption         =   "&One Player"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuTwoPlayer 
            Caption         =   "&Two Player"
         End
         Begin VB.Menu mnuNetGame 
            Caption         =   "N&et Game"
         End
      End
      Begin VB.Menu mnuClear 
         Caption         =   "C&lear"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "A&bout"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmOthello"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'http://www.armory.com/~iioa/othguide/faq/othellorules.html
Dim Turn As Boolean
Const MaxCols = 8
Const MaxRows = 8
Const Gridrows = 9
Const Gridcols = 9
Const VbDarkGreen = &H8000&
Dim GridArray(-1 To Gridrows, -1 To Gridcols) As Integer
Dim RegRow As Integer, RegCol As Integer
Dim NewRow As Integer, NewCol As Integer
Dim CheckRow As Integer, CheckCol As Integer
Dim HoldRow As Integer, HoldCol As Integer
Dim StopGame As Integer
Dim OptionClick As Boolean, HoldMove As Boolean
Dim Pos1 As Boolean
Dim HoldRow1 As Integer, HoldCol1 As Integer
Dim HoldPieces As Integer, Pieces As Integer
Dim NetGame As Boolean
Dim ClearGame As Integer, ClearGame2 As Integer
Dim PlayerName
Dim NoGameMoves As Integer, Messages As Integer
Private Sub chkClient_Click()
Select Case chkClient.Value
    Case "1"
        If txtIp = "" Then
            chkClient.Caption = "Disconnect"
            MsgBox ("Need to enter a ip please.")
            chkClient.Value = 0
            Exit Sub
        End If
        PlayerName = InputBox("Enter your name or nickname.", "Othello", "Dirtball2")
        If PlayerName = "" Then
            chkClient.Value = 0
            Exit Sub
        End If
        Winsock.RemotePort = 2121
        Winsock.RemoteHost = txtIp
        WinClear.RemotePort = 2120
        WinClear.RemoteHost = txtIp
        WinChat.RemotePort = 2100
        WinChat.RemoteHost = txtIp
        Call WinClear.Connect
        Call Winsock.Connect
        Call WinChat.Connect
        NetGame = True
        chkServer.Enabled = False
        chkClient.Caption = "Disconnect"
    Case "0"
        chkClient.Caption = "Connect"
        chkServer.Enabled = True
        Winsock.Close
        WinClear.Close
        WinChat.Close
        Winsock.LocalPort = 0
        WinClear.LocalPort = 0
        WinChat.LocalPort = 0
        Command1.Enabled = False
        Text1.Enabled = False
        FlexGrid.Enabled = False
        NetGame = False
End Select
End Sub
Private Sub chkServer_Click()
Select Case chkServer.Value
    Case "1"
        PlayerName = InputBox("Enter your name or nickname.", "Othello", "Dirtball")
        If PlayerName = "" Then
            chkServer.Value = 0
            Exit Sub
        End If
        Winsock.LocalPort = 2121
        Winsock.Listen
        WinClear.LocalPort = 2120
        WinClear.Listen
        WinChat.LocalPort = 2100
        WinChat.Listen
        NetGame = True
        chkClient.Enabled = False
        chkServer.Caption = "Disconnect"
    Case "0"
        chkClient.Enabled = True
        chkServer.Caption = "Host Game"
        Winsock.Close
        WinClear.Close
        WinChat.Close
        Winsock.LocalPort = 0
        WinClear.LocalPort = 0
        WinChat.LocalPort = 0
        Command1.Enabled = False
        Text1.Enabled = False
        FlexGrid.Enabled = False
        NetGame = False
End Select
End Sub

Private Sub cmdIp_Click()
'MsgBox ("your ip is " & Winsock.LocalIP)
txtIp.Text = Winsock.LocalIP
End Sub

Private Sub Command1_Click()
Call WinChat.SendData(PlayerName & ":" & "         " & Trim$(Text1) & vbCrLf)
Text1.Text = ""
End Sub

Private Sub Command2_Click()
CommonDialog1.ShowColor
lbltalk.ForeColor = CommonDialog1.Color
End Sub

Private Sub Form_Terminate()
Winsock.Close
WinClear.Close
End Sub

Private Sub mnuNetGame_Click()
Dim message As String
message = MsgBox("Are you sure you want to exit this game and clear it", vbYesNo + vbQuestion)
If message = vbYes Then
    Call mnuClear_Click
    mnuTwoPlayer.Checked = False
    mnuOnePLayer.Checked = False
    mnuNetGame.Checked = True
    chkServer.Visible = True
    chkClient.Visible = True
    txtIp.Visible = True
    cmdIp.Visible = True
    lblip.Visible = True
    Text1.Visible = True
    Command1.Visible = True
    FlexGrid.Enabled = False
    Command2.Visible = True
    Label1.Visible = True
Else
    Exit Sub
End If
End Sub

Private Sub mnuOnePLayer_Click()
Dim message As String
message = MsgBox("Are you sure you want to exit this game and clear it", vbYesNo + vbQuestion)
If message = vbYes Then
    Call mnuClear_Click
    mnuTwoPlayer.Checked = False
    mnuOnePLayer.Checked = True
    mnuNetGame.Checked = False
    chkServer.Visible = False
    chkClient.Visible = False
    txtIp.Visible = False
    cmdIp.Visible = False
    lblip.Visible = False
    Text1.Visible = False
    Command1.Visible = False
    Command2.Visible = False
    Label1.Visible = False
    Winsock.Close
    WinClear.Close
Else
    Exit Sub
End If
End Sub
Private Sub mnuTwoPlayer_Click()
Dim message As String
message = MsgBox("Are you sure you want to exit this game and clear it", vbYesNo + vbQuestion)
If message = vbYes Then
    Call mnuClear_Click
    mnuTwoPlayer.Checked = True
    mnuOnePLayer.Checked = False
    mnuNetGame.Checked = False
    chkServer.Visible = False
    chkClient.Visible = False
    txtIp.Visible = False
    cmdIp.Visible = False
    lblip.Visible = False
    Text1.Visible = False
    Command1.Visible = False
    Command2.Visible = False
    Label1.Visible = False
    Winsock.Close
    WinClear.Close
Else
    Exit Sub
End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command1_Click
    Text1.Text = ""
End If
End Sub

Private Sub Timer1_Timer()
lbltalk.Visible = False
Timer1.Enabled = False
lbltalk.Caption = ""
Messages = 0
End Sub

Private Sub tmaimove_Timer()
Call AiMove2
End Sub
Private Sub cmdMoves_Click()
Dim Direction As Integer
Dim ValidRow As Integer, ValidCol As Integer
Dim NoMoves As Boolean
Dim message As String
Pos1 = False
For CheckRow = 0 To MaxRows - 1
    For CheckCol = 0 To MaxCols - 1
        If GridArray(CheckRow, CheckCol) = 0 Then
            For Direction = 1 To 8
                NoMoves = CheckVal(CheckRow, CheckCol, Direction, False)
                If Pos1 = True Then
                    Exit For
                End If
            Next Direction
        End If
        If Pos1 = True Then
            Exit For
        End If
    Next CheckCol
    If Pos1 = True Then
        Exit For
    End If
Next CheckRow
If Pos1 = True Then
    message = MsgBox("there is a move", vbInformation + vbOKOnly)
Else
    message = MsgBox("there are no moves", vbInformation + vbOKOnly)
    If Turn = True Then
        hand.Visible = False
        hand2.Visible = True
        If Pos1 = False Then
            StopGame = StopGame + 1
            Turn = False
            hand.Visible = True
            hand2.Visible = False
        End If
    Else
        hand.Visible = True
        hand2.Visible = False
        If Pos1 = False Then
            StopGame = StopGame + 1
            Turn = True
            hand.Visible = False
            hand2.Visible = True
            If mnuOnePLayer.Checked = True Then
                tmAiMove.Enabled = True
            End If
        End If
    End If
    If NetGame = True Then
        NoGameMoves = 1
        Call Winsock.SendData(NoGameMoves)
        Call Winsock.SendData(NoGameMoves)
        Call Winsock.SendData(NoGameMoves)
        FlexGrid.Enabled = False
        NoGameMoves = 0
    End If
End If
If StopGame = 2 Then
    MsgBox "start a new game there are no more moves"
    FlexGrid.Enabled = False
    cmdMoves.Enabled = False
    tmAiMove.Enabled = False
    StopGame = 0
    Call Check
End If
End Sub
Private Sub FlexGrid_Click()
Dim Error As String
Dim Row As Integer, Col As Integer
Dim othellomoves As Integer
Dim Pos As Boolean
Dim Direction As Integer
Pos1 = False
NoGameMoves = 0
If Turn = True Then
    If mnuOnePLayer.Checked = True Then
        Exit Sub
    Else
        Row = FlexGrid.Row
        Col = FlexGrid.Col
        For CheckRow = Row - 1 To Row + 1
            For CheckCol = Col - 1 To Col + 1
                Direction = Direction + 1
                If CheckRow = Row And CheckCol = Col Then
                    Direction = Direction - 1
                End If
                If GridArray(CheckRow, CheckCol) = 1 Then
                    Pos = CheckPos(CheckRow, CheckCol, Direction, False)
                    If Pos1 = True Then
                        Exit For
                    End If
                End If
            Next CheckCol
            If Pos1 = True Then
                Exit For
            End If
        Next CheckRow
        If Pos1 = True Then
            If GridArray(Row, Col) = 0 Then
                Set FlexGrid.CellPicture = picWhite
                GridArray(Row, Col) = 2
                Turn = False
                StopGame = 0
            Else
                Load frmValid
                frmValid.Show
                frmOthello.Enabled = False
                FlexGrid.CellBackColor = FlexGrid.CellBackColor
                Exit Sub
            End If
        Else
            Load frmValid
            frmValid.Show
            frmOthello.Enabled = False
            Exit Sub
        End If
    End If
Else
    Row = FlexGrid.Row
    Col = FlexGrid.Col
    For CheckRow = Row - 1 To Row + 1
        For CheckCol = Col - 1 To Col + 1
            Direction = Direction + 1
            If CheckRow = Row And CheckCol = Col Then
                Direction = Direction - 1
            End If
            If GridArray(CheckRow, CheckCol) = 2 Then
                Pos = CheckPos(CheckRow, CheckCol, Direction, False)
                If Pos1 = True Then
                    Exit For
                End If
            End If
        Next CheckCol
        If Pos1 = True Then
            Exit For
        End If
    Next CheckRow
    If Pos1 = True Then
        If GridArray(Row, Col) = 0 Then
            Set FlexGrid.CellPicture = picBlack
            GridArray(Row, Col) = 1
            Turn = True
            StopGame = 0
        Else
            Load frmValid
            frmValid.Show
            Me.Enabled = False
            FlexGrid.CellBackColor = FlexGrid.CellBackColor
            Exit Sub
        End If
    Else
        Load frmValid
        frmValid.Show
        Me.Enabled = False
        Exit Sub
    End If
End If
Text1.Text = ""
Row = FlexGrid.Row
Col = FlexGrid.Col
If NetGame = True Then
    Call Winsock.SendData(Row)
    Call Winsock.SendData(Col)
    FlexGrid.Enabled = False
    If Turn = False Then
        hand2.Visible = True
        hand.Visible = False
    Else
        hand2.Visible = False
        hand.Visible = True
    End If
End If
RegRow = Row
RegCol = Col
othellomoves = Moves(Row, Col, 1, False)
'Call Print_Array
If mnuOnePLayer.Checked = True Then
    hand2.Visible = True
    hand.Visible = False
    If Turn = True Then
        tmAiMove.Enabled = True
    End If
Else
    Call Check
End If
End Sub
Sub Print_Array()
Dim Test As Boolean
'print designs a threw h
For HoldRow = -1 To MaxRows
    Text1 = Text1 & vbCrLf
    For HoldCol = -1 To MaxCols
        Text1 = Text1 & "   " & GridArray(HoldRow, HoldCol)
    Next HoldCol
Next HoldRow
End Sub

Private Sub Form_Load()

Dim Row As Integer, Col As Integer
Dim Size As Integer
hand.Visible = True
hand2.Visible = False
Turn = False
ClearGame = False
'used to set the array to 0 on start up
For Row = -1 To MaxRows
    For Col = -1 To MaxCols
        If Row = -1 Then
            GridArray(Row, Col) = 9
        End If
        If Col = -1 Then
            GridArray(Row, Col) = 9
        End If
        If Row = MaxRows Then
            GridArray(Row, Col) = 9
        End If
        If Col = MaxCols Then
            GridArray(Row, Col) = 9
        End If
    Next Col
Next Row
GridArray(4, 4) = 2
GridArray(3, 3) = 2
GridArray(3, 4) = 1
GridArray(4, 3) = 1
'used to set the back color of the grid to vbwhite
For Row = 0 To MaxRows - 1
    For Col = 0 To MaxCols - 1
        FlexGrid.Row = Row
        FlexGrid.Col = Col
        If GridArray(Row, Col) = 0 Then
            FlexGrid.CellBackColor = VbDarkGreen
        ElseIf GridArray(Row, Col) = 1 Then
            FlexGrid.CellBackColor = VbDarkGreen
            Set FlexGrid.CellPicture = picBlack
        ElseIf GridArray(Row, Col) = 2 Then
            FlexGrid.CellBackColor = VbDarkGreen
            Set FlexGrid.CellPicture = picWhite
        End If
    Next Col
Next Row
For Size = 0 To MaxRows - 1
    FlexGrid.RowHeight(Size) = 675
    FlexGrid.ColWidth(Size) = 675
Next Size
End Sub
Function Moves(Row1 As Integer, Col1 As Integer, Direction As Integer, valid As Boolean) As Integer
Dim Directrow As Integer, Row As Integer, Col As Integer, Directcol As Integer
Dim othellomoves As Integer
Select Case Direction
    Case 1
    'top left
        Directrow = Row1 - 1
        Directcol = Col1 - 1
    Case 2
    'top
        Directrow = Row1 - 1
        Directcol = Col1
    Case 3
    'top right
        Directrow = Row1 - 1
        Directcol = Col1 + 1
    Case 4
    'right
        Directrow = Row1
        Directcol = Col1 + 1
    Case 5
    'bottom right
        Directrow = Row1 + 1
        Directcol = Col1 + 1
    Case 6
    'bottom
        Directrow = Row1 + 1
        Directcol = Col1
    Case 7
    'bottom left
        Directrow = Row1 + 1
        Directcol = Col1 - 1
    Case 8
    'left
        Directrow = Row1
        Directcol = Col1 - 1
End Select
If Direction <= 8 Then
    If Turn = True Then
        Direction = Direction + 1
        If GridArray(Directrow, Directcol) = 2 Then
            Direction = Direction - 1
            othellomoves = Moves(Directrow, Directcol, Direction, False)
            If GridArray(NewRow, NewCol) = 1 And GridArray(Directrow, Directcol) = 2 Then
                valid = True
            End If
            If valid = True Then
                GridArray(Directrow, Directcol) = 1
                FlexGrid.Row = Directrow
                FlexGrid.Col = Directcol
                Set FlexGrid.CellPicture = picBlack
            End If
        ElseIf GridArray(Directrow, Directcol) = 1 Then
            othellomoves = Moves(RegRow, RegCol, Direction, False)
        ElseIf GridArray(Directrow, Directcol) = 0 Then
            othellomoves = Moves(RegRow, RegCol, Direction, False)
        ElseIf GridArray(Directrow, Directcol) = 9 Then
            othellomoves = Moves(RegRow, RegCol, Direction, False)
        End If
    ElseIf Turn = False Then
        Direction = Direction + 1
        If GridArray(Directrow, Directcol) = 1 Then
            Direction = Direction - 1
            othellomoves = Moves(Directrow, Directcol, Direction, False)
            If GridArray(NewRow, NewCol) = 2 And GridArray(Directrow, Directcol) = 1 Then
                valid = True
            End If
            If valid = True Then
                GridArray(Directrow, Directcol) = 2
                FlexGrid.Row = Directrow
                FlexGrid.Col = Directcol
                Set FlexGrid.CellPicture = picWhite
             End If
        ElseIf GridArray(Directrow, Directcol) = 2 Then
            othellomoves = Moves(RegRow, RegCol, Direction, False)
        ElseIf GridArray(Directrow, Directcol) = 0 Then
            othellomoves = Moves(RegRow, RegCol, Direction, False)
        ElseIf GridArray(Directrow, Directcol) = 9 Then
            othellomoves = Moves(RegRow, RegCol, Direction, False)
        End If
    End If
    NewRow = Directrow
    NewCol = Directcol
Else
    othellomoves = 1
End If
End Function
Sub Check()
Dim Row As Integer, Col As Integer, White As Integer, Black As Integer
Dim Row1 As Integer, Col1 As Integer
Dim Test As Boolean
Dim Winner As String, PiecesWinner As String, PiecesLoser As String
Winner = ""
PiecesWinner = ""
PiecesLoser = ""
For Row = 0 To MaxRows - 1
    For Col = 0 To MaxRows - 1
        If GridArray(Row, Col) = 1 Then
            Black = Black + 1
        End If
        If GridArray(Row, Col) = 2 Then
            White = White + 1
        End If
        If GridArray(Row, Col) = 0 Then
            Test = True
        Else
            For Row1 = 0 To MaxRows - 1
                For Col1 = 0 To MaxRows - 1
                    If GridArray(Row1, Col1) = 0 Then
                        Test = True
                        Exit For
                    End If
                Next Col1
                Exit For
            Next Row1
        End If
    Next Col
Next Row

lblwhite = White
lblblack = Black
If White = 0 Or Black = 0 Then
    If Black > White Then
        Winner = "Black"
        PiecesWinner = Black
        PiecesLoser = White
    ElseIf Black < White Then
        Winner = "White"
        PiecesWinner = White
        PiecesLoser = Black
    End If
    MsgBox (Winner & " wins " & PiecesWinner & " to " & PiecesLoser & vbCrLf & "no more moves clear the game")
    FlexGrid.Enabled = False
    cmdMoves.Enabled = False
    Exit Sub
End If
If Test = False Then
    If Black > White Then
        Winner = "Black"
        PiecesWinner = Black
        PiecesLoser = White
    ElseIf White > Black Then
        Winner = "White"
        PiecesWinner = White
        PiecesLoser = Black
    ElseIf Black = White Then
        Winner = "Tie"
        PiecesWinner = Black
        PiecesLoser = White
        MsgBox ("You Tied " & PiecesWinner & " to " & PiecesLoser)
    End If
    If Winner = "Black" Or Winner = "White" Then
        MsgBox (Winner & " wins " & PiecesWinner & " to " & PiecesLoser & vbCrLf & "no more moves clear the game")
    End If
    FlexGrid.Enabled = False
    cmdMoves.Enabled = False
    Exit Sub
End If
If Turn = True Then
    hand.Visible = False
    hand2.Visible = True
Else
    hand.Visible = True
    hand2.Visible = False
End If
End Sub

Private Sub mnuAbout_Click()
Load frmAbout
frmAbout.Show
frmOthello.Enabled = False
End Sub

Private Sub mnuClear_Click()
Dim Row As Integer, Col As Integer
If NetGame = True Then
    If chkServer.Value = 1 Then
        ClearGame = 1
        Call WinClear.SendData(ClearGame)
        Exit Sub
    Else
        Exit Sub
    End If
End If
For Row = 0 To MaxRows
    For Col = 0 To MaxCols
        GridArray(Row, Col) = 0
    Next Col
Next Row
FlexGrid.Clear
Text1.Text = ""
lblblack = 2
lblwhite = 2
FlexGrid.Enabled = True
cmdMoves.Enabled = True
Call Form_Load
End Sub
Private Sub mnuExit_Click()
End
End Sub
Function CheckPos(Row1 As Integer, Col1 As Integer, Direction As Integer, valid As Boolean) As Boolean
Dim Directrow As Integer, Row As Integer, Col As Integer, Directcol As Integer
Dim othellomoves As Integer
Select Case Direction
    Case 1
    'top left
        Directrow = Row1 - 1
        Directcol = Col1 - 1
    Case 2
    'top
        Directrow = Row1 - 1
        Directcol = Col1
    Case 3
    'top right
        Directrow = Row1 - 1
        Directcol = Col1 + 1
    Case 5
    'right
        Directrow = Row1
        Directcol = Col1 + 1
    Case 8
    'bottom right
        Directrow = Row1 + 1
        Directcol = Col1 + 1
    Case 7
    'bottom
        Directrow = Row1 + 1
        Directcol = Col1
    Case 6
    'bottom left
        Directrow = Row1 + 1
        Directcol = Col1 - 1
    Case 4
    'left
        Directrow = Row1
        Directcol = Col1 - 1
End Select
If Direction <= 8 Then
    If Turn = True Then
        If GridArray(Directrow, Directcol) = 2 Then
            NewRow = CheckRow
            NewCol = CheckCol
            othellomoves = CheckPos(Directrow, Directcol, Direction, False)
            If GridArray(NewRow, NewCol) = 1 And GridArray(Directrow, Directcol) = 2 And GridArray(Row1, Col1) = 1 Then
                Pos1 = True
            End If
        ElseIf GridArray(Directrow, Directcol) = 1 Then
            othellomoves = CheckPos(Directrow, Directcol, Direction, False)
        ElseIf GridArray(Directrow, Directcol) = 9 Then
            Pos1 = False
        End If
     ElseIf Turn = False Then
        If GridArray(Directrow, Directcol) = 1 Then
                NewRow = CheckRow
                NewCol = CheckCol
                othellomoves = CheckPos(Directrow, Directcol, Direction, False)
            If GridArray(NewRow, NewCol) = 2 And GridArray(Directrow, Directcol) = 1 And GridArray(Row1, Col1) = 2 Then
                    Pos1 = True
                End If
            ElseIf GridArray(Directrow, Directcol) = 2 Then
                othellomoves = CheckPos(Directrow, Directcol, Direction, False)
            ElseIf GridArray(Directrow, Directcol) = 9 Then
                Pos1 = False
            End If
    End If
Else
    othellomoves = 1
End If
End Function
Function CheckVal(Row1 As Integer, Col1 As Integer, Direction As Integer, valid As Boolean) As Boolean
Dim Directrow As Integer, Row As Integer, Col As Integer, Directcol As Integer
Dim othellomoves As Integer
Select Case Direction
    Case 1
    'top left
        Directrow = Row1 - 1
        Directcol = Col1 - 1
    Case 2
    'top
        Directrow = Row1 - 1
        Directcol = Col1
    Case 3
    'top right
        Directrow = Row1 - 1
        Directcol = Col1 + 1
    Case 5
    'right
        Directrow = Row1
        Directcol = Col1 + 1
    Case 8
    'bottom right
        Directrow = Row1 + 1
        Directcol = Col1 + 1
    Case 7
    'bottom
        Directrow = Row1 + 1
        Directcol = Col1
    Case 6
    'bottom left
        Directrow = Row1 + 1
        Directcol = Col1 - 1
    Case 4
    'left
        Directrow = Row1
        Directcol = Col1 - 1
End Select
If Direction <= 8 Then
    If Turn = True Then
        If GridArray(Directrow, Directcol) = 2 Then
            NewRow = CheckRow
            NewCol = CheckCol
            othellomoves = CheckVal(Directrow, Directcol, Direction, False)
            If GridArray(NewRow, NewCol) = 0 And GridArray(Directrow, Directcol) = 2 And GridArray(Row1, Col1) = 1 Then
                Pos1 = True
            End If
        ElseIf GridArray(Directrow, Directcol) = 1 Then
            If GridArray(Row1, Col1) = 2 Then
                Exit Function
            Else
            othellomoves = CheckVal(Directrow, Directcol, Direction, False)
            End If
        ElseIf GridArray(Directrow, Directcol) = 9 Then
            Pos1 = False
        End If
        
    ElseIf Turn = False Then
        If GridArray(Directrow, Directcol) = 1 Then
                NewRow = CheckRow
                NewCol = CheckCol
                othellomoves = CheckVal(Directrow, Directcol, Direction, False)
            If GridArray(NewRow, NewCol) = 0 And GridArray(Directrow, Directcol) = 1 And GridArray(Row1, Col1) = 2 Then
                    Pos1 = True
                End If
            ElseIf GridArray(Directrow, Directcol) = 2 Then
                If GridArray(Row1, Col1) = 1 Then
                    Exit Function
                Else
                    othellomoves = CheckVal(Directrow, Directcol, Direction, False)
                End If
            ElseIf GridArray(Directrow, Directcol) = 9 Then
                Pos1 = False
            End If
    End If
Else
    othellomoves = 1
End If
End Function
Function aimove(Row1 As Integer, Col1 As Integer, Direction As Integer, valid As Boolean) As Boolean
Dim Directrow As Integer, Row As Integer, Col As Integer, Directcol As Integer
Dim othellomoves As Integer
Pos1 = False
Select Case Direction
    Case 1
    'top left
        Directrow = Row1 - 1
        Directcol = Col1 - 1
    Case 2
    'top
        Directrow = Row1 - 1
        Directcol = Col1
    Case 3
    'top right
        Directrow = Row1 - 1
        Directcol = Col1 + 1
    Case 5
    'right
        Directrow = Row1
        Directcol = Col1 + 1
    Case 8
    'bottom right
        Directrow = Row1 + 1
        Directcol = Col1 + 1
    Case 7
    'down
        Directrow = Row1 + 1
        Directcol = Col1
    Case 6
    'bottom left
        Directrow = Row1 + 1
        Directcol = Col1 - 1
    Case 4
    'left
        Directrow = Row1
        Directcol = Col1 - 1
End Select
If Direction <= 8 Then
    If Turn = True Then
        Direction = Direction + 1
        If GridArray(Directrow, Directcol) = 1 Then
            Direction = Direction - 1
            othellomoves = aimove(Directrow, Directcol, Direction, False)
            Pieces = Pieces + 1
            If GridArray(NewRow, NewCol) = 2 And GridArray(Directrow, Directcol) = 1 And GridArray(CheckRow, CheckCol) = 0 Then
                Pos1 = True
            End If
        ElseIf GridArray(Directrow, Directcol) = 2 Then
            Pieces = 0
            If GridArray(Row1, Col1) = 1 Then
                NewRow = Directrow
                NewCol = Directcol
                Exit Function
            End If
        othellomoves = aimove(CheckRow, CheckCol, Direction, False)
        ElseIf GridArray(Directrow, Directcol) = 9 Then
            Pieces = 0
            othellomoves = aimove(CheckRow, CheckCol, Direction, False)
        ElseIf GridArray(Directrow, Directcol) = 0 Then
            Pieces = 0
            othellomoves = aimove(CheckRow, CheckCol, Direction, False)
        End If
    End If
    NewRow = Directrow
    NewCol = Directcol
Else
    othellomoves = 1
End If
End Function

Private Sub mnuHelp_Click()
frmHelp.Show
End Sub
Sub AiMove2()
Dim Direction As Integer, othellomoves As Integer
Dim ComputerMove As Integer
Pos1 = False
HoldMove = False
Pieces = 0
HoldPieces = 0
HoldRow1 = 0
HoldCol1 = 0
For CheckRow = 0 To MaxRows
    For CheckCol = 0 To MaxCols
        If GridArray(CheckRow, CheckCol) = 0 Then
            ComputerMove = aimove(CheckRow, CheckCol, 1, False)
            If Pos1 = True Then
                If (CheckRow = 0 And CheckCol = 0) Or (CheckRow = 0 And _
                CheckCol = 7) Or (CheckRow = 7 And CheckCol = 0) Or _
                (CheckRow = 7 And CheckCol = 7) Then
                    Pieces = Pieces * 10
                End If
                If (CheckRow = 0 And CheckCol <> 0 And CheckCol <> 7) _
                Or (CheckRow = 7 And CheckCol <> 0 And CheckCol <> 7) _
                Or (CheckCol = 0 And CheckRow <> 0 And CheckRow <> 7) _
                Or (CheckCol = 7 And CheckRow <> 0 And CheckRow <> 7) Then
                    Pieces = Pieces + 5
                End If
                If Pieces > HoldPieces Then
                    HoldPieces = Pieces
                    StopGame = 0
                    HoldMove = True
                    HoldRow1 = CheckRow
                    HoldCol1 = CheckCol
                ElseIf Pieces = HoldPieces Then
                    HoldPieces = Pieces
                    StopGame = 0
                    HoldMove = True
                    Randomize
                    Pieces = Pieces * Int((Rnd * 2) + 1)
                    If Pieces > HoldPieces Then
                        HoldRow1 = CheckRow
                        HoldCol1 = CheckCol
                    End If
                End If
            End If
            Pieces = 0
        End If
    Next CheckCol
Next CheckRow
If HoldMove = False Then
    StopGame = StopGame + 1
    If StopGame = 1 Then
        MsgBox "Computer has no move it is now your turn"
    End If
    If StopGame = 2 Then
        MsgBox "start a new game there are no more moves"
        FlexGrid.Enabled = False
        cmdMoves.Enabled = False
    Call Check
    End If
End If
Text1.Text = ""
Turn = False
If HoldMove = True Then
    othellomoves = AiMoves(HoldRow1, HoldCol1, 1, False)
    FlexGrid.Row = HoldRow1
    FlexGrid.Col = HoldCol1
    GridArray(HoldRow1, HoldCol1) = 2
    Set FlexGrid.CellPicture = picWhite
End If
'Call Print_Array
Call Check
tmAiMove.Enabled = False
End Sub
Function AiMoves(Row1 As Integer, Col1 As Integer, Direction As Integer, valid As Boolean) As Integer
Dim Directrow As Integer, Row As Integer, Col As Integer, Directcol As Integer
Dim othellomoves As Integer
Select Case Direction
    Case 1
    'top left
        Directrow = Row1 - 1
        Directcol = Col1 - 1
    Case 2
    'top
        Directrow = Row1 - 1
        Directcol = Col1
    Case 3
    'top right
        Directrow = Row1 - 1
        Directcol = Col1 + 1
    Case 4
    'right
        Directrow = Row1
        Directcol = Col1 + 1
    Case 5
    'bottom right
        Directrow = Row1 + 1
        Directcol = Col1 + 1
    Case 6
    'bottom
        Directrow = Row1 + 1
        Directcol = Col1
    Case 7
    'bottom left
        Directrow = Row1 + 1
        Directcol = Col1 - 1
    Case 8
    'left
        Directrow = Row1
        Directcol = Col1 - 1
End Select
If Direction <= 8 Then
    If Turn = False Then
        Direction = Direction + 1
        If GridArray(Directrow, Directcol) = 1 Then
            Direction = Direction - 1
            othellomoves = AiMoves(Directrow, Directcol, Direction, False)
            If GridArray(NewRow, NewCol) = 2 And GridArray(Directrow, Directcol) = 1 Then
                valid = True
            End If
            If valid = True Then
                GridArray(Directrow, Directcol) = 2
                FlexGrid.Row = Directrow
                FlexGrid.Col = Directcol
                Set FlexGrid.CellPicture = picWhite
             End If
        ElseIf GridArray(Directrow, Directcol) = 2 Then
            othellomoves = AiMoves(HoldRow1, HoldCol1, Direction, False)
        ElseIf GridArray(Directrow, Directcol) = 0 Then
            othellomoves = AiMoves(HoldRow1, HoldCol1, Direction, False)
        ElseIf GridArray(Directrow, Directcol) = 9 Then
            othellomoves = AiMoves(HoldRow1, HoldCol1, Direction, False)
        End If
    End If
    NewRow = Directrow
    NewCol = Directcol
Else
    othellomoves = 1
End If
End Function

Private Sub WinChat_Close()
Call WinChat.Close
Command1.Enabled = False
Text1.Enabled = False
End Sub

Private Sub WinChat_Connect()
lbltalk.Visible = True
 lbltalk.Caption = "Connection from IP address: " & _
      WinChat.RemoteHostIP & vbCrLf & "Port #: " & _
      WinChat.RemotePort & vbCrLf & vbCrLf
Command1.Enabled = True
Text1.Enabled = True
Timer1.Enabled = True
FlexGrid.Enabled = True
End Sub

Private Sub WinChat_ConnectionRequest(ByVal requestID As Long)
If WinChat.State <> sckClosed Then
 WinChat.Close
End If
 Call WinChat.Accept(requestID) 'accepts connection
 'display connection status
 lbltalk.Caption = "Connection from IP address: " & _
      WinChat.RemoteHostIP & vbCrLf & "Port #: " & _
      WinChat.RemotePort & vbCrLf & vbCrLf
Command1.Enabled = True
Text1.Enabled = True
FlexGrid.Enabled = True
End Sub

Private Sub WinChat_DataArrival(ByVal bytesTotal As Long)
Dim strMessage As String
MediaPlayer1.FileName = App.Path & "/talkbeg.wav"
MediaPlayer1.Play
Messages = Messages + 1
Call WinChat.GetData(strMessage)
lbltalk.Caption = lbltalk.Caption & Trim$(strMessage) & vbCrLf
lbltalk.Visible = True
Timer1.Interval = Messages * 5000
Timer1.Enabled = True
End Sub

Private Sub WinChat_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   Dim result As Integer
   result = MsgBox(Source & ": " & Description, _
      vbOKOnly, "TCP/IP Error")
   chkServer.Value = 0
   chkClient.Value = 0
End Sub

Private Sub WinClear_Close()
Call WinClear.Close
Command1.Enabled = False
Text1.Enabled = False
NetGame = False
End Sub

Private Sub WinClear_ConnectionRequest(ByVal requestID As Long)
If WinClear.State <> sckClosed Then
 WinClear.Close
End If
 Call WinClear.Accept(requestID) 'accepts connection
 'display connection status
 lbltalk.Caption = "Connection from IP address: " & _
      WinClear.RemoteHostIP & vbCrLf & "Port #: " & _
      WinClear.RemotePort & vbCrLf & vbCrLf
Command1.Enabled = True
Text1.Enabled = True
End Sub

Private Sub WinClear_DataArrival(ByVal bytesTotal As Long)
Dim Question As String
Dim Row As Integer, Col As Integer
Call WinClear.GetData(ClearGame)
If ClearGame = 1 Then
    If chkServer.Value = 1 Then
        Question = MsgBox("You have cleared the game")
    Else
        Question = MsgBox("Do you want to clear the game?", vbYesNo)
    End If
    If Question = vbYes Then
        For Row = 0 To MaxRows
            For Col = 0 To MaxCols
                GridArray(Row, Col) = 0
            Next Col
        Next Row
        FlexGrid.Clear
        Text1.Text = ""
        lblblack = 2
        lblwhite = 2
        FlexGrid.Enabled = True
        cmdMoves.Enabled = True
        Call Form_Load
        ClearGame = 1
        Call WinClear.SendData(ClearGame)
    ElseIf Question = vbOK Then
        For Row = 0 To MaxRows
            For Col = 0 To MaxCols
                GridArray(Row, Col) = 0
            Next Col
        Next Row
        FlexGrid.Clear
        Text1.Text = ""
        lblblack = 2
        lblwhite = 2
        FlexGrid.Enabled = True
        cmdMoves.Enabled = True
        Call Form_Load
    Else
        Exit Sub
    End If
End If
ClearGame = 0
End Sub

Private Sub WinClear_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   Dim result As Integer
   result = MsgBox(Source & ": " & Description, _
      vbOKOnly, "TCP/IP Error")
   chkServer.Value = 0
   chkClient.Value = 0
End Sub

Private Sub Winsock_Close()
Call Winsock.Close
Command1.Enabled = False
Text1.Enabled = False
lbltalk.Visible = True
Timer1.Enabled = True
lbltalk.Caption = lbltalk.Caption & "Connection closed" & vbCrLf
chkClient.Value = 0
chkServer.Value = 0
FlexGrid.Enabled = False
NetGame = False
End Sub

Private Sub Winsock_ConnectionRequest(ByVal requestID As Long)
If Winsock.State <> sckClosed Then
 Winsock.Close
End If
lbltalk.Visible = True
Timer1.Enabled = True
 Call Winsock.Accept(requestID) 'accepts connection
 'display connection status
 lbltalk.Caption = "Connection from IP address: " & _
      Winsock.RemoteHostIP & vbCrLf & "Port #: " & _
      Winsock.RemotePort & vbCrLf & vbCrLf
Command1.Enabled = True
Text1.Enabled = True
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
Dim Row As Integer, Col As Integer
Dim othellomoves As Integer
If Turn = True Then
    Turn = False
    hand2.Visible = False
    hand.Visible = True
Else
    Turn = True
    hand2.Visible = True
    hand.Visible = False
End If
Call Winsock.GetData(Row)
Call Winsock.GetData(Col)
Call Winsock.GetData(NoGameMoves)
If NoGameMoves = 1 Then
    NoGameMoves = 0
    FlexGrid.Enabled = True
    Exit Sub
End If
RegRow = Row
RegCol = Col
FlexGrid.Row = Row
FlexGrid.Col = Col
If Turn = True Then
    GridArray(Row, Col) = 1
    Set FlexGrid.CellPicture = picBlack
Else
    GridArray(Row, Col) = 2
    Set FlexGrid.CellPicture = picWhite
End If
othellomoves = Moves(Row, Col, 1, False)
FlexGrid.Enabled = True
Call Check
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim result As Integer
result = MsgBox(Source & ": " & Description, _
   vbOKOnly, "TCP/IP Error")
   chkServer.Value = 0
   chkClient.Value = 0
End Sub
