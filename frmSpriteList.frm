VERSION 5.00
Begin VB.Form frmSpriteList 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sprite List"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Properties"
      Height          =   4695
      Left            =   3120
      TabIndex        =   9
      Top             =   1920
      Width           =   4215
      Begin VB.CommandButton cmdNote 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Note"
         Height          =   375
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtProp 
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   22
         Text            =   "x"
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox txtProp 
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Text            =   "x"
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox txtProp 
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Text            =   "x"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox txtProp 
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Text            =   "x"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox txtProp 
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Text            =   "x"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtProp 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Text            =   "x"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblProp 
         Caption         =   "[Property label]"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   21
         Top             =   3960
         Width           =   3855
      End
      Begin VB.Label lblProp 
         Caption         =   "[Property label]"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   3240
         Width           =   3855
      End
      Begin VB.Label lblProp 
         Caption         =   "[Property label]"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   2520
         Width           =   3855
      End
      Begin VB.Label lblProp 
         Caption         =   "[Property label]"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label lblProp 
         Caption         =   "[Property label]"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label lblProp 
         Caption         =   "[Property label]"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selected sprite"
      Height          =   1695
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   2640
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
      Begin VB.PictureBox picSel 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   240
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   4
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblTileDesc 
         Caption         =   "Description. (0)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   7
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label lblTileName 
         Caption         =   "tilename"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblLoc 
         Caption         =   "Located at (x, x)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   1200
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sprites in use"
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.ListBox lstSprites 
         Height          =   6300
         ItemData        =   "frmSpriteList.frx":0000
         Left            =   120
         List            =   "frmSpriteList.frx":0007
         TabIndex        =   1
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label lblSpriteTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "25 sprites total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmSpriteList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selSpriteName As String
Dim selNote As String

Private Sub cmdDelete_Click()
    selSpriteIndex = lstSprites.ListIndex
    
    DeleteSelectedSprite
    frmLvlEdit.DrawScreen
    
    PopulateSpriteList
    
    If Not selSpriteIndex = 0 Then
        If Not selSpriteIndex < numOfSprites Then
            selSpriteIndex = selSpriteIndex - 1
        End If
    End If
    
    On Error Resume Next
    lstSprites.ListIndex = selSpriteIndex
End Sub

Private Sub cmdNote_Click()
MsgBox selNote, vbInformation, "Note about selected sprite"
End Sub

Private Sub cmdOk_click()
Unload Me
End Sub

Private Sub Form_Load()

PopulateSpriteList

End Sub

Private Function PopulateSpriteList()

If numOfSprites = 0 Then
    Unload Me
    Exit Function
End If

lblSpriteTotal.Caption = numOfSprites & " sprites total"

lstSprites.Clear
For i = 0 To numOfSprites - 1
    lstSprites.AddItem Tiles(tileIndex(Sprites(i, 2)), 2) & " (" & Sprites(i, 0) & ", " & Sprites(i, 1) & ")"
Next i

End Function


Private Sub lstsprites_Click()

'introuce a temporary variable to save typing!
Dim tIndex As Integer
tIndex = tileIndex(Sprites(lstSprites.ListIndex, 2))

'--

lblTileName.Caption = Tiles(tIndex, 2)
selSpriteName = Tiles(tIndex, 2)
lblTileDesc.Caption = Tiles(tIndex, 3) & " (#" & Tiles(tIndex, 1) & ")"
lblTileDesc.Caption = UCase(Mid(lblTileDesc.Caption, 1, 1)) & Mid(lblTileDesc.Caption, 2, Len(lblTileDesc.Caption) - 1) & "."
lblLoc.Caption = "Located at (" & Sprites(lstSprites.ListIndex, 0) & ", " & Sprites(lstSprites.ListIndex, 1) & ")"

picSel.Picture = LoadPicture(App.Path & "\res\tiles\" & Tiles(tIndex, 2) & ".bmp")




For i = 0 To 5
    lblProp(i).Caption = vbNullString
Next i
selNote = vbNullString

Select Case selSpriteName 'show appropriate properties
    Case "pipeexit"
        lblProp(0).Caption = "Exits to level:"
        lblProp(1).Caption = "Entrance row"
        lblProp(2).Caption = "Entrance column"
        lblProp(3).Caption = "Entrance action (see note)"
        selNote = "Entrance action:" _
            & vbCrLf & "0: use present pipe" _
            & vbCrLf & "1: falling" _
            & vbCrLf & "2: nothing (appear)"
End Select

For i = 0 To 5
    If Not lblProp(i).Caption = vbNullString Then
        txtProp(i).Text = Sprites(lstSprites.ListIndex, i + 3)
        txtProp(i).Visible = True
    Else
        txtProp(i).Text = vbNullString
        txtProp(i).Visible = False
    End If
Next i

If Not selNote = vbNullString Then
    cmdNote.Visible = True
Else
    cmdNote.Visible = False
End If




End Sub

Private Sub lstSprites_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    cmdDelete_Click
End If
End Sub

Private Sub txtProp_Change(Index As Integer)
On Error GoTo NotInteger
    Sprites(lstSprites.ListIndex, Index + 3) = txtProp(Index).Text
    If Not txtProp(Index).BackColor = RGB(255, 255, 255) Then txtProp(Index).BackColor = RGB(255, 255, 255)
    Exit Sub
    
NotInteger:
    txtProp(Index).BackColor = RGB(255, 150, 100)
    
End Sub
