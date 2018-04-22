VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Level Properties"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "More level properties"
      Height          =   2295
      Left            =   3120
      TabIndex        =   21
      Top             =   2040
      Width           =   2895
      Begin VB.TextBox txtFruit 
         Height          =   285
         Left            =   1440
         TabIndex        =   23
         Text            =   "1"
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   1920
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Fruit type:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   27
         Top             =   360
         Width           =   255
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Relative Size Preview"
      Height          =   1815
      Left            =   3120
      TabIndex        =   8
      Top             =   120
      Width           =   2895
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         Height          =   1455
         Left            =   720
         ScaleHeight     =   1395
         ScaleWidth      =   1395
         TabIndex        =   9
         Top             =   240
         Width           =   1455
         Begin VB.Label lblPreviewRect 
            BackColor       =   &H0080FF80&
            Height          =   135
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   135
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Level appearance"
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   2895
      Begin VB.PictureBox picColor 
         BackColor       =   &H000040C0&
         Height          =   255
         Index           =   0
         Left            =   1080
         ScaleHeight     =   195
         ScaleWidth      =   555
         TabIndex        =   25
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdChgColor 
         Caption         =   "&Change..."
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   24
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdChgColor 
         Caption         =   "&Change..."
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   18
         Top             =   1320
         Width           =   855
      End
      Begin VB.PictureBox picColor 
         BackColor       =   &H000040C0&
         Height          =   255
         Index           =   3
         Left            =   1080
         ScaleHeight     =   195
         ScaleWidth      =   555
         TabIndex        =   17
         Top             =   1320
         Width           =   615
      End
      Begin VB.PictureBox picColor 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   2
         Left            =   1080
         ScaleHeight     =   195
         ScaleWidth      =   555
         TabIndex        =   15
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdChgColor 
         Caption         =   "&Change..."
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.PictureBox picColor 
         BackColor       =   &H00FF00FF&
         Height          =   255
         Index           =   1
         Left            =   1080
         ScaleHeight     =   195
         ScaleWidth      =   555
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdChgColor 
         Caption         =   "&Change..."
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Pellet:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Fill:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Edge (shadow):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Edge (light):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dimensions"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2895
      Begin VB.HScrollBar scrWidth 
         Height          =   255
         Left            =   120
         Max             =   300
         Min             =   21
         TabIndex        =   4
         Top             =   600
         Value           =   21
         Width           =   2655
      End
      Begin VB.HScrollBar scrHeight 
         Height          =   255
         Left            =   120
         Max             =   300
         Min             =   23
         TabIndex        =   3
         Top             =   1320
         Value           =   23
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   4320
         TabIndex        =   20
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblWidth 
         Caption         =   "Width: x"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblHeight 
         Caption         =   "Height: y"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdChgColor_Click(Index As Integer)
    
    frmProperties.CommonDialog1.ShowColor
    picColor(Index).BackColor = CommonDialog1.Color
    
        Select Case Index
                
            Case 0
                pellet_Color = CommonDialog1.Color
                frmLvlEdit.GetCrossRef
                
            Case 1
                edge_LightColor = CommonDialog1.Color
                frmLvlEdit.GetCrossRef
                
            Case 2
                edge_ShadowColor = CommonDialog1.Color
                frmLvlEdit.GetCrossRef
                
            Case 3
                fill_Color = CommonDialog1.Color
                frmLvlEdit.GetCrossRef
        
        End Select
        
    frmLvlEdit.PopulateTileBox
    frmLvlEdit.DrawScreen


End Sub

Private Sub cmdOk_click()

Dim tResult

If scrWidth.Value + scrHeight.Value > 1100 Then
    tResult = MsgBox("Warning: The dimensions you specified result in an enormous map size, which will result in a large file size and long delays in saving and loading! Are you sure you want to make this adjustment?", vbYesNoCancel, Me.Caption)
    
    If Not tResult = 6 Then
        Exit Sub
    End If

End If

If scrWidth.Value Mod 2 = 0 Or scrHeight.Value Mod 2 = 0 Then
    tResult = MsgBox("Warning: One or more of the specified dimensions is an even number." & vbCrLf & vbCrLf & "This may cause undesired results in the center of the board with symmetric map editing turned on. Are you sure you want to make this adjustment?", vbYesNoCancel, Me.Caption)
    
    If Not tResult = 6 Then
        Exit Sub
    End If
    
End If


lvlWidth = scrWidth.Value
lvlHeight = scrHeight.Value

frmLvlEdit.DrawScreen
frmLvlEdit.ResizeScrollbars

Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()

    scrWidth.Value = lvlWidth
    scrHeight.Value = lvlHeight
    
    picColor(0).BackColor = pellet_Color
    picColor(1).BackColor = edge_LightColor
    picColor(2).BackColor = edge_ShadowColor
    picColor(3).BackColor = fill_Color
    
    txtFruit.Text = fruit_Type
    ReDrawIt
    
End Sub


Private Sub scrHeight_Scroll()
scrHeight_Change
End Sub

Private Sub scrWidth_Change()
lblWidth.Caption = scrWidth.Value
ReDrawIt
End Sub

Private Sub scrHeight_Change()
lblHeight.Caption = scrHeight.Value
ReDrawIt
End Sub

Private Sub scrWidth_Scroll()
scrWidth_Change
End Sub

Private Function ReDrawIt()
lblPreviewRect.Width = (scrWidth.Value / scrWidth.Max) * Picture1.Width
lblPreviewRect.Height = (scrHeight.Value / scrHeight.Max) * Picture1.Height
lblPreviewRect.Move Picture1.Width / 2 - lblPreviewRect.Width / 2, Picture1.Height / 2 - lblPreviewRect.Height / 2

Label2.Caption = (scrWidth.Value * 16) & " x " & (scrHeight.Value * 16)

End Function

Private Sub txtFruit_Change()
Image1.Picture = Nothing

On Error Resume Next
fruit_Type = txtFruit.Text

Image1.Picture = LoadPicture(App.Path & "\res\sprite\fruit " & fruit_Type & ".gif")
End Sub
