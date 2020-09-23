VERSION 5.00
Begin VB.Form frmMR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mind Reader"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   3015
   Icon            =   "frmMR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   3015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Start"
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   14
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Text            =   "Click Start"
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   6
      Text            =   "1"
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Index           =   0
      ItemData        =   "frmMR.frx":08CA
      Left            =   2880
      List            =   "frmMR.frx":08CC
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Index           =   1
      ItemData        =   "frmMR.frx":08CE
      Left            =   2880
      List            =   "frmMR.frx":08D0
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Index           =   2
      ItemData        =   "frmMR.frx":08D2
      Left            =   2880
      List            =   "frmMR.frx":08D4
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Index           =   3
      ItemData        =   "frmMR.frx":08D6
      Left            =   2880
      List            =   "frmMR.frx":08D8
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Index           =   4
      ItemData        =   "frmMR.frx":08DA
      Left            =   2880
      List            =   "frmMR.frx":08DC
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Index           =   5
      ItemData        =   "frmMR.frx":08DE
      Left            =   2880
      List            =   "frmMR.frx":08E0
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2880
      TabIndex        =   7
      Text            =   "2"
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2880
      TabIndex        =   8
      Text            =   "4"
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   2880
      TabIndex        =   9
      Text            =   "8"
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   2880
      TabIndex        =   10
      Text            =   "16"
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   2880
      TabIndex        =   11
      Text            =   "32"
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Restart"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   16
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   17
      Top             =   360
      Width           =   2535
   End
   Begin VB.Menu mnuGen 
      Caption         =   "Card List Generator"
   End
   Begin VB.Menu mnuSee 
      Caption         =   "See Why You Got This"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, ttl, cn, play As Integer
Public see As Boolean, seeNum As Integer

Private Sub Command1_Click()
If play = 1 Then
    ttl = ttl + Text1(cn)
    Call NextCard
End If
End Sub

Private Sub Command2_Click()
If play = 1 Then
    Call NextCard
End If
End Sub

Private Sub Command4_Click()
play = 0
cn = 0
Label1.Caption = "Think of a # between" & vbCrLf & "0 and 63. Six cards" & vbCrLf & "will be displayed."
Text2.Text = "Click Start"
mnuSee.Visible = False
mnuGen.Visible = True
Command3.Enabled = True
Command3.Visible = True
End Sub

Private Sub Command3_Click()
cn = 0
play = 1
ttl = 0
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = False
Label1.Caption = ""
    Label2.Caption = List1(cn).List(0)
For a = 1 To 5
    Label1.Caption = Label1.Caption & vbCrLf & List1(0).List(a)
Next a
Text2.Text = "Is your number on the card above?"
End Sub

Private Sub NextCard()
cn = cn + 1
If cn = 6 Then
    Call EndIt
Else
Label1.Caption = ""
    Label2.Caption = List1(cn).List(0)
For a = 1 To 5
    Label1.Caption = Label1.Caption & vbCrLf & List1(cn).List(a)
Next a
Text2.Text = "Is your number on the card above?"
End If
End Sub

Private Sub EndIt()
Dim one, two, total As String
Label2.Caption = 0
play = 2
If ttl < 10 Then
    total = " " + CStr(ttl)
Else
    total = CStr(ttl)
End If
one = Left(total, 1)
two = Right(total, 1)
Label2.Caption = ""
Label1.Caption = "THE NUMBER YOU WERE" & vbCrLf & "THINKING OF WAS " & Trim(one) & two & "."
seeNum = Trim(one) & two
mnuGen.Visible = False
mnuSee.Visible = True
Text2.Text = "Click Restart to play again"
Command3.Visible = False
Command1.Enabled = False
Command2.Enabled = False
Command4.Enabled = True
Command4.Visible = True
End Sub

Private Sub Form_Load()
Dim j As Long, i As Long, k As Long, l As String
see = True
Label1.Caption = "" & vbCrLf & "Think of a # between" & vbCrLf & "1 and 63. Six cards" & vbCrLf & "will be displayed."
'###Credit for this code goes to Lefteris Eleftheriades###
For j = 0 To 5
    List1(j).List(0) = "Card " & j + 1
    k = 1
    For i = 1 To 63
        If (i And (2 ^ j)) = (2 ^ j) Then
            If Len(List1(j).List(k)) = 0 Then
                If i < 10 Then
                    l = Str(i) & " "
                Else
                    l = i
                End If
            Else
                l = Str(i) & String(3 - Len(Str(i)), " ")
            End If
            List1(j).List(Int(k)) = List1(j).List(k) & l
            If Len(List1(j).List(k)) >= 22 Then k = k + 1
        End If
    Next
Next j
'##########################################################
End Sub

Private Sub List1_Click(Index As Integer)
    Text2.SetFocus
End Sub

Private Sub mnuGen_Click()
see = False
Form1.Show
see = True
End Sub

Private Sub mnuSee_Click()
Form1.Show
End Sub
