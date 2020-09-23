VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Value List Generator"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4920
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5530
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":08CA
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ncards As Long, nnums As Long, sp As String

Private Sub Form_Load()
Dim num() As String, yes As Boolean
If frmMR.see = True Then
    ncards = 6
Else
    ncards = InputBox("Number of cards?", "How Many?")
End If
If IsNumeric(ncards) = False Then End
nnums = 2 ^ Val(ncards)
For j = 0 To Val(ncards) - 1
    List1.AddItem "Card " & j + 1
    List1.AddItem ""
    sp = ""
    For i = 1 To nnums - 1
        If (i And (2 ^ j)) = (2 ^ j) Then
            If Len(List1.List(List1.ListCount - 1)) > 0 Then sp = ", "
            List1.List(List1.ListCount - 1) = List1.List(List1.ListCount - 1) & sp & i
        End If
    Next i
    num() = Split(List1.List(List1.ListCount - 1), ", ")
    yes = False
    For i = 0 To UBound(num)
        If frmMR.seeNum = num(i) Then num(i) = "**" & num(i) & "**": yes = True
    Next
    List1.List(List1.ListCount - 1) = Join(num, ", ")
    List1.List(List1.ListCount - 2) = List1.List(List1.ListCount - 2) & ": You said " & IIf(yes, "yes.", "no.")
    If j < Val(ncards) - 1 Then List1.AddItem ""
Next j
ListSave List1, App.Path & "\list.txt"
RichTextBox1.LoadFile App.Path & "\list.txt"
End Sub

Private Sub ListSave(List As ListBox, FilePath As String)
'Save all data in a list box
On Error GoTo error
Dim i As Integer
Dim Directory As String
Directory$ = FilePath
Open Directory$ For Output As #1
    For i = 0 To List.ListCount - 1
        Print #1, List.List(i)
    Next i
Close #1
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error!"
End Sub
