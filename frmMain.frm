VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error Dictionary"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9045
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Look up error number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   5460
      TabIndex        =   5
      Top             =   3885
      Width           =   2220
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   7
         Top             =   420
         Width           =   750
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Search"
         Enabled         =   0   'False
         Height          =   330
         Left            =   1155
         TabIndex        =   6
         Top             =   420
         Width           =   750
      End
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7875
      TabIndex        =   2
      Top             =   4515
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Load Errors"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7875
      TabIndex        =   1
      Top             =   3990
      Width           =   1065
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3570
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   8835
   End
   Begin VB.Label Label3 
      Caption         =   "To look up an error using the error number, type the error number into the text box above and click search."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      TabIndex        =   4
      Top             =   4410
      Width           =   5160
   End
   Begin VB.Label Label2 
      Caption         =   "Click load to load all the error numbers, with their descriptions, into the listbox to the left."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      TabIndex        =   3
      Top             =   3885
      Width           =   5160
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1.Text = "" Then Exit Sub
    If errDescription(Text1.Text) = "" Then
        MsgBox "Sorry, I do not have that error number in my dictionary", vbInformation + vbOKOnly, "Error Not Found"
    Else
        MsgBox errDescription(Text1.Text), , "Error Description"
    End If
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
   Text1.SetFocus

End Sub

Private Sub Command2_Click()

    Dim num

    List1.Clear
    Me.Caption = "Error Dictionary - Loading error descriptions..."
    
    num = 100 / 11031
    
    For i = 0 To 11031
        Me.Caption = "Error Dictionary - " & Left(num * i, 4) & "%"
        
        If Len(errDescription(i)) > 0 Then
            List1.AddItem Str(i) & "= " & errDescription(i)
        End If
    Next i
    
    
    Me.Caption = "Error Dictionary (" & List1.ListCount & " errors listed)"
    Command1.Enabled = True
    Command1.Default = True
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub List1_DblClick()
    MsgBox List1.Text, , "Error Description"
End Sub

Private Sub Text1_Change()
    If Len(Text1.Text) > 0 Then
        Command1.Enabled = True
        Command1.Default = True
    Else
        Command1.Enabled = False
        Command1.Default = False
        Command2.Default = True
    End If
End Sub
