VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   11400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   23745
   LinkTopic       =   "Form3"
   ScaleHeight     =   11400
   ScaleWidth      =   23745
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3375
      Left            =   4080
      TabIndex        =   8
      Top             =   3480
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   5953
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Index           =   4
      Left            =   4080
      TabIndex        =   7
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Index           =   3
      Left            =   6120
      TabIndex        =   6
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   735
      Index           =   2
      Left            =   6840
      TabIndex        =   5
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   735
      Index           =   1
      Left            =   6840
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Index           =   0
      Left            =   6840
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   735
      Index           =   2
      Left            =   4080
      TabIndex        =   2
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   735
      Index           =   1
      Left            =   4080
      TabIndex        =   1
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   0
      Left            =   4080
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Height          =   735
      Index           =   6
      Left            =   19080
      TabIndex        =   15
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Height          =   735
      Index           =   5
      Left            =   19080
      TabIndex        =   14
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Height          =   735
      Index           =   4
      Left            =   9000
      TabIndex        =   13
      Top             =   6960
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Height          =   735
      Index           =   3
      Left            =   12120
      TabIndex        =   12
      Top             =   6960
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Height          =   735
      Index           =   2
      Left            =   14880
      TabIndex        =   11
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Height          =   735
      Index           =   1
      Left            =   11880
      TabIndex        =   10
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Height          =   735
      Index           =   0
      Left            =   8880
      TabIndex        =   9
      Top             =   960
      Width           =   2895
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)

    If Command1(0).Value = True Then
                
        




    End If

End Sub
