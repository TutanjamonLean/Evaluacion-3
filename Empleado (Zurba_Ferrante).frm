VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   14505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Iniciar Sesion"
      Height          =   615
      Left            =   9840
      TabIndex        =   2
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   9840
      TabIndex        =   1
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   9840
      TabIndex        =   0
      Top             =   3960
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pepito As New Conexion
Dim Usuario, Contra, nombre As String
Private Sub Command1_Click()
    
    Usuario = Text1.Text
    Contra = Text2.Text


    If Pepito.Chekeo(CStr(Usuario), CStr(Contra)) = True Then
        MsgBox "Bienvenido al Sistema Leandro ", vbInformation
        Form2.Show
        Unload Me

    Else
    
        MsgBox "Incorrecto", vbCritical

    End If



End Sub

Private Sub Form_Activate()

    Pepito.Conector ("Empleado")
    
    Form1.Width = 13960
    Form1.Height = 9840

End Sub



