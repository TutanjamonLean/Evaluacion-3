VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form2 
   BackColor       =   &H8000000C&
   Caption         =   "Form2"
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19710
   LinkTopic       =   "Form2"
   ScaleHeight     =   10080
   ScaleWidth      =   19710
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   4080
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   10095
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   2
         Left            =   2520
         TabIndex        =   15
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   1
         Left            =   2520
         TabIndex        =   14
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   0
         Left            =   2520
         TabIndex        =   13
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Mostrar"
         Height          =   615
         Left            =   6120
         TabIndex        =   12
         Top             =   2880
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Eliminar"
         Height          =   615
         Left            =   6120
         TabIndex        =   11
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Modificar"
         Height          =   615
         Left            =   6120
         TabIndex        =   10
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar"
         Height          =   615
         Left            =   6120
         TabIndex        =   9
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Height          =   615
         Left            =   1800
         TabIndex        =   8
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000001&
         Caption         =   "Cuit:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   840
         TabIndex        =   7
         Top             =   3600
         Width           =   3855
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000001&
         Caption         =   "Direccion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   840
         TabIndex        =   6
         Top             =   2760
         Width           =   3855
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000001&
         Caption         =   "Apellido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   840
         TabIndex        =   5
         Top             =   1920
         Width           =   3855
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000001&
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   840
         TabIndex        =   4
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   "Datos del Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   5055
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grilla 
      Height          =   3495
      Left            =   480
      TabIndex        =   1
      Top             =   6360
      Visible         =   0   'False
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   6165
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      ItemData        =   "Clientes(Zueba_Ferrante).frx":0000
      Left            =   720
      List            =   "Clientes(Zueba_Ferrante).frx":0002
      TabIndex        =   0
      Top             =   1080
      Width           =   3015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ConductorCli As New CLientes
Dim nom, ape, dire, CUIT As String
Dim nuevoNom, nuevoApe, nuevoDire As String
Private Sub Command1_Click()

    nom = Text1(0).Text
    ape = Text1(1).Text
    dire = Text1(2).Text
    CUIT = Text4.Text
    
    
    If nom = "" Or ape = "" Or dire = "" Or CUIT = "" Then
        
        MsgBox "Todos los campos son obligatorios", vbCritical, "Error"
        
        Exit Sub
    End If

    If ConductorCli.Añadir(CStr(nom), CStr(ape), CStr(dire), CStr(CUIT)) = True Then
        
        MsgBox "añadido al sistema", vbInformation, "Exito"
        
        
            With Grilla
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = "Cliente"
                .TextMatrix(.Rows - 1, 1) = nom
                .TextMatrix(.Rows - 1, 2) = ape
                .TextMatrix(.Rows - 1, 3) = dire
                .TextMatrix(.Rows - 1, 4) = CUIT
            End With

              
    End If
    
End Sub
Private Sub Command2_Click()

    nuevoNom = Text1(0).Text
    nuevoApe = Text1(1).Text
    nuevoDire = Text1(2).Text

    If CStr(nuevoNom) = "" Or CStr(nuevoApe) = "" Or CStr(nuevoDire) = "" Or CStr(CUIT) = "" Then
        
        MsgBox "Todos los campos son obligatorios", vbCritical, "Error"
        
        Exit Sub
    End If


    If ConductorCli.ModificarCampo(CStr(nuevoNom), CStr(nuevoApe), CStr(nuevoDire), CStr(CUIT)) Then
    
        
        MsgBox "no esta ningun CUIT", vbCritical
        
    Else
    
        MsgBox "Registro modificado", vbInformation, "exito"

    End If
    


End Sub
Private Sub ActualizarGrilla(cuitGrilla As String, nomGrilla As String, apeGrilla As String, dire As String)
Dim fila As Integer
Dim search As Boolean
    
    
    
End Sub
Private Sub Command3_Click()
Dim resultado As Boolean

    If CUIT = "" Then
    
        MsgBox "Campo obligatorio", vbCritical, "Error"
        
    End If


    If MsgBox("¿Seguro de eliminar este campo?", vbYesNo + vbQuestion, "Confirmar") = vbNo Then
        Exit Sub
    End If
    
    
    resultado = ConductorCli.BorrarDato(CUIT)
        
        
    If resultado = True Then
        MsgBox "Campo eliminado", vbInformation, "Exito"
        
    Else
    
        MsgBox "No se encontro CUIt", vbCritical, "Error"
        
    End If

End Sub

Private Sub Form_Load()
    
    ConductorCli.Conector2 ("Clientes")
    
    List1.AddItem "cliente"
    List1.AddItem "producto"
    
   
   With Grilla
        
        .Rows = 1
        .Cols = 5
        .ColAlignment(1) = 0
        .ColAlignment(2) = 0
        .ColAlignment(3) = 0
        .ColAlignment(4) = 0
        .TextMatrix(0, 0) = "Cliente"
        .TextMatrix(0, 1) = "Nombre"
        .TextMatrix(0, 2) = "Apellido"
        .TextMatrix(0, 3) = "Direccion"
        .TextMatrix(0, 4) = "Cuit"
        .ColWidth(1) = 3000
        .ColWidth(2) = 3000
        .ColWidth(3) = 3000
        .ColWidth(4) = 2000
   
   
   
   End With
    

End Sub

Private Sub List1_Click()

   If List1.ListIndex = 0 Then
        
        Frame1.Visible = True
        Grilla.Visible = True
        
   End If
        


End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii >= 48 And KeyAscii <= 57 Then

        KeyAscii = 0

    Else

        KeyAscii = KeyAscii

    End If
    
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    
    If KeyAscii >= 48 And KeyAscii <= 57 Then

        KeyAscii = KeyAscii

    Else

        KeyAscii = 0

    End If
    
End Sub
