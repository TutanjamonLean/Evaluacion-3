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
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   855
      Left            =   960
      TabIndex        =   33
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000018&
      Caption         =   "Productos"
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
      Left            =   10920
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   10095
      Begin VB.TextBox Text6 
         Height          =   495
         Index           =   1
         Left            =   1920
         TabIndex        =   32
         Top             =   3720
         Width           =   2415
      End
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   1920
         TabIndex        =   30
         Top             =   4560
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Index           =   0
         Left            =   1920
         TabIndex        =   29
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   3960
         TabIndex        =   28
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   3840
         TabIndex        =   27
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Mostrar"
         Height          =   615
         Left            =   6960
         TabIndex        =   21
         Top             =   3120
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Eliminar"
         Height          =   615
         Left            =   6960
         TabIndex        =   20
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Modificar"
         Height          =   615
         Left            =   6960
         TabIndex        =   19
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Agregar"
         Height          =   615
         Left            =   6960
         TabIndex        =   18
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000001&
         Caption         =   "Stock:"
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
         Left            =   720
         TabIndex        =   26
         Top             =   4440
         Width           =   3855
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000001&
         Caption         =   "Venta:"
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
         Left            =   720
         TabIndex        =   25
         Top             =   3600
         Width           =   3855
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000001&
         Caption         =   "Costo:"
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
         Left            =   720
         TabIndex        =   24
         Top             =   2760
         Width           =   3855
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000001&
         Caption         =   "Nombre del Producto:"
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
         Left            =   720
         TabIndex        =   23
         Top             =   1920
         Width           =   5775
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000001&
         Caption         =   "Codigo del Producto:"
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
         Left            =   720
         TabIndex        =   22
         Top             =   960
         Width           =   5775
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         Caption         =   "Datos de Productos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   600
         TabIndex        =   17
         Top             =   480
         Width           =   6015
      End
   End
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
      Left            =   3360
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   10095
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   15
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   2520
         TabIndex        =   14
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Left            =   240
      List            =   "Clientes(Zueba_Ferrante).frx":0002
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
   Begin MSFlexGridLib.MSFlexGrid Grilla2 
      Height          =   3495
      Left            =   9000
      TabIndex        =   31
      Top             =   6360
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6165
      _Version        =   393216
      Appearance      =   0
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ConductorCli As New CLientes
Dim ConductorPro As New Productos
Dim nom, ape, dire, CUIT As String
Dim Codpro, nomPro, costo, sell, stock As String
Dim nuevoNom, nuevoApe, nuevoDire, Cuitmod As String
Private Sub Command1_Click()

    nom = Text1(0).Text
    ape = Text1(1).Text
    dire = Text2.Text
    CUIT = Text3.Text
    
    
    If nom = "" Or ape = "" Or dire = "" Or CUIT = "" Then
        
        MsgBox "Todos los campos son obligatorios", vbCritical, "Error"
        
        

    ElseIf ConductorCli.Añadir(CStr(nom), CStr(ape), CStr(dire), CStr(CUIT)) = True Then
        
        MsgBox "añadido al sistema", vbInformation, "Exito"
        
        
        With Grilla
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = "Cliente"
                .TextMatrix(.Rows - 1, 1) = nom
                .TextMatrix(.Rows - 1, 2) = ape
                .TextMatrix(.Rows - 1, 3) = dire
                .TextMatrix(.Rows - 1, 4) = CUIT
            End With
    
    Else
    
        
        MsgBox "No se puede agregar el mismo", vbCritical, "Error"
        
              
    End If
    
End Sub
Private Sub Command2_Click()
    
    nuevoNom = Text1(0).Text
    nuevoApe = Text1(1).Text
    nuevoDire = Text2.Text
    Cuitmod = Text3.Text
    

    If nuevoNom = "" Or nuevoApe = "" Or nuevoDire = "" Or Cuitmod = "" Then
        
        MsgBox "Todos los campos son obligatorios", vbCritical, "Error"
    

    ElseIf ConductorCli.ModificarCampo(CStr(nuevoNom), CStr(nuevoApe), CStr(nuevoDire), CStr(Cuitmod)) = True Then
    
        MsgBox "Modificacion"
        
        ActualizarGrilla CStr(Cuitmod), CStr(nuevoNom), CStr(nuevoApe), CStr(nuevoDire)
        
    Else
    
        MsgBox "No se puede pa"
        
    End If
    


End Sub
Private Sub ActualizarGrilla(cuitGrilla As String, nomGrilla As String, apeGrilla As String, grillaDire As String)
Dim fila As Integer
Dim search As Boolean


    search = False

    For fila = 1 To Grilla.Rows - 1
        If Grilla.TextMatrix(fila, 4) = cuitGrilla Then
            Grilla.TextMatrix(fila, 1) = nomGrilla
            Grilla.TextMatrix(fila, 2) = apeGrilla
            Grilla.TextMatrix(fila, 3) = grillaDire
            search = True
            Exit For
        End If
    Next fila
            
    
    
    
End Sub
Private Sub Command3_Click()
Dim resultado As Boolean
Dim cuitDel As String
Dim A, B, col As Integer


    cuitDel = Text3.Text

    If cuitDel = "" Then
    
        MsgBox "Campo obligatorio", vbCritical, "Error"
        
    End If


    If MsgBox("¿Seguro de eliminar este campo?", vbYesNo + vbQuestion, "Confirmar") = vbNo Then
        Exit Sub
    End If
    
    
    resultado = ConductorCli.BorrarDato(cuitDel)
        
        
    If resultado = True Then
        
        MsgBox "Campo eliminado", vbInformation, "Exito"
        
        For A = 1 To Grilla.Rows - 1
                Grilla.Row = A
                
                If cuitDel = Text3.Text Then
                    For B = A To Grilla.Rows - 2
                        For col = 0 To Grilla.Cols - 1
                            Grilla.TextMatrix(B, col) = Grilla.TextMatrix(B + 1, col)
                        Next col
                    Next B
                    
                Grilla.Rows = Grilla.Rows - 1
                  
                Exit For
                End If
        Next A
    Else
    
        MsgBox "No se encontro CUIt", vbCritical, "Error"
          
    End If

End Sub
Private Sub Command4_Click()
Dim cont As Integer
cont = 1
    
        Do Until ConductorCli.registro.EOF
            With Grilla
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
                
                .Rows = .Rows + 1
                
                .TextMatrix(cont, 0) = "Cliente"
                .TextMatrix(cont, 1) = ConductorCli.registro.Fields("Nombre")
                .TextMatrix(cont, 2) = ConductorCli.registro.Fields("Apellido")
                .TextMatrix(cont, 3) = ConductorCli.registro.Fields("Direccion")
                .TextMatrix(cont, 4) = ConductorCli.registro.Fields("Cuit")
            End With
            
            cont = cont + 1

        Loop


End Sub

Private Sub Command5_Click()
    
    Codpro = Text4.Text
    nomPro = Text5.Text
    costo = Text6(0).Text
    sell = Text6(1).Text
    stock = Text7.Text
    
    If Codpro = "" Or nomPro = "" Or costo = "" Or sell = "" Or stock = "" Then
        
        MsgBox "Todos los campos son obligatorios", vbCritical, "Error"
        
        Exit Sub
    End If

    If ConductorPro.AñadirPro(CStr(Codpro), CStr(nomPro), CStr(costo), CStr(sell), CStr(stock)) = True Then
        
        MsgBox "añadido al sistema", vbInformation, "Exito"
        
        
        With Grilla2
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = "Cliente"
                .TextMatrix(.Rows - 1, 1) = Codpro
                .TextMatrix(.Rows - 1, 2) = nomPro
                .TextMatrix(.Rows - 1, 3) = costo
                .TextMatrix(.Rows - 1, 4) = sell
                .TextMatrix(.Rows - 1, 5) = stock
            End With
    
    Else
    
        
        MsgBox "No se puede agregar el mismo", vbCritical, "Error"
        
              
    End If
    
End Sub

Private Sub Command9_Click()
    Form3.Show
    Unload Me
End Sub

Private Sub Form_Load()
    
    ConductorCli.Conector2 ("Cliente")
    ConductorPro.Conector2 ("Productos")
    
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
   
   With Grilla2
        
        .Rows = 1
        .Cols = 6
        .ColAlignment(1) = 0
        .ColAlignment(2) = 0
        .ColAlignment(3) = 0
        .ColAlignment(4) = 0
        .ColAlignment(5) = 0
        .TextMatrix(0, 0) = "Producto"
        .TextMatrix(0, 1) = "Codigo del Producto"
        .TextMatrix(0, 2) = "Nombre del Producto"
        .TextMatrix(0, 3) = "Costo"
        .TextMatrix(0, 4) = "Venta"
        .TextMatrix(0, 5) = "Stock"
        .ColWidth(1) = 3000
        .ColWidth(2) = 3000
        .ColWidth(3) = 1500
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
   
   End With
    

End Sub

Private Sub Grilla_Click()
    
    Grilla
    
End Sub

Private Sub List1_Click()

   If List1.ListIndex = 0 Then
        
        Frame1.Visible = True
        Grilla.Visible = True
        Frame2.Visible = False
        Grilla2.Visible = False
        
    ElseIf List1.ListIndex = 1 Then
        
        Frame2.Visible = True
        Grilla2.Visible = True
        Frame1.Visible = False
        Grilla.Visible = False
        
   End If
        


End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Then
        
        KeyAscii = KeyAscii
        
    Else
    
        KeyAscii = 0
        
    End If
    
    
    
    
    
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)

    If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
        
        KeyAscii = KeyAscii
        
    Else
    
        KeyAscii = 0
        
    End If



End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 45 Then

        KeyAscii = KeyAscii

    Else

        KeyAscii = 0

    End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    
    If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
        
        KeyAscii = KeyAscii
        
    Else
    
        KeyAscii = 0
        
    End If


End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)

    If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Then
        
        KeyAscii = KeyAscii
        
    Else
    
        KeyAscii = 0
        
    End If





End Sub
Private Sub Text6_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 36 Then

       KeyAscii = KeyAscii

    Else

        KeyAscii = 0

    End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
 If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    
       KeyAscii = KeyAscii
        
    Else
    
        KeyAscii = 0
        
    End If
End Sub
