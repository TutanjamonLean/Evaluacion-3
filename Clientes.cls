VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private base As Database
Public registro As Recordset
Public Function Conector2(NombreDetabla As String)
    
   Set base = OpenDatabase(App.Path & "\database\Database 2.mdb")
   
   Set registro = base.OpenRecordset(NombreDetabla, dbOpenTable)
   
   
End Function
Public Function Añadir(Nombre As String, Apellido As String, Direccion As String, CUIT As String) As Boolean
    

    registro.Index = "IndiceCuit"
    registro.Seek "=", CUIT


    If registro.NoMatch = False Then
       'Si existe es falso
        Exit Function
    Else
    
            registro.AddNew
            registro.Fields("Nombre") = Nombre
            registro.Fields("Apellido") = Apellido
            registro.Fields("Direccion") = Direccion
            registro.Fields("Cuit") = CUIT
            registro.Update
            Añadir = True
    End If
            
    
     
End Function
Public Function ModificarCampo(newNombre As String, newApellido As String, newDireccion As String, CUIT2 As String) As Boolean

    
    registro.Index = "IndiceCuit"
    registro.Seek "=", CUIT2

     If Not registro.NoMatch Then
        If registro.Fields("Cuit") = CUIT2 Then
            registro.Edit
            registro.Fields("Nombre") = newNombre
            registro.Fields("Apellido") = newApellido
            registro.Fields("Direccion") = newDireccion
            registro.Update
            ModificarCampo = True
        End If
    End If

End Function
Public Function BorrarDato(cuitDel As String) As Boolean

    registro.Index = "IndiceCuit"
    registro.Seek "=", cuitDel

    If registro.NoMatch = False Then
        If registro.Fields("Cuit") = cuitDel Then
            registro.Delete
            registro.Update
            BorrarDato = True
        End If
    End If
    
    Exit Function


End Function




