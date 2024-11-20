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
Private registro As Recordset
Public Function Conector2(NombreDetabla As String)
    
   Set base = OpenDatabase(App.Path & "\database\Database 2.mdb")
   
   Set registro = base.OpenRecordset(NombreDetabla, dbOpenTable)
   
   
End Function
Public Function Añadir(Nombre As String, Apellido As String, Direccion As String, CUIT As String) As Boolean


            registro.AddNew
            registro.Fields("Nombre") = Nombre
            registro.Fields("Apellido") = Apellido
            registro.Fields("Direccion") = Direccion
            registro.Fields("Cuit") = CUIT
            registro.Update
            Añadir = True
            
            
    
     
End Function
Public Function ModificarCampo(Nombre2 As String, Apellido2 As String, Direccion2 As String, CUIT2 As String) As Boolean


    If registro.NoMatch = True Then
        registro.Edit
        Nombre2 = registro.Fields("Nombre")
        registro.Fields("Apellido") = Apellido2
        registro.Fields("Direccion") = Direccion2
        registro.Fields("Cuit") = CUIT
        registro.Update
        ModificarCampo = True
    End If
  

End Function
