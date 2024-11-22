VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Productos"
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
Public Function AñadirPro(Codpro As String, NombrePro As String, costo As String, Venta As String, Sto As String) As Boolean

    

    registro.Index = "IndiceProducto"
    registro.Seek "=", Codpro


    If registro.NoMatch = False Then
       'Si existe es falso
        Exit Function
    Else
    
            registro.AddNew
            registro.Fields("CodigoProducto") = Codpro
            registro.Fields("NombreProducto") = NombrePro
            registro.Fields("Costo") = costo
            registro.Fields("Venta") = Venta
            registro.Fields("Stock") = Sto
            registro.Update
            AñadirPro = True
    End If
            
            
End Function
