VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Conexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private base As Database
Private registro As Recordset
Public Function Conector(NombreTabla As String)

    Set base = OpenDatabase("C:\Documents and Settings\asdakj\Escritorio\ABMCF\database\DataBase.mdb")
    
    Set registro = base.OpenRecordset(NombreTabla, dbOpenTable)

End Function
Public Function Chekeo(User As String, Pass As String) As Boolean

    registro.Index = "IndiceUser"
    registro.Seek "=", User
    
    If registro.NoMatch = False Then
        If registro.Fields("Contraseña") = Pass Then
            Chekeo = True
        
        End If
    End If


End Function
