VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_ACTUALIZADOR_BD_LOCAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'PROPIEDADES DEL CAMPO DE BASE DE DATOS
Private PropiedadCampo_LLAVE_PRIMARIA As Boolean, PropiedadCampo_LLAVE_FORANEA As Boolean
Private PropiedadCampo_TIPO As String, PropiedadCampo_TABLA_ORIGINARIA_LLAVE_FORANEA As String, PropiedadCampo_CAMPO_ORIGINARIO_LLAVE_FORANEA As String
Private PropiedadCampo_REGLA_ON_UPDATE As String, PropiedadCampo_REGLA_ON_DELETE As String, PropiedadCampo_NOMBRE_RESTRICCION_LLAVE_FORANEA As String


'Variables privadas de tipo "String"
Private QueryLocal As String, QueryLocal1 As String
'Variables privadas de otros tipos
Private NombreTabla As Variant, NombreCampo As Variant
Private BD_LOCAL_ACTUAL As Object, BD_LOCAL_ANTERIOR As Object
Private CREADOR_BD_LOCAL_ACTUAL As New cls_ESTRUCTURA_BD_LOCAL
Private CREADOR_BD_LOCAL_ANTERIOR As New cls_RECUPERA_ESTRUCTURA_BDLOCAL
Private SentenciasSQL As Object, SentenciasSQL_CreacionTablas As Object













Public Sub PROCESAR()
Set SentenciasSQL = CreateObject("Scripting.Dictionary"): Set SentenciasSQL_CreacionTablas = CreateObject("Scripting.Dictionary")
pRECUPERA_BD_LOCAL_ACTUAL
pRECUPERA_BD_LOCAL_ANTERIOR
pCOMPARA_ESTRUCTURAS_ACTUAL_VS_ANTERIOR
pORDENA_SENTENCIAS_SQL
pEJECUTA_SENTENCIAS_SQL
End Sub









Private Sub pRECUPERA_BD_LOCAL_ACTUAL()
Set BD_LOCAL_ACTUAL = CREADOR_BD_LOCAL_ACTUAL.PROCESAR
End Sub
Private Sub pRECUPERA_BD_LOCAL_ANTERIOR()
fRECONECTA_BD '---> FUNCION GLOBAL PARA RECONECTAR BD LOCAL
Set BD_LOCAL_ANTERIOR = CREADOR_BD_LOCAL_ANTERIOR.PROCESAR(CnBD, False)
End Sub














Private Sub pCOMPARA_ESTRUCTURAS_ACTUAL_VS_ANTERIOR()
'TABLAS QUE ES NECESARIO AGREGAR EN LA ESTRUCTURA DE LA BASE DE DATOS ANTERIOR
For Each NombreTabla In BD_LOCAL_ACTUAL.keys
    If BD_LOCAL_ANTERIOR.exists(NombreTabla) Then
        pTABLA_SI_EXISTE_EN_BD_LOCAL_ANTERIOR
    Else
        pTABLA_NO_EXISTE_EN_BD_LOCAL_ANTERIOR '---> CASO CUBIERTO
    End If
Next NombreTabla

'TABLAS QUE YA NO EXISTEN EN LA ESTRUCTURA DE BASE DE DATOS ACTUAL Y ES NECESARIO ELIMINARLAS
For Each NombreTabla In BD_LOCAL_ANTERIOR.keys
    If Not BD_LOCAL_ACTUAL.exists(NombreTabla) Then
        pTABLA_ELIMINADA_EN_BD_LOCAL_ACTUAL
    End If
Next NombreTabla
End Sub













Private Sub pTABLA_SI_EXISTE_EN_BD_LOCAL_ANTERIOR()
'CAMPOS NUEVOS POR AGREGAR
For Each NombreCampo In BD_LOCAL_ACTUAL(NombreTabla)("CAMPOS").keys
    If Not BD_LOCAL_ANTERIOR(NombreTabla)("CAMPOS").exists(NombreCampo) Then
        pCAMPO_NO_EXISTE_EN_BD_LOCAL_ANTERIOR
    Else
        pCAMPO_SI_EXISTE_EN_BD_LOCAL_ANTERIOR
    End If
Next NombreCampo

'CAMPOS OBSOLETOS POR ELIMINAR
For Each NombreCampo In BD_LOCAL_ANTERIOR(NombreTabla)("CAMPOS").keys
    If Not BD_LOCAL_ACTUAL(NombreTabla)("CAMPOS").exists(NombreCampo) Then
        pCAMPO_ELIMINADO_EN_BD_LOCAL_ACTUAL
    End If
Next NombreCampo
End Sub
Private Sub pCAMPO_NO_EXISTE_EN_BD_LOCAL_ANTERIOR()
pOBTENER_PROPIEDADES_CAMPO

QueryLocal = QueryLocal & "ALTER TABLE " & NombreTabla & " ADD COLUMN " & NombreCampo & " " & PropiedadCampo_TIPO & vbCrLf
pAGREGA_SENTENCIA_SQL NombreTabla, QueryLocal, New Collection
QueryLocal = ""

If PropiedadCampo_LLAVE_FORANEA Then
    QueryLocal = QueryLocal & "ALTER TABLE " & NombreTabla & vbCrLf
    QueryLocal = QueryLocal & "ADD FOREIGN KEY (" & NombreCampo & ") REFERENCES " & PropiedadCampo_TABLA_ORIGINARIA_LLAVE_FORANEA & "(" & PropiedadCampo_CAMPO_ORIGINARIO_LLAVE_FORANEA & ")" & vbCrLf & "ON UPDATE " & PropiedadCampo_REGLA_ON_UPDATE & vbCrLf & "ON DELETE " & PropiedadCampo_REGLA_ON_DELETE
    QueryLocal = QueryLocal & ";" & vbCrLf
End If
pAGREGA_SENTENCIA_SQL NombreTabla, QueryLocal, New Collection
QueryLocal = ""
End Sub
Private Sub pCAMPO_SI_EXISTE_EN_BD_LOCAL_ANTERIOR()
'Variables de tipo "Boolean"
Dim flag0 As Boolean
'Variables de tipo "String"
Dim Restriccion As String
'Variables de tipo "Collection"
Dim TablasRequeridas As New Collection

pOBTENER_PROPIEDADES_CAMPO

If Not BD_LOCAL_ANTERIOR(NombreTabla)("CAMPOS")(NombreCampo)("TIPO DATO") = PropiedadCampo_TIPO Then
    QueryLocal = QueryLocal & "ALTER TABLE " & NombreTabla & " ALTER COLUMN " & NombreCampo & " " & PropiedadCampo_TIPO
    pAGREGA_SENTENCIA_SQL NombreTabla, QueryLocal, New Collection
    QueryLocal = ""
End If

If PropiedadCampo_LLAVE_FORANEA Then
    Restriccion = PropiedadCampo_NOMBRE_RESTRICCION_LLAVE_FORANEA
    
    If BD_LOCAL_ANTERIOR(NombreTabla)("CAMPOS")(NombreCampo)("LLAVE FORANEA") Then
        If Not (BD_LOCAL_ANTERIOR(NombreTabla)("CAMPOS")(NombreCampo)("REGLA ON UPDATE") = PropiedadCampo_REGLA_ON_UPDATE _
            And BD_LOCAL_ANTERIOR(NombreTabla)("CAMPOS")(NombreCampo)("REGLA ON DELETE") = PropiedadCampo_REGLA_ON_DELETE _
            And BD_LOCAL_ANTERIOR(NombreTabla)("CAMPOS")(NombreCampo)("TABLA ORIGINARIA LLAVE FORANEA") = PropiedadCampo_TABLA_ORIGINARIA_LLAVE_FORANEA And BD_LOCAL_ANTERIOR(NombreTabla)("CAMPOS")(NombreCampo)("CAMPO ORIGINARIO LLAVE FORANEA") = PropiedadCampo_CAMPO_ORIGINARIO_LLAVE_FORANEA _
            ) Then
            QueryLocal = QueryLocal & "ALTER TABLE " & NombreTabla & " DROP CONSTRAINT " & BD_LOCAL_ANTERIOR(NombreTabla)("CAMPOS")(NombreCampo)("NOMBRE RESTRICCION LLAVE FORANEA")
            pAGREGA_SENTENCIA_SQL NombreTabla, QueryLocal, New Collection
            QueryLocal = ""
            flag0 = True
        End If
    End If
    
    If (Not BD_LOCAL_ANTERIOR(NombreTabla)("CAMPOS")(NombreCampo)("LLAVE FORANEA")) Or flag0 Then
        If Not flag0 Then
            'FORZA LA DISPONIBILIDAD DE LA RESTRICCION ELIMINANDOLA
            QueryLocal = QueryLocal & "ALTER TABLE " & NombreTabla & vbCrLf
            QueryLocal = QueryLocal & "DROP CONSTRAINT " & Restriccion & vbCrLf
            On Error Resume Next
            EJECUTA_QUERY_LOCAL QueryLocal, CnBD
            QueryLocal = ""
            On Error GoTo 0
        End If
    
        QueryLocal = "ALTER TABLE " & NombreTabla & " ADD CONSTRAINT " & Restriccion & vbCrLf
        QueryLocal = QueryLocal & "FOREIGN KEY (" & NombreCampo & ") " & vbCrLf
        QueryLocal = QueryLocal & "REFERENCES " & PropiedadCampo_TABLA_ORIGINARIA_LLAVE_FORANEA & " (" & PropiedadCampo_CAMPO_ORIGINARIO_LLAVE_FORANEA & ") " & vbCrLf
        QueryLocal = QueryLocal & "ON UPDATE " & PropiedadCampo_REGLA_ON_UPDATE & vbCrLf
        QueryLocal = QueryLocal & "ON DELETE " & PropiedadCampo_REGLA_ON_DELETE & ";" & vbCrLf
        TablasRequeridas.Add PropiedadCampo_TABLA_ORIGINARIA_LLAVE_FORANEA
        pAGREGA_SENTENCIA_SQL NombreTabla, QueryLocal, TablasRequeridas
        QueryLocal = ""
    End If
End If
End Sub
Private Sub pCAMPO_ELIMINADO_EN_BD_LOCAL_ACTUAL()
QueryLocal = QueryLocal & "ALTER TABLE " & NombreTabla & " DROP COLUMN " & NombreCampo & ";" & vbCrLf
End Sub






Private Sub pTABLA_NO_EXISTE_EN_BD_LOCAL_ANTERIOR()
'Variables de tipo "String"
Dim QueryLocalLlavePrimaria As String, QueryLocalLlaveForanea As String
'Variables de tipo "Collection"
Dim TablasRequeridas As New Collection

QueryLocal = QueryLocal & "CREATE TABLE " & NombreTabla & " (" & vbCrLf

With BD_LOCAL_ACTUAL(NombreTabla)("CAMPOS")
    QueryLocalLlavePrimaria = ", PRIMARY KEY ("
    QueryLocalLlaveForanea = ""
    
    For Each NombreCampo In .keys
        pOBTENER_PROPIEDADES_CAMPO
        QueryLocal = QueryLocal & NombreCampo & " " & PropiedadCampo_TIPO & ", "
        
        If PropiedadCampo_LLAVE_PRIMARIA Then
            QueryLocalLlavePrimaria = QueryLocalLlavePrimaria & NombreCampo & ","
        End If
        If PropiedadCampo_LLAVE_FORANEA Then
            TablasRequeridas.Add PropiedadCampo_TABLA_ORIGINARIA_LLAVE_FORANEA
            QueryLocalLlaveForanea = QueryLocalLlaveForanea & ", FOREIGN KEY (" & NombreCampo & ") REFERENCES " & PropiedadCampo_TABLA_ORIGINARIA_LLAVE_FORANEA & "(" & PropiedadCampo_CAMPO_ORIGINARIO_LLAVE_FORANEA & ")" & vbCrLf & "ON UPDATE " & PropiedadCampo_REGLA_ON_UPDATE & vbCrLf & "ON DELETE " & PropiedadCampo_REGLA_ON_DELETE
        End If
    Next NombreCampo
End With

QueryLocal = left(QueryLocal, Len(QueryLocal) - 2)
If Not QueryLocalLlavePrimaria = ", PRIMARY KEY (" Then
    QueryLocalLlavePrimaria = left(QueryLocalLlavePrimaria, Len(QueryLocalLlavePrimaria) - 1) & ")"
End If
QueryLocal = QueryLocal & QueryLocalLlavePrimaria
QueryLocal = QueryLocal & QueryLocalLlaveForanea
QueryLocal = QueryLocal & vbCrLf & ");" & vbCrLf

pAGREGA_SENTENCIA_SQL NombreTabla, QueryLocal, TablasRequeridas
pAGREGA_SENTENCIA_SQL_CREACION_TABLA NombreTabla, SentenciasSQL.Count
QueryLocal = ""
End Sub
Private Sub pOBTENER_PROPIEDADES_CAMPO()
With BD_LOCAL_ACTUAL(NombreTabla)("CAMPOS")(NombreCampo)
    PropiedadCampo_TIPO = .Item("TIPO DATO")
    PropiedadCampo_LLAVE_PRIMARIA = .Item("LLAVE PRIMARIA")
    PropiedadCampo_LLAVE_FORANEA = .Item("LLAVE FORANEA")
    PropiedadCampo_TABLA_ORIGINARIA_LLAVE_FORANEA = .Item("TABLA ORIGINARIA LLAVE FORANEA")
    PropiedadCampo_CAMPO_ORIGINARIO_LLAVE_FORANEA = .Item("CAMPO ORIGINARIO LLAVE FORANEA")
    PropiedadCampo_REGLA_ON_UPDATE = .Item("REGLA ON UPDATE")
    PropiedadCampo_REGLA_ON_DELETE = .Item("REGLA ON DELETE")
    PropiedadCampo_NOMBRE_RESTRICCION_LLAVE_FORANEA = .Item("NOMBRE RESTRICCION LLAVE FORANEA")
End With
End Sub
Private Sub pAGREGA_SENTENCIA_SQL(ByRef Tabla As Variant, ByVal Sentencia As String, ByRef TablasRequeridas As Collection, Optional NumeroSentencia As Variant)
'Variables de otros tipos
Dim Posicion As Variant

With SentenciasSQL
    If IsMissing(NumeroSentencia) Then
        Posicion = (.Count + 1)
    Else
        Posicion = NumeroSentencia
    End If
    
    .Add Posicion, CreateObject("Scripting.Dictionary")
    
    With .Item(Posicion)
        If Not .exists("TABLA") Then .Add "TABLA", Tabla
        If Not .exists("SENTENCIA SQL") Then .Add "SENTENCIA SQL", New Collection
        With .Item("SENTENCIA SQL")
            .Add Sentencia
        End With
        If Not .exists("TABLAS NECESARIAS") Then .Add "TABLAS NECESARIAS", TablasRequeridas
        If Not .exists("LLAVE") Then .Add "LLAVE", Posicion
    End With
End With
End Sub
Private Sub pAGREGA_SENTENCIA_SQL_CREACION_TABLA(ByRef Tabla As Variant, ByRef NumeroSentencia As Variant)
With SentenciasSQL_CreacionTablas
    .Add Tabla, CreateObject("Scripting.Dictionary")
    With .Item(Tabla)
        .Add "TABLA", Tabla
        .Add "NUMERO SENTENCIA", NumeroSentencia
    End With
End With
End Sub
Private Sub pORDENA_SENTENCIAS_SQL()
'Variables de otros tipos
Dim LlaveTemp As Variant, LlaveCreacionTabla As Variant, SentenciaSQL As Variant

For Each SentenciaSQL In SentenciasSQL.items
    For Each TablaNecesaria In SentenciaSQL("TABLAS NECESARIAS")
        If SentenciasSQL_CreacionTablas.exists(TablaNecesaria) Then
            LlaveTemp = SentenciaSQL("LLAVE")
            LlaveCreacionTabla = SentenciasSQL_CreacionTablas(TablaNecesaria)("NUMERO SENTENCIA")
            
            temp = LlaveTemp
            SentenciasSQL.key(LlaveTemp) = "anything"
            SentenciasSQL.key(LlaveCreacionTabla) = temp
            SentenciasSQL.key("anything") = LlaveCreacionTabla
            
            SentenciasSQL(LlaveCreacionTabla)("LLAVE") = LlaveCreacionTabla
            SentenciasSQL(LlaveTemp)("LLAVE") = LlaveTemp
        End If
    Next TablaNecesaria
Next SentenciaSQL
End Sub
Private Sub pEJECUTA_SENTENCIAS_SQL()
'Variables de otros tipos
Dim x As Variant
Dim SentenciaSQL As Variant

For x = 1 To SentenciasSQL.Count
    For Each SentenciaSQL In SentenciasSQL(x)("SENTENCIA SQL")
        EJECUTA_QUERY_LOCAL CStr(SentenciaSQL), CnBD
    Next SentenciaSQL
Next x
End Sub














Private Sub pTABLA_ELIMINADA_EN_BD_LOCAL_ACTUAL()
QueryLocal = QueryLocal & "DROP TABLE " & NombreTabla & ";" & vbCrLf
pAGREGA_SENTENCIA_SQL NombreTabla, QueryLocal, New Collection
End Sub




















Private Function EJECUTA_QUERY_LOCAL(ByRef QueryLocal As String, ByRef Conexion As ADODB.Connection, Optional ByRef rs As ADODB.Recordset, Optional BorraQuery As Boolean = True) As ADODB.Recordset
Set rs = New ADODB.Recordset

rs.CursorLocation = adUseClient
rs.Open QueryLocal, Conexion, adOpenStatic
rs.ActiveConnection = Nothing
rs.StayInSync = False
If BorraQuery Then QueryLocal = ""

Set EJECUTA_QUERY_LOCAL = rs
End Function
