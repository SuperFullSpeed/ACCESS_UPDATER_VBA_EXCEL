VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_RECUPERA_ESTRUCTURA_BDLOCAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Variables privadas de tipo "String"
Private CodigoVBAExcel As String
'Variables privadas de otros tipos
Private TABLAS As Object
Private innerConexionBDLocal As ADODB.Connection









Public Function PROCESAR(ByRef ConexionBDLocal As ADODB.Connection, ByRef GenerarCodigoVBAExcel As Boolean) As Object
Set innerConexionBDLocal = ConexionBDLocal

REGISTRA_NOMBRES_TABLAS
REGISTRA_CAMPOS_TABLAS
REGISTRA_LLAVES_PRIMARIAS
REGISTRA_LLAVES_FORANEAS
If GenerarCodigoVBAExcel Then
    GENERA_CODIGO_VBA_EXCEL_PARA_REPLICAR_DICCIONARIO
    GUARDA_CODIGO_VBA_EXCEL_PARA_REPLICAR_DICCIONARIO
End If

Set PROCESAR = TABLAS
End Function









Private Sub REGISTRA_NOMBRES_TABLAS()
'Variables de otros tipos
Dim rsTablas As Object

'RECONECTA BASE DE DATOS LOCAL
fRECONECTA_BD

'20 CORRESPONDE AL TIPO DE ESQUEMA PARA LAS TABLAS
Set rsTablas = innerConexionBDLocal.OpenSchema(adSchemaTables)

'RECORRE LOS RESULTADOS
Do Until rsTablas.EOF
    'IGNORA LAS TABLAS DEL SISTEMA (COMIENZAN CON "MSys")
    If left(rsTablas("TABLE_NAME").Value, 4) <> "MSys" Then
        TABLAS.Add rsTablas("TABLE_NAME").Value, CreateObject("Scripting.Dictionary")
    End If
    rsTablas.MoveNext
Loop
End Sub
Private Sub REGISTRA_CAMPOS_TABLAS()
'Variables de tipo "String"
Dim NombreCampo As String
' Variables de otros tipos
Dim NombreTabla As Variant
Dim rsCampos As Object

'PARA CADA TABLA
With TABLAS
    For Each NombreTabla In .keys
        With .Item(NombreTabla)
            'AGREGA UN DICCIONARIO PARA LOS CAMPOS
            .Add "CAMPOS", CreateObject("Scripting.Dictionary")
            
            'OBTENER LOS CAMPOS DE LA TABLA, 4 CORRESPONDE AL TIPO DE ESQUEMA PARA LAS COLUMNAS
            Set rsCampos = innerConexionBDLocal.OpenSchema(adSchemaColumns, Array(Empty, Empty, NombreTabla))
            
            'CON EL DICCIONARIO CAMPOS
            With .Item("CAMPOS")
                'RECORREMOS LOS RESULTADOS Y LOS REGISTRAMOS
                Do Until rsCampos.EOF
                    'NOMBRE DEL CAMPO
                    NombreCampo = rsCampos("COLUMN_NAME").Value
                    
                    'AGREGA EL CAMPO
                    .Add NombreCampo, CreateObject("Scripting.Dictionary")
                    
                    'PROPIEDADES DEL CAMPO
                    With .Item(NombreCampo)
                        .Add "TIPO DATO", fNOMBRE_TIPO_DATO(rsCampos("DATA_TYPE").Value)
                        .Add "LLAVE PRIMARIA", False
                        .Add "LLAVE FORANEA", False
                        .Add "TABLA ORIGINARIA LLAVE FORANEA", ""
                        .Add "CAMPO ORIGINARIO LLAVE FORANEA", ""
                        .Add "REGLA ON UPDATE", ""
                        .Add "REGLA ON DELETE", ""
                        .Add "NOMBRE RESTRICCION LLAVE FORANEA", ""
                    End With
                    
                    rsCampos.MoveNext
                Loop
            End With
        End With
    Next NombreTabla
End With
End Sub
Private Sub REGISTRA_LLAVES_PRIMARIAS()
'Variables de tipo "String"
Dim NombreTabla As String, NombreCampo As String, ConstraintName As String
'Variables de otros tipos
Dim rsIndices As Object

'RECONECTA BASE DE DATOS LOCAL
fRECONECTA_BD

'OBTENER LAS LLAVES PRIMARIAS
Set rsIndices = innerConexionBDLocal.OpenSchema(adSchemaPrimaryKeys)

'RECORREMOS LOS RESULTADOS
While Not rsIndices.EOF
    'DATOS DE LA LLAVE PRIMARIA
    NombreTabla = rsIndices("TABLE_NAME").Value
    NombreCampo = rsIndices("COLUMN_NAME").Value
    
    'IGNORA LAS TABLAS DEL SISTEMA (COMIENZAN CON "MSys")
    If left(NombreTabla, 4) <> "MSys" Then
        'REGISTRAR EL NOMBRE DEL CAMPO
        TABLAS(NombreTabla)("CAMPOS")(NombreCampo)("LLAVE PRIMARIA") = True
    End If
    
    rsIndices.MoveNext
Wend
End Sub
Private Sub REGISTRA_LLAVES_FORANEAS()
'Variables de tipo "String"
Dim NombreTablaOriginaria As String, NombreCampoOriginaria As String
Dim NombreTablaForanea As String, NombreCampoForaneo As String
Dim ReglaOnUpdate As String, ReglaOnDelete As String, NombreRestriccion As String
'Variables de otros tipos
Dim rsIndices As Object

'RECONECTA BASE DE DATOS LOCAL
fRECONECTA_BD

'OBTENER LAS LLAVES PRIMARIAS
Set rsIndices = innerConexionBDLocal.OpenSchema(adSchemaForeignKeys)

'RECORREMOS LOS RESULTADOS
While Not rsIndices.EOF
    'DATOS DE LA LLAVE PRIMARIA
    NombreTablaOriginaria = rsIndices("PK_TABLE_NAME").Value
    NombreCampoOriginaria = rsIndices("PK_COLUMN_NAME").Value
    NombreTablaForanea = rsIndices("FK_TABLE_NAME").Value
    NombreCampoForaneo = rsIndices("FK_COLUMN_NAME").Value
    ReglaOnUpdate = rsIndices("UPDATE_RULE").Value
    ReglaOnDelete = rsIndices("DELETE_RULE").Value
    NombreRestriccion = rsIndices("FK_NAME").Value
    
    'IGNORA LAS TABLAS DEL SISTEMA (COMIENZAN CON "MSys")
    If left(NombreTablaOriginaria, 4) <> "MSys" Then
        'REGISTRAR EL NOMBRE DEL CAMPO
        TABLAS(NombreTablaForanea)("CAMPOS")(NombreCampoForaneo)("LLAVE FORANEA") = True
        TABLAS(NombreTablaForanea)("CAMPOS")(NombreCampoForaneo)("TABLA ORIGINARIA LLAVE FORANEA") = NombreTablaOriginaria
        TABLAS(NombreTablaForanea)("CAMPOS")(NombreCampoForaneo)("CAMPO ORIGINARIO LLAVE FORANEA") = NombreCampoOriginaria
        TABLAS(NombreTablaForanea)("CAMPOS")(NombreCampoForaneo)("REGLA ON UPDATE") = ReglaOnUpdate
        TABLAS(NombreTablaForanea)("CAMPOS")(NombreCampoForaneo)("REGLA ON DELETE") = ReglaOnDelete
        TABLAS(NombreTablaForanea)("CAMPOS")(NombreCampoForaneo)("NOMBRE RESTRICCION LLAVE FORANEA") = NombreRestriccion
    End If
    
    rsIndices.MoveNext
Wend
End Sub
















Private Sub GENERA_CODIGO_VBA_EXCEL_PARA_REPLICAR_DICCIONARIO()
'Variables de otros tipos
Dim NombreTabla As Variant, NombreCampo As Variant, PropiedadCampo As Variant
Dim ValorPropiedad As Variant

'VARIABLES
CodigoVBAExcel = CodigoVBAExcel & "'Variables privadas de otros tipos" & vbCrLf
CodigoVBAExcel = CodigoVBAExcel & "Private TABLAS As Object" & vbCrLf
CodigoVBAExcel = CodigoVBAExcel & vbCrLf & vbCrLf & vbCrLf & vbCrLf

'RUTINA PRINCIPAL
CodigoVBAExcel = CodigoVBAExcel & "Public Function PROCESAR() As Object" & vbCrLf
CodigoVBAExcel = CodigoVBAExcel & "REGISTRA_NOMBRES_TABLAS" & vbCrLf
CodigoVBAExcel = CodigoVBAExcel & "REGISTRA_CAMPOS_TABLAS" & vbCrLf
CodigoVBAExcel = CodigoVBAExcel & "Set PROCESAR = TABLAS" & vbCrLf
CodigoVBAExcel = CodigoVBAExcel & "End Function" & vbCrLf
CodigoVBAExcel = CodigoVBAExcel & vbCrLf & vbCrLf & vbCrLf & vbCrLf

'REGISTRO DE LAS TABLAS
CodigoVBAExcel = CodigoVBAExcel & "Private Sub REGISTRA_NOMBRES_TABLAS" & vbCrLf
CodigoVBAExcel = CodigoVBAExcel & "With TABLAS" & vbCrLf
For Each NombreTabla In TABLAS.keys
    CodigoVBAExcel = CodigoVBAExcel & vbTab & ".Add """ & NombreTabla & """,CreateObject(""Scripting.Dictionary"")" & vbCrLf
    CodigoVBAExcel = CodigoVBAExcel & vbTab & ".item(""" & NombreTabla & """).Add ""CAMPOS"",CreateObject(""Scripting.Dictionary"")" & vbCrLf
Next NombreTabla
CodigoVBAExcel = CodigoVBAExcel & "End With" & vbCrLf
CodigoVBAExcel = CodigoVBAExcel & "End Sub" & vbCrLf

'REGISTRO DE LOS CAMPOS DE LAS TABLAS
CodigoVBAExcel = CodigoVBAExcel & "Private Sub REGISTRA_CAMPOS_TABLAS" & vbCrLf
CodigoVBAExcel = CodigoVBAExcel & "With TABLAS" & vbCrLf
With TABLAS
    For Each NombreTabla In .keys
        CodigoVBAExcel = CodigoVBAExcel & vbTab & "With .item(""" & NombreTabla & """)(""CAMPOS"")" & vbCrLf
        With .Item(NombreTabla)("CAMPOS")
            For Each NombreCampo In .keys
                CodigoVBAExcel = CodigoVBAExcel & vbTab & vbTab & ".Add """ & NombreCampo & """, CreateObject(""Scripting.Dictionary"")" & vbCrLf
                CodigoVBAExcel = CodigoVBAExcel & String(2, vbTab) & "With .item(""" & NombreCampo & """)" & vbCrLf
                For Each PropiedadCampo In .Item(NombreCampo).keys
                    ValorPropiedad = .Item(NombreCampo)(PropiedadCampo)
                    
                    Select Case TypeName(ValorPropiedad)
                        Case "String"
                            ValorPropiedad = """" & ValorPropiedad & """"
                        Case "Boolean"
                            ValorPropiedad = IIf(ValorPropiedad, "True", "False")
                    End Select
                    
                    CodigoVBAExcel = CodigoVBAExcel & String(3, vbTab) & ".add """ & PropiedadCampo & """, " & ValorPropiedad & vbCrLf
                Next PropiedadCampo
                CodigoVBAExcel = CodigoVBAExcel & String(2, vbTab) & "End With" & vbCrLf
            Next NombreCampo
        End With
        CodigoVBAExcel = CodigoVBAExcel & vbTab & "End With" & vbCrLf
    Next NombreTabla
End With
CodigoVBAExcel = CodigoVBAExcel & "End With" & vbCrLf
CodigoVBAExcel = CodigoVBAExcel & "End Sub" & vbCrLf


CodigoVBAExcel = CodigoVBAExcel & String(6, vbCrLf) & vbCrLf
CodigoVBAExcel = CodigoVBAExcel & "Private Sub Class_Initialize()" & vbCrLf
CodigoVBAExcel = CodigoVBAExcel & "Set TABLAS = CreateObject(""Scripting.Dictionary"")" & vbCrLf
CodigoVBAExcel = CodigoVBAExcel & "End Sub"
End Sub
Private Sub GUARDA_CODIGO_VBA_EXCEL_PARA_REPLICAR_DICCIONARIO()
'Variables de tipo "String"
Dim RutaArchivo As String
'Variables de otros tipos
Dim fso As Object, archivoTexto As Object

'INICIALIZAMOS EL OBJETO
Set fso = CreateObject("Scripting.FileSystemObject")

'RUTA EN EL QUE SE GUARDARA EL ARCHIVO
RutaArchivo = ThisWorkbook.Path & "\" & "CodigoVBAExcel.txt"

'CREAMOS EL ARCHIVO DE TEXTO
Set archivoTexto = fso.CreateTextFile(RutaArchivo, True)

'ESCRIBIMOS EN EL
archivoTexto.Write CodigoVBAExcel

'CERRAMOS EL ARCHIVO
archivoTexto.Close
End Sub

















Private Function fNOMBRE_TIPO_DATO(ByVal NumTipoDato As Variant) As String
Select Case NumTipoDato
    Case 2
        fNOMBRE_TIPO_DATO = "SMALLINT"
    Case 3
        fNOMBRE_TIPO_DATO = "INTEGER"
    Case 4
        fNOMBRE_TIPO_DATO = "REAL"
    Case 5
        fNOMBRE_TIPO_DATO = "DOUBLE"
    Case 6
        fNOMBRE_TIPO_DATO = "MONEY"
    Case 7
        fNOMBRE_TIPO_DATO = "DATETIME"
    Case 11
        fNOMBRE_TIPO_DATO = "BIT"
    Case 17
        fNOMBRE_TIPO_DATO = "BYTE"
    Case 72
        fNOMBRE_TIPO_DATO = "GUID"
    Case 128
        fNOMBRE_TIPO_DATO = "BINARY"
    Case 130
        fNOMBRE_TIPO_DATO = "VARCHAR"
    Case 131
        fNOMBRE_TIPO_DATO = "MEMO"
    Case Else
        fNOMBRE_TIPO_DATO = "DESCONOCIDO"
End Select
End Function



















Private Sub Class_Initialize()
'INICIA DICCIONARIOS
Set TABLAS = CreateObject("Scripting.Dictionary")
End Sub
