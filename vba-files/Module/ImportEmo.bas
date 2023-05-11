Attribute VB_Name = "ImportEmo"
Option Explicit
Public emo_header_origin_dictionary As Scripting.Dictionary, emo_header_destiny_dictionary As Scripting.Dictionary, diagnostics_header_destiny_dictionary As Scripting.Dictionary, diagnostics_header_origin_dictionary As Scripting.Dictionary, emphasis_header_destiny_dictionary As Scripting.Dictionary, emphasis_header_origin_dictionary As Scripting.Dictionary
Dim num As Integer, data_import As Range

Public Sub emo_db(ByVal header As String)
  Dim rng_header_destiny As Range, rng_header_origin As Range, item As Variant, separateName() As String

  Set emo_header_origin_dictionary = CreateObject("Scripting.Dictionary")
  Set emo_header_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set diagnostics_header_origin_dictionary = CreateObject("Scripting.Dictionary")
  Set diagnostics_header_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set emphasis_header_origin_dictionary = CreateObject("Scripting.Dictionary")
  Set emphasis_header_destiny_dictionary = CreateObject("Scripting.Dictionary")

  '' Configuracion de la cabeceras para el libro destino ''
  Set rng_header_destiny = destiny.Worksheets(header).Range("$A1", destiny.Worksheets(header).Range("$A1").End(xlToRight))

  '' Se recorre las cabeceras de la tabla de destino para agregar los indices de las columnas en el diccionario  datos diferentes de diagnosticos y enfasis ''
  For Each item In rng_header_destiny
    If Not emo_header_destiny_dictionary.exists(header_emo(item)) Then
      On Error Resume Next
      emo_header_destiny_dictionary.Add header_emo(item), (item.Column - 1)
      DoEvents
    End If
  Next item

  '' Se recorre las cabeceras de la tabla de destino para agregar los indices de las columnas en el diccionario de los diagnosticos y enfasis ''
  For Each item In rng_header_destiny
    On Error Resume Next
    diagnostics_header_destiny_dictionary.Add header_diag(item), (item.Column - 1)
    DoEvents
  Next item

  For Each item In rng_header_destiny
    On Error Resume Next
    emphasis_header_destiny_dictionary.Add header_emphasis(item), (item.Column - 1)
    DoEvents
  Next item

  '' Configuracion de la cabeceras para el libro origen ''

  '' se realiza la separacion de la extencion DB o MT para seleccionar la hoja del libro origen ''
  separateName = VBA.Split(header, "_")

  Set rng_header_origin = origin.Worksheets(separateName(0)).Range("$A1", origin.Worksheets(separateName(0)).Range("$A1").End(xlToRight))

  '' configuracion de la cabeceras para el libro origen ''
  For Each item In rng_header_origin
    If Not emo_header_origin_dictionary.exists(header_emo(item)) Then
      On Error Resume Next
      emo_header_origin_dictionary.Add header_emo(item), (item.Column - 1)
      DoEvents
    End If
  Next item

  '' Se recorre las cabeceras de la tabla de origen para agregar los indices de las columnas en el diccionario de los diagnosticos y enfasis ''
  For Each item In rng_header_origin
    On Error Resume Next
    diagnostics_header_origin_dictionary.Add header_diag(item), (item.Column - 1)
    DoEvents
  Next item

  For Each item In rng_header_origin
    On Error Resume Next
    emphasis_header_origin_dictionary.Add header_diag(item), (item.Column - 1)
    DoEvents
  Next item

  Set data_import = origin.Worksheets(separateName(0)).Range("$A2", origin.Worksheets(separateName(0)).Range("$A2").End(xlDown))

  '' Suma de los datos de la hoja de emo de mi libro origen ''
  nCount.sumDataInformation((data_import.count * emo_header_origin_dictionary.count))

  '' Suma de los registros en la hoja de emo de mi libro origen ''
  nCount.sumRegister(data_import.count)

  '' Traspaso de informacion a al hoja de emo de mi libro destino ''
  Windows(destiny.Name).Activate

  destiny.Worksheets(header).Select
  Range("$A2").Select
  If Not IsEmpty(ActiveCell.Offset(1, 0).value) Then
    ActiveCell.End(xlDown).Select
    Selection.Offset(1, 0).Select
  End If

  On Error GoTo 0
  For Each item In data_import
    With ActiveCell
    .Offset(, emo_header_destiny_dictionary("NOMBRE CONTRATO"))  = validateKey(emo_header_origin_dictionary, "NOMBRE CONTRATO", item)
    .Offset(, emo_header_destiny_dictionary("ORIGEN"))  = validateKey(emo_header_origin_dictionary, "ORIGEN", item)
    .Offset(, emo_header_destiny_dictionary("DESTINO"))  = validateKey(emo_header_origin_dictionary, "DESTINO", item)
    .Offset(, emo_header_destiny_dictionary("CIUDAD"))  = validateKey(emo_header_origin_dictionary, "CIUDAD", item)
    .Offset(, emo_header_destiny_dictionary("INGRESO"))  = validateKey(emo_header_origin_dictionary, "INGRESO", item)
    .Offset(, emo_header_destiny_dictionary("TIPO EXAMEN"))  = validateKey(emo_header_origin_dictionary, "TIPO EXAMEN", item)
    .Offset(, emo_header_destiny_dictionary("FECHA"))  = validateKey(emo_header_origin_dictionary, "FECHA", item)
    .Offset(, emo_header_destiny_dictionary("PACIENTE"))  = validateKey(emo_header_origin_dictionary, "PACIENTE", item)
    .Offset(, emo_header_destiny_dictionary("NRO IDENTIFICACION"))  = validateKey(emo_header_origin_dictionary, "NRO IDENTIFICACION", item)
    .Offset(, emo_header_destiny_dictionary("EDAD"))  = validateKey(emo_header_origin_dictionary, "EDAD", item)
    .Offset(, emo_header_destiny_dictionary("ESTRATO"))  = validateKey(emo_header_origin_dictionary, "ESTRATO", item)
    .Offset(, emo_header_destiny_dictionary("GENERO"))  = validateKey(emo_header_origin_dictionary, "GENERO", item)
    .Offset(, emo_header_destiny_dictionary("NRO HIJOS"))  = validateKey(emo_header_origin_dictionary, "NRO HIJOS", item)
    .Offset(, emo_header_destiny_dictionary("RAZA"))  = validateKey(emo_header_origin_dictionary, "RAZA", item)
    .Offset(, emo_header_destiny_dictionary("ESTADO CIVIL"))  = validateKey(emo_header_origin_dictionary, "ESTADO CIVIL", item)
    .Offset(, emo_header_destiny_dictionary("ESCOLARIDAD"))  = validateKey(emo_header_origin_dictionary, "ESCOLARIDAD", item)
    .Offset(, emo_header_destiny_dictionary("CARGO"))  = validateKey(emo_header_origin_dictionary, "CARGO", item)
    .Offset(, emo_header_destiny_dictionary("LAB DURACION"))  = validateKey(emo_header_origin_dictionary, "LAB DURACION", item)
    .Offset(, emo_header_destiny_dictionary("ANTIGUEDAD"))  = validateKey(emo_header_origin_dictionary, "ANTIGUEDAD", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO FISICO / RUIDO"))  = validateKey(emo_header_origin_dictionary, "RIESGO FISICO / RUIDO", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO FISICO / ILUMINACION"))  = validateKey(emo_header_origin_dictionary, "RIESGO FISICO / ILUMINACION", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO FISICO / VIBRACION"))  = validateKey(emo_header_origin_dictionary, "RIESGO FISICO / VIBRACION", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO FISICO / TEMP EXTREMAS"))  = validateKey(emo_header_origin_dictionary, "RIESGO FISICO / TEMP EXTREMAS", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO FISICO / PRES ATMOSFERICA"))  = validateKey(emo_header_origin_dictionary, "RIESGO FISICO / PRES ATMOSFERICA", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO FISICO / RAD IONIZANTES"))  = validateKey(emo_header_origin_dictionary, "RIESGO FISICO / RAD IONIZANTES", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO FISICO / RAD NO IONIZANTES"))  = validateKey(emo_header_origin_dictionary, "RIESGO FISICO / RAD NO IONIZANTES", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO DE OTROS FACTORES FISICOS"))  = validateKey(emo_header_origin_dictionary, "RIESGO DE OTROS FACTORES FISICOS", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO BIOLOGICO / VIRUS"))  = validateKey(emo_header_origin_dictionary, "RIESGO BIOLOGICO / VIRUS", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO BIOLOGICO / BACTERIAS"))  = validateKey(emo_header_origin_dictionary, "RIESGO BIOLOGICO / BACTERIAS", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO BIOLOGICO / HONGOS"))  = validateKey(emo_header_origin_dictionary, "RIESGO BIOLOGICO / HONGOS", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO BIOLOGICO / RICKETSIAS"))  = validateKey(emo_header_origin_dictionary, "RIESGO BIOLOGICO / RICKETSIAS", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO BIOLOGICO / PARASITOS"))  = validateKey(emo_header_origin_dictionary, "RIESGO BIOLOGICO / PARASITOS", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO BIOLOGICO / FLUIDOS"))  = validateKey(emo_header_origin_dictionary, "RIESGO BIOLOGICO / FLUIDOS", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO BIOLOGICO / PICADURAS"))  = validateKey(emo_header_origin_dictionary, "RIESGO BIOLOGICO / PICADURAS", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO BIOLOGICO / MORDEDURAS"))  = validateKey(emo_header_origin_dictionary, "RIESGO BIOLOGICO / MORDEDURAS", item)
    .Offset(, emo_header_destiny_dictionary("OTROS RIESGOS BIOLOGICOS"))  = validateKey(emo_header_origin_dictionary, "OTROS RIESGOS BIOLOGICOS", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO_QCO/POLVOS"))  = validateKey(emo_header_origin_dictionary, "RIESGO_QCO/POLVOS", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO QUIMICO / FIBRAS"))  = validateKey(emo_header_origin_dictionary, "RIESGO QUIMICO / FIBRAS", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO QUIMICO / LIQUIDOS"))  = validateKey(emo_header_origin_dictionary, "RIESGO QUIMICO / LIQUIDOS", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO QUIMICO /GASES"))  = validateKey(emo_header_origin_dictionary, "RIESGO QUIMICO /GASES", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO QUIMICO / VAPORES"))  = validateKey(emo_header_origin_dictionary, "RIESGO QUIMICO / VAPORES", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO_QCO/HUMOS"))  = validateKey(emo_header_origin_dictionary, "RIESGO_QCO/HUMOS", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO"))  = validateKey(emo_header_origin_dictionary, "RIESGO QUIMICO /MATERIAL PARTICULADO", item)
    .Offset(, emo_header_destiny_dictionary("OTROS RIESGOS QUIMICOS"))  = validateKey(emo_header_origin_dictionary, "OTROS RIESGOS QUIMICOS", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO PSICO / GESTION ORGANIZACIONAL"))  = validateKey(emo_header_origin_dictionary, "RIESGO PSICO / GESTION ORGANIZACIONAL", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO PSICO / CARACT DEL GRUPO"))  = validateKey(emo_header_origin_dictionary, "RIESGO PSICO / CARACT DEL GRUPO", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO PSICO / INTERFACES TAREA"))  = validateKey(emo_header_origin_dictionary, "RIESGO PSICO / INTERFACES TAREA", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO PSICO / CARACT ORGANIZACION"))  = validateKey(emo_header_origin_dictionary, "RIESGO PSICO / CARACT ORGANIZACION", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO PSICO / CONDICIONES"))  = validateKey(emo_header_origin_dictionary, "RIESGO PSICO / CONDICIONES", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO PSICO / JORNADA"))  = validateKey(emo_header_origin_dictionary, "RIESGO PSICO / JORNADA", item)
    .Offset(, emo_header_destiny_dictionary("OTROS PSICO LABORAL"))  = validateKey(emo_header_origin_dictionary, "OTROS PSICO LABORAL", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO_BIOMECANICO_POSTURA"))  = validateKey(emo_header_origin_dictionary, "RIESGO_BIOMECANICO_POSTURA", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO_BIOMECANICO_ESFUERZO"))  = validateKey(emo_header_origin_dictionary, "RIESGO_BIOMECANICO_ESFUERZO", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO_BIOMECANICO_MOVREPETITIVO"))  = validateKey(emo_header_origin_dictionary, "RIESGO_BIOMECANICO_MOVREPETITIVO", item)
    .Offset(, emo_header_destiny_dictionary("RIESGO_BIOMECANICO_MANIPULACION_CARGA"))  = validateKey(emo_header_origin_dictionary, "RIESGO_BIOMECANICO_MANIPULACION_CARGA", item)
    .Offset(, emo_header_destiny_dictionary("OTROS RIESGOS BIOMECANICOS"))  = validateKey(emo_header_origin_dictionary, "OTROS RIESGOS BIOMECANICOS", item)
    .Offset(, emo_header_destiny_dictionary("CONDICIONES DE SEGURIDAD / MECANICOS"))  = validateKey(emo_header_origin_dictionary, "CONDICIONES DE SEGURIDAD / MECANICOS", item)
    .Offset(, emo_header_destiny_dictionary("CONDICIONES DE SEGURIDAD / ELECTRICOS"))  = validateKey(emo_header_origin_dictionary, "CONDICIONES DE SEGURIDAD / ELECTRICOS", item)
    .Offset(, emo_header_destiny_dictionary("CONDICIONES DE SEGURIDAD / LOCATIVO"))  = validateKey(emo_header_origin_dictionary, "CONDICIONES DE SEGURIDAD / LOCATIVO", item)
    .Offset(, emo_header_destiny_dictionary("CONDICIONES DE SEGURIDAD / TECNOLOGICO"))  = validateKey(emo_header_origin_dictionary, "CONDICIONES DE SEGURIDAD / TECNOLOGICO", item)
    .Offset(, emo_header_destiny_dictionary("CONDICIONES DE SEGURIDAD / ACC DE TRANSITO"))  = validateKey(emo_header_origin_dictionary, "CONDICIONES DE SEGURIDAD / ACC DE TRANSITO", item)
    .Offset(, emo_header_destiny_dictionary("CONDICIONES DE SEGURIDAD / PUBLICOS"))  = validateKey(emo_header_origin_dictionary, "CONDICIONES DE SEGURIDAD / PUBLICOS", item)
    .Offset(, emo_header_destiny_dictionary("CONDICIONES DE SEGURIDAD / TRABAJO EN ALTURAS"))  = validateKey(emo_header_origin_dictionary, "CONDICIONES DE SEGURIDAD / TRABAJO EN ALTURAS", item)
    .Offset(, emo_header_destiny_dictionary("CONDICIONES DE SEGURIDAD / ESPACIOS CONFINADOS"))  = validateKey(emo_header_origin_dictionary, "CONDICIONES DE SEGURIDAD / ESPACIOS CONFINADOS", item)
    .Offset(, emo_header_destiny_dictionary("CONDICIONES DE SEGURIDAD / OTROS DE SEGURIDAD"))  = validateKey(emo_header_origin_dictionary, "CONDICIONES DE SEGURIDAD / OTROS DE SEGURIDAD", item)
    .Offset(, emo_header_destiny_dictionary("FENOMENOS NATURALES / SISMO"))  = validateKey(emo_header_origin_dictionary, "FENOMENOS NATURALES / SISMO", item)
    .Offset(, emo_header_destiny_dictionary("FENOMENOS NATURALES / TERREMOTO"))  = validateKey(emo_header_origin_dictionary, "FENOMENOS NATURALES / TERREMOTO", item)
    .Offset(, emo_header_destiny_dictionary("FENOMENOS NATURALES / VENDAVAL"))  = validateKey(emo_header_origin_dictionary, "FENOMENOS NATURALES / VENDAVAL", item)
    .Offset(, emo_header_destiny_dictionary("FENOMENOS NATURALES / INUNDACION"))  = validateKey(emo_header_origin_dictionary, "FENOMENOS NATURALES / INUNDACION", item)
    .Offset(, emo_header_destiny_dictionary("FENOMENOS NATURALES / DERRUMBE"))  = validateKey(emo_header_origin_dictionary, "FENOMENOS NATURALES / DERRUMBE", item)
    .Offset(, emo_header_destiny_dictionary("FENOMENOS NATURALES / PRECIPITACIONES"))  = validateKey(emo_header_origin_dictionary, "FENOMENOS NATURALES / PRECIPITACIONES", item)
    .Offset(, emo_header_destiny_dictionary("FENOMENOS NATURALES / OTROS NATURALES"))  = validateKey(emo_header_origin_dictionary, "FENOMENOS NATURALES / OTROS NATURALES", item)
    .Offset(, emo_header_destiny_dictionary("FECHA ACCIDENTE"))  = validateKey(emo_header_origin_dictionary, "FECHA ACCIDENTE", item)
    .Offset(, emo_header_destiny_dictionary("ACCIDENTE_PASO_EN_EMPRESA"))  = validateKey(emo_header_origin_dictionary, "ACCIDENTE_PASO_EN_EMPRESA", item)
    .Offset(, emo_header_destiny_dictionary("TIPO ACCIDENTE"))  = validateKey(emo_header_origin_dictionary, "TIPO ACCIDENTE", item)
    .Offset(, emo_header_destiny_dictionary("NATURALEZA LESION"))  = validateKey(emo_header_origin_dictionary, "NATURALEZA LESION", item)
    .Offset(, emo_header_destiny_dictionary("PARTE AFECTADA"))  = validateKey(emo_header_origin_dictionary, "PARTE AFECTADA", item)
    .Offset(, emo_header_destiny_dictionary("INCAPACIDAD"))  = validateKey(emo_header_origin_dictionary, "INCAPACIDAD", item)
    .Offset(, emo_header_destiny_dictionary("SECUELAS"))  = validateKey(emo_header_origin_dictionary, "SECUELAS", item)
    .Offset(, emo_header_destiny_dictionary("NOMBRE ENFERMEDAD"))  = validateKey(emo_header_origin_dictionary, "NOMBRE ENFERMEDAD", item)
    .Offset(, emo_header_destiny_dictionary("ETAPA"))  = validateKey(emo_header_origin_dictionary, "ETAPA", item)
    .Offset(, emo_header_destiny_dictionary("OBSERVACIONES DE ENFERMEDAD"))  = validateKey(emo_header_origin_dictionary, "OBSERVACIONES DE ENFERMEDAD", item)
    .Offset(, emo_header_destiny_dictionary("ACT_ FISICA"))  = validateKey(emo_header_origin_dictionary, "ACT_ FISICA", item)
    .Offset(, emo_header_destiny_dictionary("FUMA"))  = validateKey(emo_header_origin_dictionary, "FUMA", item)
    .Offset(, emo_header_destiny_dictionary("CONSUMO DE ALCOHOL"))  = validateKey(emo_header_origin_dictionary, "CONSUMO DE ALCOHOL", item)
    .Offset(, emo_header_destiny_dictionary("PESO"))  = validateKey(emo_header_origin_dictionary, "PESO", item)
    .Offset(, emo_header_destiny_dictionary("TALLA"))  = validateKey(emo_header_origin_dictionary, "TALLA", item)
    .Offset(1, 0).Select
    End With
    DoEvents
  Next item

  '' Removemos toda la informacion de los diccionarios ''
  emo_header_origin_dictionary.RemoveAll
  emo_header_destiny_dictionary.RemoveAll
  diagnostics_header_destiny_dictionary.RemoveAll
  diagnostics_header_origin_dictionary.RemoveAll
  emphasis_header_destiny_dictionary.RemoveAll
  emphasis_header_origin_dictionary.RemoveAll

End Sub

Public Sub emo_mt(ByVal header As String)
  Debug.Print header
End Sub

Private Function header_emo(ByVal value As String) As String
  Select Case Trim(value)
   Case "NOMBRE CONTRATO"
    header_emo = "NOMBRE CONTRATO"
   Case "ORIGEN"
    header_emo = "ORIGEN"
   Case "DESTINO"
    header_emo = "DESTINO"
   Case "CIUDAD"
    header_emo = "CIUDAD"
   Case "INGRESO"
    header_emo = "INGRESO"
   Case "TIPO EXAMEN"
    header_emo = "TIPO EXAMEN"
   Case "FECHA"
    header_emo = "FECHA"
   Case "PACIENTE"
    header_emo = "PACIENTE"
   Case "NRO IDENTIFICACION", "NRO IDENFICACION"
    header_emo = "NRO IDENTIFICACION"
   Case "EDAD", "EDADREC"
    header_emo = "EDAD"
   Case "ESTRATO"
    header_emo = "ESTRATO"
   Case "GENERO"
    header_emo = "GENERO"
   Case "NRO HIJOS"
    header_emo = "NRO HIJOS"
   Case "RAZA"
    header_emo = "RAZA"
   Case "ESTADO CIVIL"
    header_emo = "ESTADO CIVIL"
   Case "ESCOLARIDAD"
    header_emo = "ESCOLARIDAD"
   Case "CARGO"
    header_emo = "CARGO"
   Case "CARGO_REC"
    header_emo = "CARGO_REC"
   Case "LAB DURACION", "ANTIGUEDAD"
    header_emo = "LAB DURACION"
   Case "RIESGO FISICO / RUIDO"
    header_emo = "RIESGO FISICO / RUIDO"
   Case "RIESGO FISICO / ILUMINACION"
    header_emo = "RIESGO FISICO / ILUMINACION"
   Case "RIESGO FISICO / VIBRACION"
    header_emo = "RIESGO FISICO / VIBRACION"
   Case "RIESGO FISICO / TEMP EXTREMAS"
    header_emo = "RIESGO FISICO / TEMP EXTREMAS"
   Case "RIESGO FISICO / PRES ATMOSFERICA"
    header_emo = "RIESGO FISICO / PRES ATMOSFERICA"
   Case "RIESGO FISICO / RAD IONIZANTES"
    header_emo = "RIESGO FISICO / RAD IONIZANTES"
   Case "RIESGO FISICO / RAD NO IONIZANTES"
    header_emo = "RIESGO FISICO / RAD NO IONIZANTES"
   Case "RIESGO DE OTROS FACTORES FISICOS"
    header_emo = "RIESGO DE OTROS FACTORES FISICOS"
   Case "RIESGO BIOLOGICO / VIRUS"
    header_emo = "RIESGO BIOLOGICO / VIRUS"
   Case "RIESGO BIOLOGICO / BACTERIAS"
    header_emo = "RIESGO BIOLOGICO / BACTERIAS"
   Case "RIESGO BIOLOGICO / HONGOS"
    header_emo = "RIESGO BIOLOGICO / HONGOS"
   Case "RIESGO BIOLOGICO / RICKETSIAS"
    header_emo = "RIESGO BIOLOGICO / RICKETSIAS"
   Case "RIESGO BIOLOGICO / PARASITOS"
    header_emo = "RIESGO BIOLOGICO / PARASITOS"
   Case "RIESGO BIOLOGICO / FLUIDOS"
    header_emo = "RIESGO BIOLOGICO / FLUIDOS"
   Case "RIESGO BIOLOGICO / PICADURAS"
    header_emo = "RIESGO BIOLOGICO / PICADURAS"
   Case "RIESGO BIOLOGICO / MORDEDURAS"
    header_emo = "RIESGO BIOLOGICO / MORDEDURAS"
   Case "OTROS RIESGOS BIOLOGICOS"
    header_emo = "OTROS RIESGOS BIOLOGICOS"
   Case "RIESGO_QCO/POLVOS"
    header_emo = "RIESGO_QCO/POLVOS"
   Case "RIESGO QUIMICO / FIBRAS"
    header_emo = "RIESGO QUIMICO / FIBRAS"
   Case "RIESGO QUIMICO / LIQUIDOS"
    header_emo = "RIESGO QUIMICO / LIQUIDOS"
   Case "RIESGO QUIMICO /GASES", "RIESGO QUIMICO /HUMOS-GASES"
    header_emo = "RIESGO QUIMICO /GASES"
   Case "RIESGO QUIMICO / VAPORES"
    header_emo = "RIESGO QUIMICO / VAPORES"
   Case "RIESGO_QCO/HUMOS"
    header_emo = "RIESGO_QCO/HUMOS"
   Case "RIESGO QUIMICO /MATERIAL PARTICULADO"
    header_emo = "RIESGO QUIMICO /MATERIAL PARTICULADO"
   Case "OTROS RIESGOS QUIMICOS"
    header_emo = "OTROS RIESGOS QUIMICOS"
   Case "RIESGO PSICO / GESTION ORGANIZACIONAL"
    header_emo = "RIESGO PSICO / GESTION ORGANIZACIONAL"
   Case "RIESGO PSICO / CARACT DEL GRUPO"
    header_emo = "RIESGO PSICO / CARACT DEL GRUPO"
   Case "RIESGO PSICO / INTERFACES TAREA"
    header_emo = "RIESGO PSICO / INTERFACES TAREA"
   Case "RIESGO PSICO / CARACT ORGANIZACION", "RIESGO PSICO / CARACT ORGANIZACI" & Chr(211) & "N"
    header_emo = "RIESGO PSICO / CARACT ORGANIZACION"
   Case "RIESGO PSICO / CONDICIONES"
    header_emo = "RIESGO PSICO / CONDICIONES"
   Case "RIESGO PSICO / JORNADA"
    header_emo = "RIESGO PSICO / JORNADA"
   Case "OTROS PSICO LABORAL"
    header_emo = "OTROS PSICO LABORAL"
   Case "RIESGO_BIOMECANICO_POSTURA"
    header_emo = "RIESGO_BIOMECANICO_POSTURA"
   Case "RIESGO_BIOMECANICO_ESFUERZO"
    header_emo = "RIESGO_BIOMECANICO_ESFUERZO"
   Case "RIESGO_BIOMECANICO_MOVREPETITIVO"
    header_emo = "RIESGO_BIOMECANICO_MOVREPETITIVO"
   Case "RIESGO_BIOMECANICO_MANIPULACION_CARGA"
    header_emo = "RIESGO_BIOMECANICO_MANIPULACION_CARGA"
   Case "OTROS RIESGOS BIOMECANICOS"
    header_emo = "OTROS RIESGOS BIOMECANICOS"
   Case "CONDICIONES DE SEGURIDAD / MECANICOS"
    header_emo = "CONDICIONES DE SEGURIDAD / MECANICOS"
   Case "CONDICIONES DE SEGURIDAD / ELECTRICOS"
    header_emo = "CONDICIONES DE SEGURIDAD / ELECTRICOS"
   Case "CONDICIONES DE SEGURIDAD / LOCATIVO"
    header_emo = "CONDICIONES DE SEGURIDAD / LOCATIVO"
   Case "CONDICIONES DE SEGURIDAD / TECNOLOGICO"
    header_emo = "CONDICIONES DE SEGURIDAD / TECNOLOGICO"
   Case "CONDICIONES DE SEGURIDAD / ACC DE TRANSITO"
    header_emo = "CONDICIONES DE SEGURIDAD / ACC DE TRANSITO"
   Case "CONDICIONES DE SEGURIDAD / PUBLICOS"
    header_emo = "CONDICIONES DE SEGURIDAD / PUBLICOS"
   Case "CONDICIONES DE SEGURIDAD / TRABAJO EN ALTURAS"
    header_emo = "CONDICIONES DE SEGURIDAD / TRABAJO EN ALTURAS"
   Case "CONDICIONES DE SEGURIDAD / ESPACIOS CONFINADOS"
    header_emo = "CONDICIONES DE SEGURIDAD / ESPACIOS CONFINADOS"
   Case "CONDICIONES DE SEGURIDAD / OTROS DE SEGURIDAD"
    header_emo = "CONDICIONES DE SEGURIDAD / OTROS DE SEGURIDAD"
   Case "FENOMENOS NATURALES / SISMO"
    header_emo = "FENOMENOS NATURALES / SISMO"
   Case "FENOMENOS NATURALES / TERREMOTO"
    header_emo = "FENOMENOS NATURALES / TERREMOTO"
   Case "FENOMENOS NATURALES / VENDAVAL"
    header_emo = "FENOMENOS NATURALES / VENDAVAL"
   Case "FENOMENOS NATURALES / INUNDACION"
    header_emo = "FENOMENOS NATURALES / INUNDACION"
   Case "FENOMENOS NATURALES / DERRUMBE"
    header_emo = "FENOMENOS NATURALES / DERRUMBE"
   Case "FENOMENOS NATURALES / PRECIPITACIONES"
    header_emo = "FENOMENOS NATURALES / PRECIPITACIONES"
   Case "FENOMENOS NATURALES / OTROS NATURALES"
    header_emo = "FENOMENOS NATURALES / OTROS NATURALES"
   Case "FECHA ACCIDENTE"
    header_emo = "FECHA ACCIDENTE"
   Case "ACCIDENTE_PASO_EN_EMPRESA"
    header_emo = "ACCIDENTE_PASO_EN_EMPRESA"
   Case "TIPO ACCIDENTE"
    header_emo = "TIPO ACCIDENTE"
   Case "NATURALEZA LESION"
    header_emo = "NATURALEZA LESION"
   Case "PARTE AFECTADA"
    header_emo = "PARTE AFECTADA"
   Case "INCAPACIDAD", "DIAS INCAPACIDAD"
    header_emo = "INCAPACIDAD"
   Case "SECUELAS"
    header_emo = "SECUELAS"
   Case "NOMBRE ENFERMEDAD"
    header_emo = "NOMBRE ENFERMEDAD"
   Case "ETAPA"
    header_emo = "ETAPA"
   Case "OBSERVACIONES DE ENFERMEDAD", "DESCRIPCION"
    header_emo = "OBSERVACIONES DE ENFERMEDAD"
   Case "ACT_ FISICA"
    header_emo = "ACT_ FISICA"
   Case "FUMA"
    header_emo = "FUMA"
   Case "CONSUMO DE ALCOHOL"
    header_emo = "CONSUMO DE ALCOHOL"
   Case "PESO"
    header_emo = "PESO"
   Case "TALLA"
    header_emo = "TALLA"
   Case "TENSION ARTERIAL"
    header_emo = "TENSION ARTERIAL"
   Case "FREC_ CARDIACA"
    header_emo = "FREC_ CARDIACA"
   Case "FREC_ RESPIRATORIA"
    header_emo = "FREC_ RESPIRATORIA"
   Case "PERIMETRO ABDOMINAL"
    header_emo = "PERIMETRO ABDOMINAL"
   Case "LATERALIDAD"
    header_emo = "LATERALIDAD"
   Case "CODIGO DIAG PPAL"
    header_emo = "CODIGO DIAG PPAL"
   Case "DIAG PPAL"
    header_emo = "DIAG PPAL"
   Case "OBS DIAGS"
    header_emo = "OBS DIAGS"
   Case "CONCEPTO DE EVALUACION"
    header_emo = "CONCEPTO DE EVALUACION"
   Case "OBSERVACIONES DEL CONCEPTO"
    header_emo = "OBSERVACIONES DEL CONCEPTO"
   Case Else
    header_emo = 0
  End Select
End Function

Private Function header_diag(ByVal value As String) As String
  Select Case Trim(Ucase(value))
   Case "CODIGO DIAG REL" & num, "CODIGO DIAG REL " & num, "CODIGO DIAG REL" & num, "CODIGO DIAG REL" & num & ","
    header_diag = "CODIGO DIAG REL" & num
   Case "DIAG REL " & num, "DIAG REL" & num
    header_diag = "DIAG REL " & num
    num = num + 1
   Case Else
    header_diag = "0"
  End Select
End Function

Private Function header_emphasis(ByVal value As String) As String
  Select Case Trim(Ucase(value))
   Case "CONCEPTO AL ENFASIS_" & num, "CONCEPTO AL ENFASIS " & num, "CONCEPTO_AL_ENFASIS_" & num
    header_emphasis = "CONCEPTO AL ENFASIS_" & num
   Case "OBSERVACIONES_AL_ENFASIS_" & num, "OBSERVACIONES AL ENFASIS " & num, "OBSERVACIONES AL ENFASIS_" & num, "OBSERVACIONES_AL_ENFASIS " & num
    header_emphasis = "OBSERVACIONES_AL_ENFASIS_" & num
    num = num + 1
   Case Else
    header_emphasis = "0"
  End Select
End Function
