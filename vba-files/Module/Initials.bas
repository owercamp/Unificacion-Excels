Attribute VB_Name = "Initials"
Public destiny As Object, origin As Object, errors As LongPtr, nCount As Counters, tbl As ListObject, newRow As ListRow
Option Explicit

Public Sub directoryValidate()

  With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
  End With

  Dim fso As Object, filesDirectory As Object, file As Object, data As Object
  Dim typeArchive() As String, exts As String
  Dim item As Variant, position as String
  Dim color_db As LongPtr, color_mt As LongPtr

  Set data = CreateObject("Scripting.Dictionary")
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set nCount = New Counters
  Set destiny = ThisWorkbook
  Set tbl = Sheets("Index").ListObjects("tbl_registros")

  typeArchive = VBA.Split(ThisWorkbook.Sheets(1).Range("E8").value, "-")

  Select Case Trim(typeArchive(0))
   Case "DB"
    color_db = 4626167
    exts = "_DB"
   Case "MT"
    color_mt = 5066944
    exts = "_MT"
  End Select


  If fso.folderExists(ThisWorkbook.Sheets(1).Range("E6").value) Then
    Set filesDirectory = fso.getFolder(ThisWorkbook.Sheets(1).Range("E6").value)
    For Each file In filesDirectory.files
      Set origin = Workbooks.Open(ThisWorkbook.Sheets(1).Range("E6").value & Application.PathSeparator & file.Name, ReadOnly:=True)
      nCount.counter = 0
      nCount.register = 0
      For Each item In origin.Worksheets
        Select Case item.Name
         Case "EMO"
          If CLngPtr(destiny.Worksheets("EMO" & exts).Tab.Color) = color_db Then
            Call emo_db("EMO" & exts)
          ElseIf CLngPtr(destiny.Worksheets("EMO" & exts).Tab.Color) = color_mt Then
            Call emo_mt("EMO" & exts)
          End If
         Case "AUDIO"
          If CLngPtr(destiny.Worksheets("AUDIO" & exts).Tab.Color) = color_db Then
            Call audio_db("AUDIO" & exts)
          ElseIf CLngPtr(destiny.Worksheets("AUDIO" & exts).Tab.Color) = color_mt Then
            Call audio_mt("AUDIO" & exts)
          End If
         Case "OPTO"
          If CLngPtr(destiny.Worksheets("OPTO" & exts).Tab.Color) = color_db Then
            Call opto_db("OPTO" & exts)
          ElseIf CLngPtr(destiny.Worksheets("OPTO" & exts).Tab.Color) = color_mt Then
            Call opto_mt("OPTO" & exts)
          End If
         Case "VISIO"
          If CLngPtr(destiny.Worksheets("VISIO" & exts).Tab.Color) = color_db Then
            Call visio_db("VISIO" & exts)
          ElseIf CLngPtr(destiny.Worksheets("VISIO" & exts).Tab.Color) = color_mt Then
            Call visio_mt("VISIO" & exts)
          End If
         Case "ESPIRO"
          If CLngPtr(destiny.Worksheets("ESPIRO" & exts).Tab.Color) = color_db Then
            Call espiro_db("ESPIRO" & exts)
          ElseIf CLngPtr(destiny.Worksheets("ESPIRO" & exts).Tab.Color) = color_mt Then
            Call espiro_mt("ESPIRO" & exts)
          End If
         Case "OSTEO"
          If CLngPtr(destiny.Worksheets("OSTEO" & exts).Tab.Color) = color_db Then
            Call osteo_db("OSTEO" & exts)
          ElseIf CLngPtr(destiny.Worksheets("OSTEO" & exts).Tab.Color) = color_mt Then
            Call osteo_mt("OSTEO" & exts)
          End If
         Case "COMPLEMENTARIOS", "COMPLEMENTARIO"
          If CLngPtr(destiny.Worksheets("COMPLEMENTARIOS" & exts).Tab.Color) = color_db Then
            Call complementarios_db("COMPLEMENTARIOS" & exts)
          ElseIf CLngPtr(destiny.Worksheets("COMPLEMENTARIOS" & exts).Tab.Color) = color_mt Then
            Call complementarios_mt("COMPLEMENTARIOS" & exts)
          End If
         Case "TEST DE INSOMNIO"
          If CLngPtr(destiny.Worksheets("TEST DE INSOMNIO" & exts).Tab.Color) = color_db Then
            Call test_insomnio_db("TEST DE INSOMNIO" & exts)
          ElseIf CLngPtr(destiny.Worksheets("TEST DE INSOMNIO" & exts).Tab.Color) = color_mt Then
            Call test_insomnio_mt("TEST DE INSOMNIO" & exts)
          End If
         Case "VALORACION RESPIRATORIA X FISIO"
          If CLngPtr(destiny.Worksheets("VALORACION RESPIRATORIA X FISIO").Tab.Color) = color_db Then
            Call valoracion_respiratoria_db("VALORACION RESPIRATORIA X FISIO")
          ElseIf CLngPtr(destiny.Worksheets("VALORACION RESPIRATORIA X FISIO").Tab.Color) = color_mt Then
            Call valoracion_respiratoria_mt("VALORACION RESPIRATORIA X FISIO")
          End If
         Case "PSICOTECNICA", "PSICOLOGIA"
          If CLngPtr(destiny.Worksheets("PSICOTECNICA" & exts).Tab.Color) = color_db Then
            Call psicotecnica_db("PSICOTECNICA" & exts)
          ElseIf CLngPtr(destiny.Worksheets("PSICOTECNICA" & exts).Tab.Color) = color_mt Then
            Call psicotecnica_mt("PSICOTECNICA" & exts)
          End If
         Case "PSICOSENSOMETRICA", "PSICOMOTRIZ"
          If CLngPtr(destiny.Worksheets("PSICOMOTRIZ" & exts).Tab.Color) = color_db Then
            Call psicosensometrica_db("PSICOMOTRIZ" & exts)
          ElseIf CLngPtr(destiny.Worksheets("PSICOMOTRIZ" & exts).Tab.Color) = color_mt Then
            Call psicosensometrica_mt("PSICOMOTRIZ" & exts)
          End If
         Case "LABORATORIOS", "LABORATORIO", "LAB", "LABS"
          If CLngPtr(destiny.Worksheets("LABORATORIOS" & exts).Tab.Color) = color_db Then
            Call laboratorios_db("LABORATORIOS" & exts)
          ElseIf CLngPtr(destiny.Worksheets("LABORATORIOS" & exts).Tab.Color) = color_mt Then
            Call laboratorios_mt("LABORATORIOS" & exts)
          End If
         Case "TEST DE FRAMINGHAM"
          If CLngPtr(destiny.Worksheets("TEST DE FRAMINGHAM" & exts).Tab.Color) = color_db Then
            Call test_framingham_db("TEST DE FRAMINGHAM" & exts)
          ElseIf CLngPtr(destiny.Worksheets("TEST DE FRAMINGHAM" & exts).Tab.Color) = color_mt Then
            Call test_framingham_mt("TEST DE FRAMINGHAM" & exts)
          End If
        End Select
      Next item

      destiny.Worksheets("Index").Select
      If Not tbl.DataBodyRange Is Nothing Then
        Set newRow = tbl.ListRows.Add
        newRow.Range(1) = tbl.ListRows.count
        newRow.Range(2) = origin.Name
        newRow.Range(3) = nCount.counter
        newRow.Range(4) = nCount.register
        newRow.Range(5) = nCount.datainformation
        newRow.Range(3).Select
        position = Selection.Address
        Call status_id(nCount.register, position)
      Else
        tbl.Range(2, 1) = tbl.ListRows.count + 1
        tbl.Range(2, 2) = origin.Name
        tbl.Range(2, 3) = nCount.counter
        tbl.Range(2, 4) = nCount.register
        tbl.Range(2, 5) = nCount.datainformation
        tbl.Range(2, 3).Select
        position = Selection.Address
        Call status_id(nCount.register, position)
      End If

      Application.ScreenUpdating = True
      Application.Wait (Now + TimeValue("0:00:10"))
      Application.ScreenUpdating = False

      origin.Close
    Next file
  End If

  With Application
    .ScreenUpdating = True
    .Calculation = xlAutomatic
    .EnableEvents = True
  End With

End Sub
