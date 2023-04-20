Attribute VB_Name = "Initials"
public destiny As Object, origin As Object
Option Explicit

Public Sub directoryValidate()

  Dim fso As Object, filesDirectory As Object, file As Object, data As Object
  Dim typeArchive() As String, exts As String
  Dim item As Variant
  Dim color_db As LongPtr, color_mt As LongPtr

  Set data = CreateObject("Scripting.Dictionary")
  Set fso = CreateObject("Scripting.FileSystemObject")

  typeArchive = VBA.Split(ThisWorkbook.Sheets(1).Range("E8").value,"-")

  Select Case Trim(typeArchive(0))
   Case "DB"
    color_db = 4626167
    exts = "_DB"
   Case "MT"
    color_mt = 5066944
    exts = "_MT"
  End Select

  Set destiny = ThisWorkbook

  If fso.folderExists(ThisWorkbook.Sheets(1).Range("E6").Value) Then
    Set filesDirectory = fso.getFolder(ThisWorkbook.Sheets(1).Range("E6").Value)
    For Each file In filesDirectory.files
      Set origin = Workbooks.Open(ThisWorkbook.Sheets(1).Range("E6").Value & Application.PathSeparator & file.Name, ReadOnly:=True)
      For Each item In origin.Worksheets
        Debug.Print item.Name
        Select Case item.Name
         Case "EMO"
          If CLngPtr(destiny.Worksheets("EMO" & exts).Tab.Color) = color_db Then
            Call emo_db("EMO" & exts)
          ElseIf CLngPtr(destiny.Worksheets("EMO" & exts).Tab.Color) = color_mt Then
            Call emo_mt("EMO" & exts)
          End if
         Case "AUDIO"
          If CLngPtr(destiny.Worksheets("AUDIO" & exts).Tab.Color) = color_db Then
            Call audio_db("AUDIO" & exts)
          ElseIf CLngPtr(destiny.Worksheets("AUDIO" & exts).Tab.Color) = color_mt Then
            Call audio_mt("AUDIO" & exts)
          End if
         Case "OPTO"
          If CLngPtr(destiny.Worksheets("OPTO" & exts).Tab.Color) = color_db Then
            Call opto_db("OPTO" & exts)
          ElseIf CLngPtr(destiny.Worksheets("OPTO" & exts).Tab.Color) = color_mt Then
            Call opto_mt("OPTO" & exts)
          End if
         Case "VISIO"
          If CLngPtr(destiny.Worksheets("VISIO" & exts).Tab.Color) = color_db Then
            Call visio_db("VISIO" & exts)
          ElseIf CLngPtr(destiny.Worksheets("VISIO" & exts).Tab.Color) = color_mt Then
            Call visio_mt("VISIO" & exts)
          End if
         Case "ESPIRO"
          If CLngPtr(destiny.Worksheets("ESPIRO" & exts).Tab.Color) = color_db Then
            Call espiro_db("ESPIRO" & exts)
          ElseIf CLngPtr(destiny.Worksheets("ESPIRO" & exts).Tab.Color) = color_mt Then
            Call espiro_mt("ESPIRO" & exts)
          End if
         Case "OSTEO"
          If CLngPtr(destiny.Worksheets("OSTEO" & exts).Tab.Color) = color_db Then
            Call osteo_db("OSTEO" & exts)
          ElseIf CLngPtr(destiny.Worksheets("OSTEO" & exts).Tab.Color) = color_mt Then
            Call osteo_mt("OSTEO" & exts)
          End if
         Case "COMPLEMENTARIOS","COMPLEMENTARIO"
          If CLngPtr(destiny.Worksheets("COMPLEMENTARIOS" & exts).Tab.Color) = color_db Then
            Call complementarios_db("COMPLEMENTARIOS" & exts)
          ElseIf CLngPtr(destiny.Worksheets("COMPLEMENTARIOS" & exts).Tab.Color) = color_mt Then
            Call complementarios_mt("COMPLEMENTARIOS" & exts)
          End if
         Case "TEST DE INSOMNIO"
          If CLngPtr(destiny.Worksheets("TEST DE INSOMNIO" & exts).Tab.Color) = color_db Then
            Call test_insomnio_db("TEST DE INSOMNIO" & exts)
          ElseIf CLngPtr(destiny.Worksheets("TEST DE INSOMNIO" & exts).Tab.Color) = color_mt Then
            Call test_insomnio_mt("TEST DE INSOMNIO" & exts)
          End if
         Case "VALORACION RESPIRATORIA X FISIO"
          If CLngPtr(destiny.Worksheets("VALORACION RESPIRATORIA X FISIO").Tab.Color) = color_db Then
            Call valoracion_respiratoria_db("VALORACION RESPIRATORIA X FISIO")
          ElseIf CLngPtr(destiny.Worksheets("VALORACION RESPIRATORIA X FISIO").Tab.Color) = color_mt Then
            Call valoracion_respiratoria_mt("VALORACION RESPIRATORIA X FISIO")
          End if
         Case "PSICOTECNICA","PSICOLOGIA"
          If CLngPtr(destiny.Worksheets("PSICOTECNICA" & exts).Tab.Color) = color_db Then
            Call psicotecnica_db("PSICOTECNICA" & exts)
          ElseIf CLngPtr(destiny.Worksheets("PSICOTECNICA" & exts).Tab.Color) = color_mt Then
            Call psicotecnica_mt("PSICOTECNICA" & exts)
          End if
         Case "PSICOSENSOMETRICA","PSICOMOTRIZ"
          If CLngPtr(destiny.Worksheets("PSICOMOTRIZ" & exts).Tab.Color) = color_db Then
            Call psicosensometrica_db("PSICOMOTRIZ" & exts)
          ElseIf CLngPtr(destiny.Worksheets("PSICOMOTRIZ" & exts).Tab.Color) = color_mt Then
            Call psicosensometrica_mt("PSICOMOTRIZ" & exts)
          End if
         Case "LABORATORIOS","LABORATORIO","LAB","LABS"
          If CLngPtr(destiny.Worksheets("LABORATORIOS" & exts).Tab.Color) = color_db Then
            Call laboratorios_db("LABORATORIOS" & exts)
          ElseIf CLngPtr(destiny.Worksheets("LABORATORIOS" & exts).Tab.Color) = color_mt Then
            Call laboratorios_mt("LABORATORIOS" & exts)
          End if
         Case "TEST DE FRAMINGHAM"
          If CLngPtr(destiny.Worksheets("TEST DE FRAMINGHAM" & exts).Tab.Color) = color_db Then
            Call test_framingham_db("TEST DE FRAMINGHAM" & exts)
          ElseIf CLngPtr(destiny.Worksheets("TEST DE FRAMINGHAM" & exts).Tab.Color) = color_mt Then
            Call test_framingham_mt("TEST DE FRAMINGHAM" & exts)
          End if
        End Select
      Next item
      origin.Close
    Next file
  End If

End Sub
