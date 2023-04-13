Attribute VB_Name = "Initials"
public report As Object
Option Explicit

Public Sub directoryValidate()

  Dim fso As Object, filesDirectory As Object, file As Object, data As Object, FileToRead As Object, report As Object
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

  Set report = ThisWorkbook

  If fso.folderExists(ThisWorkbook.Sheets(1).Range("E6").Value) Then
    Set filesDirectory = fso.getFolder(ThisWorkbook.Sheets(1).Range("E6").Value)
    For Each file In filesDirectory.files
      Set FileToRead = Workbooks.Open(ThisWorkbook.Sheets(1).Range("E6").Value & Application.PathSeparator & file.Name, ReadOnly:=True)
      For Each item In FileToread.Worksheets
        Debug.Print item.Name
        Select Case item.Name
         Case "EMO"
          If CLngPtr(report.Worksheets("EMO" & exts).Tab.Color) = color_db Then: Call emo_db("EMO" & exts)
          If CLngPtr(report.Worksheets("EMO" & exts).Tab.Color) = color_mt Then: Call emo_mt("EMO" & exts)
         Case "AUDIO"
          If CLngPtr(report.Worksheets("AUDIO" & exts).Tab.Color) = color_db Then: Call audio_db("AUDIO" & exts)
          If CLngPtr(report.Worksheets("AUDIO" & exts).Tab.Color) = color_mt Then: Call audio_mt("AUDIO" & exts)
         Case "OPTO"
          If CLngPtr(report.Worksheets("OPTO" & exts).Tab.Color) = color_db Then: Call opto_db("OPTO" & exts)
          If CLngPtr(report.Worksheets("OPTO" & exts).Tab.Color) = color_mt Then: Call opto_mt("OPTO" & exts)
         Case "VISIO"
          If CLngPtr(report.Worksheets("VISIO" & exts).Tab.Color) = color_db Then: Call visio_db("VISIO" & exts)
          If CLngPtr(report.Worksheets("VISIO" & exts).Tab.Color) = color_mt Then: Call visio_mt("VISIO" & exts)
         Case "ESPIRO"
          If CLngPtr(report.Worksheets("ESPIRO" & exts).Tab.Color) = color_db Then: Call espiro_db("ESPIRO" & exts)
          If CLngPtr(report.Worksheets("ESPIRO" & exts).Tab.Color) = color_mt Then: Call espiro_mt("ESPIRO" & exts)
         Case "OSTEO"
          If CLngPtr(report.Worksheets("OSTEO" & exts).Tab.Color) = color_db Then: Call osteo_db("OSTEO" & exts)
          If CLngPtr(report.Worksheets("OSTEO" & exts).Tab.Color) = color_mt Then: Call osteo_mt("OSTEO" & exts)
         Case "COMPLEMENTARIOS","COMPLEMENTARIO"
          If CLngPtr(report.Worksheets("COMPLEMENTARIOS" & exts).Tab.Color) = color_db Then: Call complementarios_db("COMPLEMENTARIOS" & exts)
          If CLngPtr(report.Worksheets("COMPLEMENTARIOS" & exts).Tab.Color) = color_mt Then: Call complementarios_mt("COMPLEMENTARIOS" & exts)
         Case "TEST DE INSOMNIO"
          If CLngPtr(report.Worksheets("TEST DE INSOMNIO" & exts).Tab.Color) = color_db Then: Call test_insomnio_db("TEST DE INSOMNIO" & exts)
          If CLngPtr(report.Worksheets("TEST DE INSOMNIO" & exts).Tab.Color) = color_mt Then: Call test_insomnio_mt("TEST DE INSOMNIO" & exts)
         Case "VALORACION RESPIRATORIA X FISIO"
          If CLngPtr(report.Worksheets("VALORACION RESPIRATORIA X FISIO").Tab.Color) = color_db Then: Call valoracion_respiratoria_db("VALORACION RESPIRATORIA X FISIO")
          If CLngPtr(report.Worksheets("VALORACION RESPIRATORIA X FISIO").Tab.Color) = color_mt Then: Call valoracion_respiratoria_mt("VALORACION RESPIRATORIA X FISIO")
         Case "PSICOTECNICA","PSICOLOGIA"
          If CLngPtr(report.Worksheets("PSICOTECNICA" & exts).Tab.Color) = color_db Then: Call psicotecnica_db("PSICOTECNICA" & exts)
          If CLngPtr(report.Worksheets("PSICOTECNICA" & exts).Tab.Color) = color_mt Then: Call psicotecnica_mt("PSICOTECNICA" & exts)
         Case "PSICOSENSOMETRICA","PSICOMOTRIZ"
          If CLngPtr(report.Worksheets("PSICOMOTRIZ" & exts).Tab.Color) = color_db Then: Call psicosensometrica_db("PSICOMOTRIZ" & exts)
          If CLngPtr(report.Worksheets("PSICOMOTRIZ" & exts).Tab.Color) = color_mt Then: Call psicosensometrica_mt("PSICOMOTRIZ" & exts)
         Case "LABORATORIOS","LABORATORIO","LAB","LABS"
          If CLngPtr(report.Worksheets("LABORATORIOS" & exts).Tab.Color) = color_db Then: Call laboratorios_db("LABORATORIOS" & exts)
          If CLngPtr(report.Worksheets("LABORATORIOS" & exts).Tab.Color) = color_mt Then: Call laboratorios_mt("LABORATORIOS" & exts)
         Case "TEST DE FRAMINGHAM"
          If CLngPtr(report.Worksheets("TEST DE FRAMINGHAM" & exts).Tab.Color) = color_db Then: Call test_framingham_db("TEST DE FRAMINGHAM" & exts)
          If CLngPtr(report.Worksheets("TEST DE FRAMINGHAM" & exts).Tab.Color) = color_mt Then: Call test_framingham_mt("TEST DE FRAMINGHAM" & exts)
        End Select
      Next item
      FileToRead.Close
    Next file
  End If

End Sub
