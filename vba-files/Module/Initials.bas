Attribute VB_Name = "Initials"
public report As Object
Option Explicit

Public Sub directoryValidate()

  Dim fso, filesDirectory, file, data, FileToRead As Object
  Dim item As Variant
  Set data = CreateObject("Scripting.Dictionary")
  Set fso = CreateObject("Scripting.FileSystemObject")

  Set report = ThisWorkbook

  If fso.folderExists(ThisWorkbook.Sheets(1).Range("E6").Value) Then
    Set filesDirectory = fso.getFolder(ThisWorkbook.Sheets(1).Range("E6").Value)
    For Each file In filesDirectory.files
      Set FileToRead = Workbooks.Open(ThisWorkbook.Sheets(1).Range("E6").Value & Application.PathSeparator & file.Name, ReadOnly:=True)
      For Each item In FileToread.Worksheets
        Select Case item.Name
         Case "EMO"
          Call emo
         Case "AUDIO"
          Call audio
         Case "OPTO"
          Call opto
         Case "VISIO"
          Call visio
         Case "ESPIRO"
          Call espiro
         Case "OSTEO"
          Call osteo
         Case "COMPLEMENTARIOS","COMPLEMENTARIO"
          Call complementarios
         Case "TEST DE INSOMNIO"
          Call test_insomnio
         Case "VALORACION RESPIRATORIA X FISIO"
          Call valoracion_respiratoria
         Case "PSICOTECNICA","PSICOLOGIA"
          Call psicotecnica
         Case "PSICOSENSOMETRICA","PSICOMOTRIZ"
          Call psicosensometrica
         Case "LABORATORIOS","LABORATORIO"
          Call laboratorios
         Case "TEST DE FRAMINGHAM"
          Call test_framingham
        End Select
      Next item
      FileToRead.close
    Next file
  End If

End Sub
