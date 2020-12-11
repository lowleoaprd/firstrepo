Attribute VB_Name = "horario"
Sub hora()

 Dim hora As String
 Dim start As Long, fim As Long

'##### Indicadores #####
   'rs.MoveFirst
   'Set wb2 = Workbooks.Open(rs.Fields("FileName").Value, Local:=True)

'HORARIOS
   start = InStr(wb2.Name, "0-")
   fim = InStr(wb2.Name, ".csv")
   hora = Mid(wb2.Name, start + 2, fim - inicio - 1)
   
   hora = Replace(hora, ".csv", "")
   hora = Replace(hora, "-", ":")
   CS.Sheets("CS_Ind").Range("A5").Value = hora
'------
   'wb2.Sheets("PLANILHA 7-12-2020-16-59").Columns("A:W").Copy
   'CS.Sheets("Indicadores").Range("A1").PasteSpecial
   'CS.Sheets("Indicadores").Columns("A:W").AutoFit
   'Application.CutCopyMode = False
   'wb2.Close
   
End Sub

