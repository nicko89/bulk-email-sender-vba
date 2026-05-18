Attribute VB_Name = "Módulo1"
Sub ValidarCorreos()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$"
    re.IgnoreCase = True
    re.Global = False

    Dim ultimaFila As Long
    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim fila As Long
    Dim celda As Range
    Dim partes() As String
    Dim correo As Variant
    Dim resultado As String
    Dim columnas As Variant
    columnas = Array("C", "D", "E", "F", "G") ' Nuevas columnas

    For fila = 2 To ultimaFila
        resultado = ""

        For Each col In columnas
            Set celda = ws.Range(col & fila)

            If Not IsEmpty(celda.Value) Then
                partes = Split(celda.Value, ";")
                For Each correo In partes
                    correo = Trim(correo)
                    If re.Test(correo) Then
                        If resultado = "" Then
                            resultado = correo
                        Else
                            resultado = resultado & ";" & correo
                        End If
                    End If
                Next correo
            End If
        Next col

        If resultado = "" Then
            ws.Range("H" & fila).Value = "Sin correos"
        Else
            ws.Range("H" & fila).Value = resultado
        End If
    Next fila

    MsgBox "Correos validados correctamente. Revisa la columna H.", vbInformation
End Sub
