Attribute VB_Name = "Módulo2"
Sub EnviarFormularioClientes(modoPrueba As Boolean, preview As Boolean)
    Dim olApp As Object
    Dim olMail As Object
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim urlFormulario As String
    Dim asuntoCorreo As String
    Dim clienteCodigo As String
    Dim clienteNombre As String
    Dim correos As Variant
    Dim correo As Variant
    Dim cuerpoMensaje As String
    Dim firma As String
    Dim correosEnviados As Integer
    Dim totalCorreos As Long
    Dim listaCorreos As String

    ' Establecer hoja de trabajo
    Set ws = ThisWorkbook.Sheets("Masivo correos Actualizacion")
    
    ' Última fila con datos en columna A
    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Asunto del correo
    asuntoCorreo = "Request to Update Your Information – Decowraps."
    
    ' URL del formulario
    urlFormulario = "https://forms.office.com/r/123488273"  'put the Form or link required'
    
    ' Contar total de correos para confirmar
    totalCorreos = 0
    For i = 2 To ultimaFila
        If Not IsEmpty(ws.Cells(i, 3).Value) Then
            correos = Split(ws.Cells(i, 3).Value, ";")
            For Each correo In correos
                If Trim(correo) <> "" Then totalCorreos = totalCorreos + 1
            Next correo
        End If
    Next i

    ' Confirmación del usuario antes de enviar
    If Not preview Then
        If MsgBox("¿Estás seguro que deseas enviar " & totalCorreos & " correos?", vbYesNo + vbQuestion, "Confirmar envío") <> vbYes Then
            MsgBox "Envío cancelado.", vbExclamation
            Exit Sub
        End If
    End If

    ' Inicializar Outlook
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0

    correosEnviados = 0

    ' Recorre las filas para enviar los correos
    For i = 2 To ultimaFila
        clienteCodigo = ws.Cells(i, 1).Value ' Columna A
        clienteNombre = ws.Cells(i, 2).Value ' Columna B

        If Trim(ws.Cells(i, 8).Value) = "" Or LCase(Trim(ws.Cells(i, 8).Value)) = "Sin correos" Then GoTo siguienteCliente ' Columna H con correos
        
        correos = Split(ws.Cells(i, 8).Value, ";")
        
        ' En modo preview solo muestra un correo
        If preview And correosEnviados >= 1 Then Exit For

        listaCorreos = "" ' Reinicia la lista de correos para este cliente

        For Each correo In correos
            correo = Trim(correo)
            If correo <> "" Then
                If listaCorreos = "" Then
                    listaCorreos = correo
                Else
                    listaCorreos = listaCorreos & ";" & correo
                End If
            End If
        Next correo

        If listaCorreos <> "" Then
            Set olMail = olApp.CreateItem(0)
            olMail.To = listaCorreos
            olMail.Subject = asuntoCorreo
            olMail.CC = "inputyoureamil@ficticialemail.com" 'Replace to copy the request sent'

            cuerpoMensaje = _
                "<p style='font-family: Calibri; font-size: 11pt;'>Dear " & clienteNombre & ",</p>" & _
                "<p style='font-family: Calibri; font-size: 11pt;'>Client Code: <b>" & clienteCodigo & "</b></p>" & _
                "<p style='font-family: Calibri; font-size: 11pt;'>We hope this message finds you well.</p>" & _
                "<p style='font-family: Calibri; font-size: 11pt;'>As part of our ongoing commitment to maintaining accurate and secure client records, we kindly ask you to complete a short update form. This will help us ensure that we have the most up-to-date contact and business information for your account.</p>" & _
                "<p style='font-family: Calibri; font-size: 11pt;'>You can access the secure form by clicking the link below:<br>" & _
                "<a href='" & urlFormulario & "' target='_blank'>" & urlFormulario & "</a></p>" & _
                "<p style='font-family: Calibri; font-size: 11pt;'>The process should take only a few minutes to complete. Your time and cooperation are greatly appreciated.</p>" & _
                "<p style='font-family: Calibri; font-size: 11pt;'>If you prefer a PDF version of the form instead, please let us know and we’ll be happy to provide it.</p>" & _
                "<p style='font-family: Calibri; font-size: 11pt;'>If you are not the correct contact for this matter, we kindly ask that you forward this email to the appropriate person within your organization.</p>" & _
                "<p style='font-family: Calibri; font-size: 11pt;'>This message was sent by Decowraps as part of a verified client update process. If you have any questions or would like to confirm the authenticity of this request, please don’t hesitate to reach out to our team:</p>" & _
                "<ul style='font-family: Calibri; font-size: 11pt;'>" & _
                "<li>US Receivables: <a href='mailto:areceivables@decowraps.com'>areceivables@decowraps.com</a></li>" & _
                "<li>Customer Service: <a href='mailto:AS@decowraps.com'>AS@decowraps.com</a></li>" & _
                "<li>Phone: <b>+1 (305) 436-1415</b></li>" & _
                "</ul>" & _
                "<p style='font-family: Calibri; font-size: 11pt;'>Thank you in advance for your support and continued partnership.</p>"



            olMail.Display
            firma = olMail.HTMLBody
            olMail.HTMLBody = cuerpoMensaje & firma

            If Not preview Then olMail.Send

            correosEnviados = correosEnviados + 1

            If preview Then Exit For
        End If

siguienteCliente:
    Next i

    MsgBox IIf(preview, "Previsualización completada para 1 correo.", "Correos enviados correctamente."), vbInformation
End Sub

Sub PreviewFormulario()
    EnviarFormularioClientes True, True
End Sub

Sub EnviarFormulario()
    EnviarFormularioClientes False, False
End Sub


