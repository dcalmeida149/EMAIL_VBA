Sub gerapdf()
    With Planilha1
        .Activate
        Dim mydate
        Dim plan
        mydate = Date
        mydate = VBA.Format(mydate, "dd-mm-yyyy")
        plan = Sheets(1).Name
        
        Caminho_Arquivo = ThisWorkbook.Path & "\Emails_enviados\Controle_semanal_" & plan & "-" & mydate & ".pdf"
        
        ActiveSheet.Range("a1:k80").ExportAsFixedFormat Type:=xlTypePDF, Filename:=Caminho_Arquivo 'Exporta para pdf no caminho indicado. Lembrar de configurar Area de impressão
        
    End With
End Sub
Sub enviarEmail()
    Call gerapdf
    
    Dim Outlook As Object, novoEmail As Object
    Dim mydate
        mydate = Date
    Dim plan
        plan = Sheets(1).Name
    Dim Caminho_Arquivo
        mydate = VBA.Format(mydate, "dd-mm-yyyy")
        Caminho_Arquivo = ThisWorkbook.Path & "\Emails_enviados\Controle_semanal_" & plan & "-" & mydate & ".pdf"
               
        Set Outlook = CreateObject("Outlook.application")
        Set novoEmail = Outlook.createitem(0)
        
    With novoEmail
        .to = ActiveSheet.[m3].Value
        .cc = ActiveSheet.[m4].Value
        .Subject = "KPI - Métrica semanal, controle de custos " & mydate & " - " & plan
        .htmlbody = ActiveSheet.[m5].Value
        .Attachments.Add Caminho_Arquivo
        .display
    End With
End Sub

