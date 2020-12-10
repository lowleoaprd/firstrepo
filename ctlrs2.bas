Attribute VB_Name = "ctlrs"
Sub sendmail()

    'Variáveis Outlook
    Dim oLookApp As Outlook.Application
    Dim oLookItm As Outlook.MailItem
    Dim oLookIns As Outlook.Inspector
    
    'Variáveis Word
    Dim oWrdDoc As Word.Document
    Dim oWrdRng As Word.Range
    
    'Variável Excel
    Dim ExcRng As Range
    
    'Ativar Outlook
    Set oLookApp = GetObject(, "Outlook.Application")
    Set oLookApp = New Outlook.Application
            
            
'################PRIMEIRO EMAIL###################
    Set oLookItm = oLookApp.CreateItem(olMailItem)
    Set ExcRng = Sheets(2).Range("A2:M27")

    With oLookItm
    
        .To = "email@"
        .CC = ""
        .Subject = "Subj 1"
    
        Set oLookIns = .GetInspector
        Set oWrdDoc = oLookIns.WordEditor

        .Display
        
        ExcRng.Copy
        Set oWrdRng = oWrdDoc.Application.ActiveDocument.Content
                    
        'Colar
        oWrdRng.PasteSpecial DataType:=wdPasteMetafilePicture
   
       '.Send
    End With

'########################SEGUNDO EMAIL#######################
    
    Set oLookItm = oLookApp.CreateItem(olMailItem)
    Set ExcRng = Sheets(3).Range("A2:M27")
    
    With oLookItm
    
        .To = "email@"
        .CC = ""
        .Subject = "Subj 2"
       
        Set oLookIns = .GetInspector
        Set oWrdDoc = oLookIns.WordEditor

        .Display
        
        ExcRng.Copy
        Set oWrdRng = oWrdDoc.Application.ActiveDocument.Content
                    
        'Colar
        oWrdRng.PasteSpecial DataType:=wdPasteMetafilePicture
   
        '.Send
    End With
        
End Sub

