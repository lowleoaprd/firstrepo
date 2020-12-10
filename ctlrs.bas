Attribute VB_Name = "ctlrs"
Sub sendmail()

    'Declare Outlook Variables
    Dim oLookApp As Outlook.Application
    Dim oLookItm As Outlook.MailItem
    Dim oLookIns As Outlook.Inspector
    
    'Declare Word Variables
    Dim oWrdDoc As Word.Document
    Dim oWrdRng As Word.Range
    
    'Delcare Excel Variables
    Dim ExcRng As Range
    
    On Error Resume Next
    
    'Get the Active instance of Outlook if there is one
    Set oLookApp = GetObject(, "Outlook.Application")
    Set oLookApp = New Outlook.Application
            
    'Create a new email
    Set oLookItm = oLookApp.CreateItem(olMailItem)
    
          
    'Create an array to hold ranges
    Set ExcRng = Sheets(2).Range("A2:M27")

    With oLookItm
    
        'Define some basic info of our email
        .To = "uniegufigueiredo@fei.edu.br"
        .CC = ""
        .Subject = "Here are all of my Ranges"
        .Body = "Here are all the Ranges from my worksheet."
        
        'Display the email
        
        
        'Get the Active Inspector
        Set oLookIns = .GetInspector
        
        'Get the document within the inspector
        Set oWrdDoc = oLookIns.WordEditor

        .Display
        
        ExcRng.Copy
        
        'Define the range, insert a blank line, collapse the selection.
        Set oWrdRng = oWrdDoc.Application.ActiveDocument.Content
                    
        'Paste the object.
        oWrdRng.PasteSpecial DataType:=wdPasteMetafilePicture
   
        .Send
    End With
        
        
End Sub

