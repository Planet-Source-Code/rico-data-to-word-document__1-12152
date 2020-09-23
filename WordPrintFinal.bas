Attribute VB_Name = "WordPrint"
Public Sub Printen(Name, Address, City, Telephone)
    Dim Range As Word.Range
    Dim Myw As Object
    
    'if Word isnt open VB will raise an error, we dont need that :)
    On Error Resume Next
    
    'check if Word is open
    Set Myw = GetObject(, "Word.application")
    
    'if Word is not open
    If Myw Is Nothing Then
        Set Myw = GetObject("", "Word.Application")
        Myw.Visible = True
    End If
    
    'pass data to the bookmarks in the template I created in Word and print it
    If Not Myw Is Nothing Then
        'this is the location of the template I created, fill in your location
        Myw.WindowState = 2
        Myw.Documents.Add "c:/Word.dot"
        'I work with bookmarks to be able to get the data in the right spot
        Set Range = Myw.ActiveDocument.Bookmarks("bkName").Range
        Range.InsertAfter Name
        Set Range = Myw.ActiveDocument.Bookmarks("bkAddr").Range
        Range.InsertAfter Address
        Set Range = Myw.ActiveDocument.Bookmarks("bkCity").Range
        Range.InsertAfter City
        Set Range = Myw.ActiveDocument.Bookmarks("bkTele").Range
        Range.InsertAfter Telephone
        Myw.ActiveDocument.PrintOut
        'I dont save the document you can if you want to, just remove DoNot
        Myw.ActiveDocument.Close wdDoNotSaveChanges
        'remove word from memory
        Set Myw = Nothing
    End If

    End Sub


