# Create folder Refrence :-Microsoft Scripting Runtime

    Sub create_folder()
    
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    mPath = ThisWorkbook.Path
    If Not fso.FolderExists(mPath & "\Raw_data") Then
        raw_path = mPath & "\Raw_data"
        fso.CreateFolder raw_path
    End If
    
    End Sub
    
   

