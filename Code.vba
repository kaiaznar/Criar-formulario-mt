'when a (----) appear means that must be separeted bly modules in vba excel'
'---------------------'
Sub Button1_Click()
    
    Principal.call

End Sub
'---------------------'
Public Function Localizar_Dir()

    Dim ObjShell, ObjFolder, Way, SecuriteSlash
    
    Set ObjShell = CreateObject("Shell.Application")
    Set ObjFolder = ObjShell.browseforfolder(&H0&, "Procurar por um Diretório", &H1&)
    
    On Error Resume Next
    
    Way = ObjFolder.ParentFolder.parsename(ObjFolder.Title).Path & ""
    
    If ObjFolder.Title = "Bureau" Then
        
        Way = "C:WindowsBureau"
    
    End If
    If ObjFolder.Title = "" Then
        
        Way = ""
    
    End If
    
    SecuriteSlash = InStr(ObjFolder.Title, ":")
    If SecuriteSlash > 0 Then
        
        Way = Mid(ObjFolder.Title, SecuriteSlash - 1, 2) & ""
    
    End If
    
    Localizar_Dir = Way
        
End Function
'---------------------'
Public Function SearchFileORFolder(action As Byte)

    Dim NumEscolha As Integer
    Dim Path As String
    
    If action = 0 Then
        Application.FileDialog(msoFileDialogFolderPicker).AllowMultiSelect = False
        
        NumEscolha = Application.FileDialog(msoFileDialogFolderPicker).Show
        
        If NumEscolha <> 0 Then
        
            Path = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
            SearchFileORFolder = Path
        
        End If
    
    Else
        
        Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
        
        NumEscolha = Application.FileDialog(msoFileDialogOpen).Show
        
        If NumEscolha <> 0 Then
        
            Path = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
            
            Workbooks.Open (Path)
            
            SearchFileORFolder = Path
        
        End If
    
    End If
    
    

End Function
'---------------------'
Public Function ChangeForm(Nomes, SPasta)
    
    With ActiveWorkbook
        
        Range("C3") = "ST-DC-103.042-087-" & Right(Nomes, 13) & "-R00"
        Range("H5") = Nomes
        Nomes = Range("C3")
        .SaveAs SPasta & "\" & Nomes & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        
    End With
    

End Function
'---------------------'
Option Explicit
Sub Formularios()
    
    Dim PastaF, Form As String
    Dim ActionSearch As Byte
    
    MsgBox "Selecione A PASTA!", vbInformation, "Pasta Principal"
    ActionSearch = 0
    PastaF = SearchFileORFolder(ActionSearch)
    If PastaF = Empty Or PastaF = "" Then
        Exit Sub
    End If
    
    MsgBox "Selecione Formulário!", vbInformation, "FORMULÁRIOS"
    ActionSearch = 1
    Form = SearchFileORFolder(ActionSearch)
    If Form = Empty Or Form = "" Then
        Exit Sub
    End If
    
    With ActiveWorkbook
        
        ListarSubPastas PastaF
        .Close
    
    End With
    

End Sub

Sub ListarSubPastas(SourceFolderName)
    
    Dim FSO As Scripting.FileSystemObject
    Dim SourceFolder As Scripting.Folder
    Dim SubPasta As Scripting.Folder
    Dim Arquivo As Scripting.File
    Dim Var As Byte
    Dim PastaLSP, FontePastaLSP As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set SourceFolder = FSO.GetFolder(SourceFolderName)
    
    For Each SubPasta In SourceFolder.SubFolders
        
        PastaLSP = SubPasta.Name
        FontePastaLSP = SubPasta.ParentFolder
        Var = ChangeForm(PastaLSP, FontePastaLSP)
    
    Next SubPasta
    
    Set SubPasta = Nothing
    Set FSO = Nothing

End Sub

