'Rewrite of CreationCut to be compatible with CorelDRAW 2022

Private Declare PtrSafe Function GetEnvironmentVariableA Lib "kernel32" ( _
    ByVal lpName As String, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long

Sub CreationCut_VDR()
    
    If Application.Documents.Count = 0 Then
        MsgBox "Abra um documento, crie um objeto e selecione-o", vbExclamation, "CreationCut"
        Exit Sub
    End If
    
    If ActiveSelection.Shapes.Count <= 0 Then
        MsgBox "Selecione ao menos um objeto", vbExclamation, "CreationCut"
        Exit Sub
    End If
    
    Dim Path As String
    Path = Space(1024)
    Dim ResultSize As Long
    ResultSize = GetEnvironmentVariableA("TEMP", Path, 1024)
    Path = Left(Path, ResultSize)
    
    Dim FileName As String
    FileName = "\PcutOut.plt"
    Dim FilePath As String
    FilePath = Path & FileName
    Dim ExpOptions As New StructExportOptions
    Dim ExpFilter As ExportFilter
    Set ExpFilter = Application.ActiveDocument.ExportEx(FilePath, cdrHPGL, cdrSelection, ExpOptions)
    
    With ExpFilter
        '.PlotterOrigin = 0 ' FilterHPGLLib.pltBottomLeft
        .ScaleFactor = 100
        .PlotterUnits = 1016
        .CurveResolution = 0.0004 ' 20#  0.004 = 0.1mm
        .removeHiddenLines = False
        .AutomaticWeld = False
        .Finish
    End With
    
    Dim ShellCommand As String
    ShellCommand = "CutterRouter.exe" & " " & FilePath
    Dim ShellReturn As Double
    ShellReturn = Shell(ShellCommand, vbNormalFocus)
    
End Sub