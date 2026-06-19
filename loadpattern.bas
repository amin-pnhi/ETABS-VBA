Public EtabsObject As ETABSv1.cOAPI
Public SapModel As ETABSv1.cSapModel
Public Helper As ETABSv1.Helper
Public ETABS_Connected As Boolean
Public Sub UnlockModel()
    SapModel.SetModelIsLocked False
End Sub
Public Function ConnectToETABS() As Boolean

    On Error GoTo ErrHandler

    Set Helper = New ETABSv1.Helper
    
    Set EtabsObject = Helper.GetObject("CSI.ETABS.API.ETABSObject")
    Set SapModel = EtabsObject.SapModel
    
    ETABS_Connected = True
    ConnectToETABS = True
    Exit Function

ErrHandler:
    ETABS_Connected = False
    MsgBox "No running ETABS model found. Open ETABS and a model first."
    ConnectToETABS = False

End Function
Function SelectEDB() As String

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Select ETABS Model"
        .Filters.Clear
        .Filters.Add "ETABS Model", "*.edb"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            SelectEDB = .SelectedItems(1)
        Else
            SelectEDB = ""
        End If
    End With

End Function
Sub OpenModel_From_Selector()

    Dim ModelPath As String
    Dim ret As Long
    
    'Choose model
    ModelPath = SelectEDB()
    If ModelPath = "" Then
        MsgBox "No file selected"
        Exit Sub
    End If
    
    'Create helper
    Set Helper = New ETABSv1.Helper
    
    'FIX 2 method (direct EXE)
    Set EtabsObject = Helper.CreateObject("C:\Program Files\Computers and Structures\ETABS 22\ETABS.exe")

    'Start ETABS
    EtabsObject.ApplicationStart
    
    'Get model
    Set SapModel = EtabsObject.SapModel
    
    'Initialize
    SapModel.InitializeNewModel
    
    'Open selected file
    ret = SapModel.File.OpenFile(ModelPath)

    If ret <> 0 Then
        MsgBox "ETABS could not open the file"
        Exit Sub
    End If
    ETABS_Connected = True
    'Unlock model (important later for combos)
    UnlockModel

    MsgBox "Model opened successfully!"

End Sub
Sub Close_ETABS()

    If ETABS_Connected = False Then Exit Sub
    
    EtabsObject.ApplicationExit False
    
    Set SapModel = Nothing
    Set EtabsObject = Nothing
    
    ETABS_Connected = False

    MsgBox "ETABS closed"

End Sub

Sub Export_Load_Combinations_To_Excel()

    If ETABS_Connected = False Then
        If ConnectToETABS = False Then Exit Sub
    End If
    
    Dim ret As Long

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    ws.Cells.Clear


    Dim NumberCombos As Long
    Dim ComboNames() As String

    '------------------------------------
    '1) Get all combination names
    '------------------------------------
    ret = SapModel.RespCombo.GetNameList(NumberCombos, ComboNames)

    If NumberCombos = 0 Then
        MsgBox "No load combinations found."
        Exit Sub
    End If

    'Dictionary to store unique load cases
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long, j As Long

    '------------------------------------
    '2) First pass: find all load cases
    '------------------------------------
    For i = 0 To NumberCombos - 1
        
        Dim NumItems As Long
        Dim CaseNames() As String
        Dim ScaleFactors() As Double
        Dim CaseTypes() As eCNameType
        
        ret = SapModel.RespCombo.GetCaseList(ComboNames(i), NumItems, CaseTypes, CaseNames, ScaleFactors)

        For j = 0 To NumItems - 1
            If Not dict.exists(CaseNames(j)) Then
                dict.Add CaseNames(j), dict.Count + 2
            End If
        Next j
        
    Next i

    '------------------------------------
    '3) Write header row
    '------------------------------------
    ws.Cells(1, 1) = "Combo Name"

    Dim key As Variant
    For Each key In dict.keys
        ws.Cells(1, dict(key)) = key
    Next key

    '------------------------------------
    '4) Second pass: write coefficients
    '------------------------------------
    Dim row As Long
    row = 2

    For i = 0 To NumberCombos - 1
        
        ws.Cells(row, 1) = ComboNames(i)

        Dim NumItems2 As Long
        Dim CaseNames2() As String
        Dim ScaleFactors2() As Double
        Dim CaseTypes2() As eCNameType
        
        ret = SapModel.RespCombo.GetCaseList(ComboNames(i), NumItems2, CaseTypes2, CaseNames2, ScaleFactors2)
                

        For j = 0 To NumItems2 - 1
            ws.Cells(row, dict(CaseNames2(j))) = ScaleFactors2(j)
        Next j
        
        row = row + 1
    Next i

    MsgBox "Load combinations exported successfully!"

End Sub
Sub Import_Load_Combinations_From_Excel()

    'Make sure connected to ETABS
    If ETABS_Connected = False Then
        If ConnectToETABS = False Then Exit Sub
    End If

    UnlockModel

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")

    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim r As Long, c As Long
    Dim ComboName As String
    Dim LoadCase As String
    Dim Factor As Double
    Dim ret As Long

    'Get existing combos in model
    Dim NumCombos As Long
    Dim ExistingCombos() As String
    SapModel.RespCombo.GetNameList NumCombos, ExistingCombos

    Dim dictExisting As Object
    Set dictExisting = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 0 To NumCombos - 1
        dictExisting.Add ExistingCombos(i), True
    Next i

    'Loop Excel combos
    For r = 2 To lastRow

        ComboName = ws.Cells(r, 1).Value
        If ComboName = "" Then GoTo NextCombo

        '-----------------------------------
        'Check if combo already exists
        '-----------------------------------
        If dictExisting.exists(ComboName) Then
            
            Dim ans As VbMsgBoxResult
            ans = MsgBox("Load Combination '" & ComboName & "' already exists." & vbCrLf & _
                         "Yes = Overwrite" & vbCrLf & _
                         "No = Skip" & vbCrLf & _
                         "Cancel = Stop Import", _
                         vbYesNoCancel + vbQuestion)

            If ans = vbCancel Then Exit Sub

            If ans = vbNo Then GoTo NextCombo

            'Overwrite ? delete old combo first
            SapModel.RespCombo.Delete ComboName

        End If

        '-----------------------------------
        'Create new combo
        '-----------------------------------
        ret = SapModel.RespCombo.Add(ComboName, 0)

        '-----------------------------------
        'Add load cases and factors
        '-----------------------------------
        For c = 2 To lastCol

            LoadCase = ws.Cells(1, c).Value
            Factor = ws.Cells(r, c).Value

            If LoadCase <> "" Then
                If Factor <> 0 Then
                    ret = SapModel.RespCombo.SetCaseList(ComboName, 0, LoadCase, Factor)
                End If
            End If

        Next c

NextCombo:
    Next r

    MsgBox "Load combination import completed!"

End Sub
