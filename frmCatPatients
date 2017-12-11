Option Compare Database
Option Explicit

Private mryOwnerDataTable() As Variant
Private arycboOwnerData() As Variant

Private mryVeterinarianDataTable() As Variant
Private arycboVeterinarianData() As Variant

Private Sub Form_Load()
    'Ensure all fields are LOCKED by default
    Call frmCatPatientsLoadLock
    'Ensure Neurological Symptoms combobox is either ENABLED if checkbox is TRUE, or DISABLED if checkbox is FALSE
    Call frmCatPatientsNeuroBox
    
    'Set target queries/tables
    Dim qryOwnersTable As DAO.Recordset
    Dim qryVeterinariansTable As DAO.Recordset
    
    'RecordCount is used instead of a set integer to set array sizes (no point in using a variable if it's just an interim that gets it data from exactly the same place)
    
    'Set counters to track progress
    Dim intOwnerCounter As Integer
    Dim intVeterinarianCounter As Integer
    
    'Set Variants cbo populating
    Dim varOwnersElement As Variant
    Dim varVetElement As Variant
    
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'Open query/table
    Set qryOwnersTable = CurrentDb.OpenRecordset("SELECT * FROM qryOwnersTable ORDER BY txtSurname", dbOpenDynaset)
    
    'Set to last record to make sure maximum size is correct
    qryOwnersTable.MoveLast
    Debug.Print "qryOwnersTable.RecordCount 1: " & qryOwnersTable.RecordCount
    
    If Not qryOwnersTable.EOF Then
        'Ensure we begin on the first row
        qryOwnersTable.MoveFirst
    
        'Start OwnerCounter
        intOwnerCounter = 0
        'The size of the array should be equal to the number of rows in the table
        'Need to size the array
        'mryOwnerDataTable is a multidimensional array containing Owner/Vet unique ID and Full Name
        'arycboOwnerData is a normal array containing Full Names to populate the Owner dropdown
        'Had trouble only displaying 1 column of a mry in a cbo, so the ary is used to fill ListIndex is used on mry to retrieve data
        ReDim mryOwnerDataTable(2, qryOwnersTable.RecordCount) As Variant
        ReDim arycboOwnerData(qryOwnersTable.RecordCount) As Variant
        
        Debug.Print "qryOwnersTable.RecordCount 2: " & qryOwnersTable.RecordCount
        
        'Set first cell to be blank for dropdown
        mryOwnerDataTable(0, 0) = ("")
        mryOwnerDataTable(1, 0) = ("")
        arycboOwnerData(0) = ("")
        Debug.Print "mryOwnerDataTable ("; 0 & "): " & mryOwnerDataTable(0, 0) & ", " & mryOwnerDataTable(1, 0)
        
        'Move to next row
        intOwnerCounter = 1
        
        Do Until qryOwnersTable.EOF
            'Get query data and populate the relevant rows with it, repeat until end of query
            'Data is sorted by order of surname
            mryOwnerDataTable(0, intOwnerCounter) = qryOwnersTable.Fields("txtID")
            mryOwnerDataTable(1, intOwnerCounter) = qryOwnersTable.Fields("txtFullname")
            arycboOwnerData(intOwnerCounter) = qryOwnersTable.Fields("txtFullname")
            Debug.Print "mryOwnerDataTable ("; intOwnerCounter & "): " & mryOwnerDataTable(0, intOwnerCounter) & ", " & mryOwnerDataTable(1, intOwnerCounter)
            Debug.Print "arycboOwnerData ("; intOwnerCounter & "): " & arycboOwnerData(intOwnerCounter)
            
            intOwnerCounter = intOwnerCounter + 1
            qryOwnersTable.MoveNext
        Loop
    
    End If
    
    If IsObject(qryOwnersTable) Then Set qryOwnersTable = Nothing
    'Reset counter for cbo populating
    intOwnerCounter = 0
    'Debug.Print UBound(arycboOwnerData)
    'Blank existing values from cbo
    Me.cboOwner.RowSource = ""
    'Fill cbo with ary data
    For Each varOwnersElement In arycboOwnerData
        Me.cboOwner.AddItem Item:=arycboOwnerData(intOwnerCounter), Index:=intOwnerCounter
        intOwnerCounter = intOwnerCounter + 1
    Next
    
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'Repeat for Veterinary Tables (yes this could probably be done in parallel but it seemed safer opening one table at a time)
    'Open your table
    Set qryVeterinariansTable = CurrentDb.OpenRecordset("SELECT * FROM qryVeterinariansTable ORDER BY txtSurname", dbOpenDynaset)
    
    qryVeterinariansTable.MoveLast
    'Debug.Print "qryVeterinariansTable.RecordCount: " & qryVeterinariansTable.RecordCount
    
    If Not qryVeterinariansTable.EOF Then
        'Ensure we begin on the first row
        qryVeterinariansTable.MoveFirst
    
        'Start OwnerCounter
        intVeterinarianCounter = 0
        'The size of the array should be equal to the number of rows in the table
        'Need to size the array
        'mryOwnerDataTable is a multidimensional array containing Owner/Vet unique ID and Full Name
        'arycboOwnerData is a normal array containing Full Names to populate the Owner dropdown
        'Had trouble only displaying 1 column of a mry in a cbo, so the ary is compared and matched to mry to retrieve data
        ReDim mryVeterinarianDataTable(2, qryVeterinariansTable.RecordCount) As Variant
        ReDim arycboVeterinarianData(qryVeterinariansTable.RecordCount) As Variant
        
        'Debug.Print "qryVeterinariansTable.RecordCount: " & qryVeterinariansTable.RecordCount
        
        'Set first cell to be blank for dropdown
        mryVeterinarianDataTable(0, 0) = ("")
        mryVeterinarianDataTable(1, 0) = ("")
        arycboVeterinarianData(0) = ("")
        'Debug.Print "mryVeterinarianDataTable ("; 0 & "): " & mryVeterinarianDataTable(0, 0) & ", " & mryVeterinarianDataTable(1, 0)
        'Debug.Print "arycboVeterinarianData ("; intVeterinarianCounter & "): " & arycboVeterinarianData(intVeterinarianCounter)
        
        'Move to next row
        intVeterinarianCounter = 1
        
        Do Until qryVeterinariansTable.EOF
            'Get query data and populate the relevant rows with it, repeat until end of query
            'Data is sorted by order of surname
            mryVeterinarianDataTable(0, intVeterinarianCounter) = qryVeterinariansTable.Fields("txtID")
            mryVeterinarianDataTable(1, intVeterinarianCounter) = qryVeterinariansTable.Fields("txtFullname")
            'Debug.Print "mryVeterinarianDataTable ("; intVeterinarianCounter & "): " & mryVeterinarianDataTable(0, intVeterinarianCounter) & ", " & mryVeterinarianDataTable(1, intVeterinarianCounter)
            
            arycboVeterinarianData(intVeterinarianCounter) = qryVeterinariansTable.Fields("txtFullname")
            'Debug.Print "arycboVeterinarianData ("; intVeterinarianCounter & "): " & arycboVeterinarianData(intVeterinarianCounter)
            
            intVeterinarianCounter = intVeterinarianCounter + 1
            qryVeterinariansTable.MoveNext
        Loop
    
    End If
    
    If IsObject(qryVeterinariansTable) Then Set qryVeterinariansTable = Nothing
    'Reset counter for cbo populating
    intVeterinarianCounter = 0
    'Debug.Print UBound(arycboVeterinarianData)
    'Blank existing values from cbo
    Me.cboVeterinarian.RowSource = ""
    'Fill cbo with ary data
    For Each varVetElement In arycboVeterinarianData
        Me.cboVeterinarian.AddItem Item:=arycboVeterinarianData(intVeterinarianCounter), Index:=intVeterinarianCounter
        intVeterinarianCounter = intVeterinarianCounter + 1
    Next

End Sub

Private Sub btnEditCatPatientsForm_Click()
    Call frmCatPatientsSwitchLock
End Sub

Private Sub btnPrintReport_Click()
    Dim strDocName As String
    Dim strWhere As String
    strDocName = "rprtCatPatients"
    strWhere = "[autoEntryID]=" & Me!autoEntryID
    DoCmd.OpenReport strDocName, acPreview, , strWhere
End Sub

Private Sub cboOwner_Change()
    
    'Set value for ListIndex
    Dim lngOwner As Long
    
    'Get cbo ListIndex
    lngOwner = cboOwner.ListIndex
    Debug.Print (cboOwner.ListIndex)
    
    'If ListIndex is greater than 0 (ie an actual value)
    If (lngOwner > 0) Then
        If (mryOwnerDataTable(1, lngOwner) = arycboOwnerData(lngOwner)) Then
            Me.Owner.Value = mryOwnerDataTable(0, lngOwner)
            Debug.Print (mryOwnerDataTable(0, lngOwner))
        Else:
            Debug.Print ("Array Mismatch!: " & arycboOwnerData(lngOwner) & " Selected vs " & mryOwnerDataTable(1, lngOwner) & " Found")
        End If
    'If ListIndex is less than or equal to 0 (ie blank)
    ElseIf (lngOwner <= 0) Then
        Me.Owner.Value = mryOwnerDataTable(0, 0)
        Debug.Print ("Blank Owner selected!")
    End If
    
End Sub

Private Sub cboVeterinarian_Change()

    'Set value for ListIndex
    Dim lngVeterinarian As Long
    
    'Get cbo ListIndex
    lngVeterinarian = cboVeterinarian.ListIndex
    Debug.Print (cboVeterinarian.ListIndex)
    
    'If ListIndex is greater than 0 (ie an actual value)
    If (lngVeterinarian > 0) Then
        If (mryVeterinarianDataTable(1, lngOwner) = arycboVeterinarianData(lngOwner)) Then
            Me.Veterinarian.Value = mryVeterinarianDataTable(0, lngVeterinarian)
            Debug.Print (mryVeterinarianDataTable(0, lngOwner))
        Else:
            Debug.Print ("Array Mismatch!: " & arycboVeterinarianData(lngOwner) & " Selected vs " & mryVeterinarianDataTable(1, lngOwner) & " Found")
        End If
    'If ListIndex is less than or equal to 0 (ie blank)
    ElseIf (lngVeterinarian <= 0) Then
        Me.Veterinarian.Value = mryVeterinarianDataTable(0, 0)
        Debug.Print ("Blank Veterinarian selected!")
    End If
    
End Sub

Private Sub chckbxNeurologicalSigns_Click()
    Call frmCatPatientsNeuroBox
End Sub

Private Sub Form_Close()
If formOpen = "o" Then
  DoCmd.Close acForm, "subfrmCatSamples"
End If
End Sub

Private Sub frmCatPatients()
Dim strCond As String

strCond = "txtPatientID1 = Forms!frmCatPatients!txtPatientID1"

   If IsLoaded("subfrmCatSamples") Then
      [Forms]![subfrmCatSamples].FilterOn = True
      [Forms]![subfrmCatSamples].Filter = strCond
   End If
   Me!txtPatientID1.SetFocus

End Sub

Private Sub btnOpenCatSamples_Click()
    On Error GoTo Err_Open_subfrmCatSamples_Click
        'Initialize strings
        Dim stDocName As String
        Dim stLinkCriteria As String
    
        stDocName = "subfrmCatSamples"
        
        
        stLinkCriteria = "[txtPatientID1]=" & Me![txtPatientID1]
        Debug.Print (stLinkCriteria)
        DoCmd.OpenForm stDocName, , , stLinkCriteria
        formOpen = "o"
Exit_Open_subfrmCatSamples_Click:
        Exit Sub
    
Err_Open_subfrmCatSamples_Click:
        MsgBox Err.Description
        Resume Exit_Open_subfrmCatSamples_Click
        
End Sub

Private Sub lblOwner_Click()
On Error GoTo Err_Open_frmOwners_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frmOwners"
    
    stLinkCriteria = "[txtFullname]=" & Me.[cboOwner].Value
    Debug.Print stLinkCriteria
    DoCmd.OpenForm stDocName
    formOpen = "o"
    Forms("frmOwners").Form.Recordset.FindFirst "[txtFullname] = '" & Me.[cboOwner].Value & "'"
Exit_Open_frmCatPatients_Click:
    'DoCmd.Close acForm, "frmCatSamples", acSavePrompt
    Exit Sub

Err_Open_frmOwners_Click:
    MsgBox Err.Description
    Resume Exit_Open_frmCatPatients_Click
End Sub

Private Sub lblVeterinarian_Click()
On Error GoTo Err_Open_frmOwners_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frmVeterinarians"
    
    stLinkCriteria = "[txtFullname]=" & Me.[cboVeterinarian].Value
    Debug.Print stLinkCriteria
    DoCmd.OpenForm stDocName
    formOpen = "o"
    Forms("frmVeterinarians").Form.Recordset.FindFirst "[txtFullname] = '" & Me.[cboVeterinarian].Value & "'"
Exit_Open_frmCatPatients_Click:
    'DoCmd.Close acForm, "frmCatSamples", acSavePrompt
    Exit Sub

Err_Open_frmOwners_Click:
    MsgBox Err.Description
    Resume Exit_Open_frmCatPatients_Click
End Sub

