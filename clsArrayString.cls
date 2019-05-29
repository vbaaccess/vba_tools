Option Compare Database
Option Explicit

Private Const CurrentModuleName = "clsArrayString"

Private Const DEFAULT_DIVIDER = ";"
Private Const DEFAULT_CriteriaDivider = ":="
Private Const DEFAULT_HEADS = "Lp;Description"

Private Enum search
    Head = 0
    Value = 1
    HeadExist = 2
    HeadLp = 3
    ValueExist = 4
End Enum

Private m_Divider As String
Private m_CriteriaDivider As String
Private m_ArrHeads As String

Private m_ArrayString() As String

Public Event PrzedDodaniemRekordu(bCancel As Boolean)
Public Event PoDodaniemRekordu()

Private Sub Class_Initialize()
    Debug.Print "Konstruktor klasy " & CurrentModuleName & " (" & VBA.Now & ")"
    
    m_Divider = DEFAULT_DIVIDER
    m_ArrHeads = DEFAULT_HEADS
    m_CriteriaDivider = DEFAULT_CriteriaDivider
    
    Call INITmArrayString
End Sub

Private Sub INITmArrayString()
    If UBoundmArray(m_ArrayString) = 0 Then
        ReDim m_ArrayString(1) As String
    End If
    
    m_ArrayString(0) = m_ArrHeads

End Sub

Public Sub GET_ArrayString(ByRef objArrayString As Variant)
    objArrayString = m_ArrayString
End Sub

Public Function GET_RecordString(Criteria As String) As String
    Dim sRecordString As String

    sRecordString = GET_Data(Criteria)
    
    GET_RecordString = sRecordString
End Function

Public Function GET_FieldString(Criteria As String, FieldName As String) As String
    Dim sFieldString As String

    sFieldString = GET_Data(Criteria, FieldName)
    
    GET_FieldString = sFieldString
End Function

Public Function SET_FieldValue(Criteria As String, FieldName As String, NewValue As String) As Boolean
    SET_FieldValue = SET_Data(Criteria, FieldName, NewValue)
End Function

' wyszukuje pierwszy rekord (pierwsza wartosc pola) spelaniajacy wszystkie warunki
Private Function SET_Data(Criteria As String, FieldName As String, NewValue As String) As Boolean
    Dim vLp As Variant
    Dim vCol As Variant
    Dim RowLp As Single
    Dim sRow As String
    Dim newRow As String
    Dim iLpCol As Single
    
    vLp = GET_Data(Criteria, , 1)
    vCol = GET_HeadLpByName(FieldName)
    
    If IsNumeric(vLp) And IsNumeric(vCol) Then
        
        sRow = GET_Data(Criteria)
        
        Dim werVal
        
        '--- analiza pol ---
        For Each werVal In Split(sRow, m_Divider)
            iLpCol = iLpCol + 1
            
            If iLpCol = CSng(vCol) Then
                newRow = newRow & m_Divider & NewValue
            Else
                newRow = newRow & m_Divider & werVal
            End If
            
        Next werVal
        
        
        '--- nadpisanie wiersza
        If Len(newRow) > 0 Then
            newRow = Mid(newRow, 2)
            m_ArrayString(CSng(vLp)) = newRow
        End If

        SET_Data = True
    End If
    
End Function

' wyszukuje pierwszy rekord (pierwsza wartosc pola) spelaniajacy wszystkie warunki
Private Function GET_Data(Criteria As String, Optional FieldName As String, Optional iFieldNameHeadLp As Single) As String
    Dim sData As String

    Dim sArr
    Dim sRow, sHead, sValu
    Dim ColLp As Single
    Dim RowLp As Single
    Dim arrLp As Long
    Dim HeadLp As Single
    
    Dim LpCriteria As Single
    Dim countCriteria As Single
    Dim countCriteriaWerOK As Single
    Dim arrCriteria() As String
    
    Dim sValueString As String
    
    Dim iRerkordFoundLp As Single
    
    Dim sReturnString As String
    
    countCriteria = SET_ArrayFromCriteria(Criteria, arrCriteria)
    
    If Len(Trim(FieldName)) = 0 Then
        If iFieldNameHeadLp > 0 Then
            FieldName = GET_HeadNameByLp(iFieldNameHeadLp)
        End If
    Else
        iFieldNameHeadLp = GET_HeadLpByName(FieldName)
    End If
    
    For Each sRow In m_ArrayString
        arrLp = arrLp + 1
        
        If arrLp = 1 Then
        '--- weryfikacja naglowkow - czy istnieja --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
            
            HeadLp = 0
            For Each sHead In Split(sRow, m_Divider)
                HeadLp = HeadLp + 1
                For LpCriteria = 1 To countCriteria
                
                    If CStr(arrCriteria(LpCriteria, search.Head)) = CStr(sHead) Then
                        arrCriteria(LpCriteria, search.HeadExist) = CSng(arrCriteria(LpCriteria, search.HeadExist)) + 1
                        arrCriteria(LpCriteria, search.HeadLp) = HeadLp
                    End If
                
                Next LpCriteria

'                If Len(Trim(FieldName)) > 0 And iFieldNameHeadLp < 1 Then
'                    If FieldName = CStr(sHead) Then
'                        iFieldNameHeadLp = HeadLp
'                    End If
'                End If
            Next sHead
            
        '--- weryfikacja - czy wszystkie naglowki sa ---
            Dim bAllHeadsExistOneTime As Boolean
            
            bAllHeadsExistOneTime = ArrayFromCriteria_WerHeadExist(arrCriteria)
            
            If Not (bAllHeadsExistOneTime) Then Exit For
        '--- ------------------------------------ --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
        Else
        '--- weryfikacja danych - czy istnieja -- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
            If Len(Trim(sRow)) > 0 Then
            
            HeadLp = 0
            countCriteriaWerOK = 0
            For Each sValu In Split(sRow, m_Divider)
                HeadLp = HeadLp + 1
                For LpCriteria = 1 To countCriteria
                    If CSng(arrCriteria(LpCriteria, search.HeadLp)) = HeadLp Then
                        '--- spradzenie czy wartosc spelnia warunek ---
                        If CStr(arrCriteria(LpCriteria, search.Value)) = CStr(sValu) Then
                            arrCriteria(LpCriteria, search.ValueExist) = CSng(arrCriteria(LpCriteria, search.ValueExist)) + 1
                            countCriteriaWerOK = countCriteriaWerOK + 1
                        End If
                        '--- -------------------------------------- ---
                    End If
                Next LpCriteria
                    
                If countCriteriaWerOK = countCriteria Then iRerkordFoundLp = arrLp
                    
                If iFieldNameHeadLp = HeadLp Then
                    sValueString = CStr(sValu)
                End If
            Next sValu
                
            ' sprwadzenie czywszystkie kryteria sa spelnione, jestli tak to mamy pierwszy rekord i z tego zwracamy wynik
            If iRerkordFoundLp > 0 Then
                If Len(Trim(FieldName)) > 0 And iFieldNameHeadLp > 0 Then
                    sReturnString = sValueString
                Else
                    sReturnString = sRow
                End If
                
                Exit For
            End If
                
            End If
        '--- ------------------------------------ --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
        End If
    Next sRow

    GET_Data = sReturnString
End Function

Private Function GET_HeadNameByLp(ByVal HeadLp As Single) As String
    Dim sHeadRow As String
    Dim sHead
    Dim RetHeadName As String
    Dim iLp As Single
    
    sHeadRow = m_ArrayString(0)
    
    For Each sHead In Split(sHeadRow, m_Divider)
        iLp = iLp + 1
        
        If HeadLp = iLp Then _
            RetHeadName = CStr(sHead)
        
    Next sHead
    
    GET_HeadNameByLp = RetHeadName
End Function

Private Function GET_HeadLpByName(ByVal HeadName As String) As Single
    Dim sHeadRow As String
    Dim RetHeadLp As Single
    Dim sHead
    Dim HeadCount As Single
    Dim HeadLp As Single
    
    sHeadRow = m_ArrayString(0)
    
    For Each sHead In Split(sHeadRow, m_Divider)
        HeadLp = HeadLp + 1
        If HeadName = CStr(sHead) Then
            RetHeadLp = HeadLp
        End If
    Next sHead
    
    If HeadCount > 1 Then RetHeadLp = -1 ' naglowek wystepuje wiecej niz raz!
    
    GET_HeadLpByName = RetHeadLp
End Function


Private Function SET_ArrayFromCriteria(ByVal Criteria As String, ByRef objArrayCriteria() As String) As Single
    Dim werCriteria
    Dim werSearch
    Dim LpCriteria As Single
    Dim countCriteriaDivider As Single
    Dim countCriteria As Single

    countCriteriaDivider = Len(Criteria) - Len(Replace(Criteria, m_Divider, ""))
    countCriteria = countCriteriaDivider + 1
    

    ReDim Preserve objArrayCriteria(1 To countCriteria, 0 To 4)
    
    For Each werCriteria In Split(Criteria, m_Divider)
        LpCriteria = LpCriteria + 1
        werSearch = Split(werCriteria, m_CriteriaDivider)
       'ReDim Preserve objArrayCriteria(1 To LpCriteria, 0 To 3)
        
        objArrayCriteria(LpCriteria, search.Head) = werSearch(search.Head)
        objArrayCriteria(LpCriteria, search.Value) = werSearch(search.Value)
        objArrayCriteria(LpCriteria, search.HeadExist) = 0
        objArrayCriteria(LpCriteria, search.HeadLp) = 0
        objArrayCriteria(LpCriteria, search.ValueExist) = 0
    Next werCriteria
    SET_ArrayFromCriteria = LpCriteria
End Function

Private Function ArrayFromCriteria_WerHeadExist(ByRef arrCriteria() As String, Optional ErrorMsg As String) As Boolean
    Dim bError As Boolean
    
    Dim iBrakNaglowkow As Single
    Dim iZdublowanychNaglowkow As Single
    Dim LpCriteria As Single
    Dim countCriteria As Single
    
    countCriteria = UBoundmArray(arrCriteria)
    
    For LpCriteria = 1 To countCriteria
        If arrCriteria(LpCriteria, search.HeadExist) = 0 Then
            iBrakNaglowkow = iBrakNaglowkow + 1
        End If
        
        If arrCriteria(LpCriteria, search.HeadExist) > 1 Then
            iZdublowanychNaglowkow = iZdublowanychNaglowkow + 1
        End If
    Next LpCriteria
    
    If iBrakNaglowkow > 0 Then
        ErrorMsg = ErrorMsg & vbLf & "-brak szukanego naglowka" & " (" & iBrakNaglowkow & "x)"
        bError = True
    End If
    
    If iZdublowanychNaglowkow > 0 Then
        ErrorMsg = ErrorMsg & vbLf & "-zdublowane naglow" & " (" & iZdublowanychNaglowkow & "x)"
        bError = True
    End If
    
    ArrayFromCriteria_WerHeadExist = Not (bError)
End Function

Private Function ArrayFromCriteria_WerValueExist(ByRef objArrayCriteria() As String, Optional ErrorMsg As String) As Boolean
    Dim bError As Boolean
    
    bError = True
    VBA.MsgBox "UNDER CONSTRUCTION", vbInformation + vbOKOnly, "under construction"
    
    ArrayFromCriteria_WerValueExist = Not (bError)
End Function

Public Property Get ArrRewsCount() As Long
    ArrRewsCount = UBoundmArray(m_ArrayString)
End Property

Public Property Get ArrHeads() As String
     ArrHeads = m_ArrHeads
End Property

Public Property Let ArrHeads(ByVal sNewHeads As String)
    Dim value_before As String
    
    value_before = m_ArrHeads
    
    m_ArrHeads = sNewHeads
    
    Call INITmArrayString
End Property

Public Property Get Divider() As String
     Divider = m_Divider
End Property

Public Property Let Divider(ByVal sNewDivider As String)
    Dim value_before As String
    
    value_before = m_Divider
    
    m_Divider = sNewDivider
End Property

Public Function AddRecord(iLp As Integer, sDane As String) As Boolean
    Dim bCancel As Boolean
    
    If iLp > UBoundmArray(m_ArrayString) Then
        ReDim Preserve m_ArrayString(iLp)
    End If
        
    RaiseEvent PrzedDodaniemRekordu(bCancel)
    
    If Not (bCancel) Then
        m_ArrayString(iLp) = sDane
        AddRecord = True
    End If
    
    RaiseEvent PoDodaniemRekordu
End Function

Public Function PutRecord(sDane As String) As Long
    Dim newLp As Long
    Dim bCancel As Boolean
    
    newLp = UBoundmArray(m_ArrayString)
    sDane = newLp & m_Divider & sDane
    
    ReDim Preserve m_ArrayString(newLp + 1)
        
    RaiseEvent PrzedDodaniemRekordu(bCancel)
    
    If Not (bCancel) Then
        m_ArrayString(newLp) = sDane
        PutRecord = newLp
    End If
    
    RaiseEvent PoDodaniemRekordu
End Function

Private Function UBoundmArray(werArray As Variant) As Long
On Error Resume Next
    UBoundmArray = UBound(werArray)
End Function

Private Sub Class_Terminate()
    Debug.Print "Destruktor klasy " & CurrentModuleName & " (" & VBA.Now & ")"
End Sub

