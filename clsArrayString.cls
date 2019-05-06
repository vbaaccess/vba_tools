Option Compare Database
Option Explicit

Private Const CurrentModuleName = "clsArrayString"

Private Const DEFAULT_DIVIDER = ";"
Private Const DEFAULT_HEADS = "Lp;Description"

Private m_Divider As String
Private m_ArrHeads As String

Private m_ArrayString() As String

Public Event PrzedDodaniemRekordu(bCancel As Boolean)
Public Event PoDodaniemRekordu()

Private Sub Class_Initialize()
    m_Divider = DEFAULT_DIVIDER
    m_ArrHeads = DEFAULT_HEADS
    
    Call INITmArrayString
End Sub

Private Sub INITmArrayString()
    ReDim m_ArrayString(1) As String
    
    m_ArrayString(0) = m_ArrHeads

End Sub

Public Sub GET_ArrayString(ByRef objArrayString As Variant)
    objArrayString = m_ArrayString
End Sub

Public Property Get ArrHeads() As String
     ArrHeads = m_ArrHeads
End Property

Public Property Let ArrHeads(ByVal sNewHeads As String)
    Dim value_before As String
    
    value_before = m_ArrHeads
    
    m_ArrHeads = sNewHeads
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
    
    newLp = UBoundmArray(m_ArrayString) + 1
    sDane = newLp & m_Divider & sDane
    
    ReDim Preserve m_ArrayString(newLp)
        
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
