'---------------------------------------------------------------------------------------
' Class Module  : clsLstSql
' Author        : paweld
' Purpose       : Klasa pozwalajaca zarzadzac dynamicznymi kolumnami
'
' Licence       : Free
'
'---------------------------------------------------------------------------------------
' Changes       :
'---------------------------------------------------------------------------------------
' 2019-05-28        , paweld : Utworzenie klasy.
'
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Const CurrentModName = "clsLstSql"

Private WithEvents objArrStr As clsArrayString
Private Const sDivider = ";"

Private Const sArrHeads = "ColLP;ColId;Width_cm;FieldSql;ColumnHeader;ColumnTipText;IsHiden;Translate"
Private Enum ENUM_COLUMN
    ColLp = 0
    ColId = 1
    Width_cm = 2
    FieldSql = 3
    ColumnHeader = 4
    ColumnTipText = 5
    IsHiden = 6
    Translate = 7
End Enum

Private m_ColumnCount%
Private m_ColumnWidths$
Private m_ColumnSqlSectionSelect$

Private Type Typ_LstParametryKolumny
    Id As Long
    LP As Long
    Width_cm As Double ' opis znakowy w centrymetrach
    Description As String
    Sql As String
    IsHiden As Boolean
End Type

Private Sub Class_Initialize()
    If objArrStr Is Nothing Then Set objArrStr = New clsArrayString
    objArrStr.ArrHeads = sArrHeads
    objArrStr.Divider = sDivider
End Sub

Public Function PutRecordColumnDescription(sDane As String) As Integer
    Dim newLp As Integer
    newLp = objArrStr.PutRecord(sDane)
    PutRecordColumnDescription = newLp
End Function

Public Property Get Divider() As String
     Divider = objArrStr.Divider
End Property

Public Property Let Divider(ByVal sNewDivider As String)
    objArrStr.Divider = sNewDivider
End Property

Public Property Get ColumnCount() As Integer
    ColumnCount = m_ColumnCount
End Property

Public Property Get ColumnWidths() As String
    ColumnWidths = m_ColumnWidths
End Property

Public Property Get ColumnSqlSectionSelect() As String
    ColumnSqlSectionSelect = m_ColumnSqlSectionSelect
End Property

Private Function CompleteTheParameters(Optional bNoVerification As Boolean) As Boolean
    Dim iCountColumnDefinition%
    
    '--- weryfikacja czy jest co uzupelniac ---
    If m_ColumnWidths = "" Or m_ColumnCount = 0 Or bNoVerification Then
        iCountColumnDefinition = objArrStr.ArrRowsCount
    End If
    
    '--- kompletowanie parametrow listy ---
    If iCountColumnDefinition > 0 Then CompleteTheParameters = True
End Function

Public Sub MakingPropertiesOfColumns(Optional bNoVerification As Boolean = True)
    
    If CompleteTheParameters(bNoVerification) = False Then Exit Sub
    
    '--- uzypelnienie (generowanie) parametrow Listy ---
    '    - sekcja SQL Select,
    '    - opisu szerokosci kolumn,
    '    - ilosci kolumn,
    
    Dim sCSS$           ' Column Sql Section Select
    Dim sColumnWidths$
    Dim iColumnCount%
    
    Dim iWerLp As Integer
    Dim werValue As String
    
    Dim ArrayString() As String
    Dim ArrFields
    
    Call objArrStr.GET_ArrayString(ArrayString)
    iColumnCount = objArrStr.ArrRowsCount
    
    For iWerLp = 1 To iColumnCount
        ArrFields = Split(ArrayString(iWerLp), sDivider)
        
        '--- set column width, includ status IsHiden ---
        werValue = "0"
        If CBool(ArrFields(ENUM_COLUMN.IsHiden)) = False Then
            werValue = ArrFields(ENUM_COLUMN.Width_cm) & "cm"
        End If
        sColumnWidths = sColumnWidths & ";" & werValue
        
        '--- set column sql ---
        werValue = ArrFields(ENUM_COLUMN.FieldSql)
        If Len(werValue) > 0 Then
            sCSS = sCSS & "," & werValue
            werValue = "": werValue = ArrFields(ENUM_COLUMN.ColumnHeader)
            If Len(werValue) > 0 Then
                If CBool(ArrFields(ENUM_COLUMN.Translate)) Then werValue = SUR.tr(werValue)
                sCSS = sCSS & " AS [" & werValue & "]"
            End If
        End If
    Next iWerLp

    If Len(sColumnWidths) > 0 Then
        sColumnWidths = Mid(sColumnWidths, 2)
    End If
    
    If Len(sCSS) > 0 Then
        sCSS = Mid(sCSS, 2)
    End If
    
    m_ColumnSqlSectionSelect = sCSS
    m_ColumnWidths = sColumnWidths
    m_ColumnCount = iColumnCount
    
End Sub
