' === 1 ==========================================================================================================================
' An example of the use of static variables
' This function allows you to reduce the frequency of parameter initialization
' It allows you to reduce the number of readings, which can improve performance, e.g. when reading them from the database
Public Function ShouldRefresStaticVariantVariable_TEST()
Const MinutesToRefresh = 2
  Static SomeVariable
  Static LastRefreshDateTime As Variant '=>Date
  Dim bRefres As Boolean
  
  If Not (bRefres) Then
    bRefres = ShouldRefresStaticVariantVariable(LastRefreshDateTime, MinutesToRefresh)
  End If
  
  If bRefres Then
    Debug.Print "Here, insert your code refreshing the variable, e.g. reading data from the database"
  End If
  
End Function

Public Function ShouldRefresStaticVariantVariable(StaticVariantVariableWithDate As Variant, MinuteRefreshInterval As Long) As Boolean

    If IsEmpty(StaticVariantVariableWithDate) Then   ' First time initialization
        StaticVariantVariableWithDate = VBA.Now()
    End If

    If IsDate(StaticVariantVariableWithDate) Then
        Dim MinutesSinceTheLastRefreshment As Long
        MinutesSinceTheLastRefreshment = VBA.DateDiff("n", VBA.CDate(StaticVariantVariableWithDate), VBA.Now())

        Debug.Print "Last refresh " & MinutesSinceTheLastRefreshment & " [min] ago"
        If MinutesSinceTheLastRefreshment >= MinuteRefreshInterval Then
            Debug.Print " ... time to refresh the variable"
            StaticVariantVariableWithDate = VBA.Now()
            ShouldRefresStaticVariantVariable = True
        End If
    End If
    
End Function

' === = ==========================================================================================================================
