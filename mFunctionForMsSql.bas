Option Compare Database
Option Explicit

Public Function SQL_getConstraintsName(sTable As String, sColumn As String, Optional sConstraintType As String = "D") As String
    Dim Sql$, sConstraintsName$
    Dim Rst As Recordset
    
    Sql = _
        "SELECT obj_table.NAME      AS 'table'," & _
        "        columns.NAME        AS 'column'," & _
        "        obj_Constraint.NAME AS 'constraint'," & _
        "        obj_Constraint.type AS 'type'" & _
        "    FROM   sys.objects obj_table" & _
        "        JOIN sys.objects obj_Constraint" & _
        "            ON obj_table.object_id = obj_Constraint.parent_object_id" & _
        "        JOIN sys.sysconstraints constraints" & _
        "             ON constraints.constid = obj_Constraint.object_id" & _
        "        JOIN sys.columns columns" & _
        "             ON columns.object_id = obj_table.object_id" & _
        "            AND columns.column_id = constraints.colid" & _
        "    WHERE obj_table.NAME='{{TABLE_NAME}}'" & _
        "      AND columns.NAME='{{COLUMN_NAME}}'" & _
        "      AND obj_Constraint.type='{{CONSTRAINT_TYPE}}'"
    
    Sql = Replace(Sql, "{{TABLE_NAME}}", sTable)
    Sql = Replace(Sql, "{{COLUMN_NAME}}", sColumn)
    Sql = Replace(Sql, "{{CONSTRAINT_TYPE}}", sConstraintType)
    
    
    Dim con As ADODB.Connection
    Set con = CurrentProject.Connection
    con.CommandTimeout = 600

    Set Rst = New ADODB.Recordset

    Call Rst.Open(Sql, con, adOpenForwardOnly, adLockReadOnly, adCmdUnknown)
    
        If Rst.RecordCount = 1 Then
            sConstraintsName = Rst!constraint
        End If

    If Rst.State = adStateOpen Then
        Rst.Close
    End If

    Set Rst = Nothing
    
    SQL_getConstraintsName = sConstraintsName
End Function
