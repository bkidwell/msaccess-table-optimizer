' TableOptimizer for Microsoft Access
'
' by Brendan Kidwell
' 17 December 2007
'
' This module is a quick and dirty procedure for optimizing the sizes
' of Text fields in Access. One important use for this module would be
' cleaning up tables imported from external sources where Access
' defaults to very large field sizes.
'
' The report script in this module create a table called
' _TableOptimizer that lists the field size, size of shortest value, and
' size of longest value in the table you specify. Then if you want to
' resize any text fields, you can edit the report table and run the
' resizing script.
'
' Usage:
'
' 1. Create a new module called "TableOptimizer". Edit that module in
' the Visual Basic window (ALT-F11) and paste this code, replacing the
' stub code you find there initially.
'
' 2. Go to the Immediate window (CTRL-G) and run the method
' "TableOptimizer.MakeReport". Enter the name of the table you want to
' analyze. The script will catalog the fields and their minimum and
' maximum sizes. WARNING: This will create a table called
' "_TableOptimizer". If you already have such a table in your database,
' you must modify the code here appropriately.
'
' 3. Go to the Database window and open the "_TableOptimizer" window.
' Optionally filter and sort the table. (If you have run the script on
' more than one table, you will probably want to filter the
' _TableOptimizer table to show only records for that table.) Look at
' the "Size", "Shortest_Value", and "Longest_Value" numbers given for
' each Text field.
'
' 4. Enter a new value in the "New_Size" column for any Text fields you
' want to resize. Leave "New_Size" empty for any other fields.
'
' 5. Go to the Immediate window (CTRL-G) and run the method
' "TableOptimizer.ChangeSizes". Again, enter the name of the table
' you're working on. The script will run a series of ALTER TABLE SQL
' statements to resize the Text fields as you specified. If there are
' any fields whose type you want to change (for example, Memo to Text),
' you must do this MANUALLY from the Design view of the target table.
'
'
'
' Copyright (c) 2007 Brendan Kidwell
'
' Permission is hereby granted, free of charge, to any person obtaining
' a copy of this software and associated documentation files (the
' "Software"), to deal in the Software without restriction, including
' without limitation the rights to use, copy, modify, merge, publish,
' distribute, sublicense, and/or sell copies of the Software, and to
' permit persons to whom the Software is furnished to do so, subject to
' the following conditions:
'
' The above copyright notice and this permission notice shall be
' included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
' EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
' MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
' CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
' TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
' SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

 
 
Option Compare Database
Option Explicit
 
Public Sub MakeReport()
 
Dim target As String, f As DAO.Field, t As DAO.TableDef
Dim rs1 As DAO.Recordset, rs2 As DAO.Recordset
Dim i As Long, sql1 As String, typeName As String, size As Variant
Dim sql2 As String, fieldName As String
 
On Error Resume Next
CurrentDb.Execute _
    "create table _TableOptimizer " & _
    "([Table] Text (50), [Field_Num] Integer, [Field] Text (50), " & _
    "[Type] Text (20), Size Long, New_Size Long, Shortest_Value Long, Longest_Value Long)"
On Error GoTo 0
 
target = InputBox("Which table", "TableOptimizer: MakeReport()", "")
If target = "" Then Exit Sub
 
CurrentDb.Execute "delete from _TableOptimizer where [Table]=""" & target & """"
 
Set rs1 = CurrentDb.OpenRecordset(target)
sql1 = "select 0 as x"
i = 0
For Each f In rs1.Fields
    i = i + 1
    size = f.size
    typeName = FieldTypeName(f)
    If typeName = "Memo" Then size = "null"
 
    sql2 = _
        "insert into _TableOptimizer " & _
        "([Table], [Field_Num], [Field], [Type], [Size]) " & _
        "values (""" & target & """, " & i & ", """ & f.Name & """, " & _
        """" & typeName & """, " & size & ")"
    CurrentDb.Execute sql2
 
    Select Case typeName
    Case "Text", "Memo"
        sql1 = sql1 & _
            ", Max(Len([" & f.Name & "])) as [Max" & i & "] " & _
            ", Min(Len([" & f.Name & "])) as [Min" & i & "] "
    End Select
Next
rs1.Close
 
sql1 = sql1 & "from [" & target & "]"
Set rs1 = CurrentDb.OpenRecordset(sql1)
rs1.MoveFirst
 
Set rs2 = CurrentDb.OpenRecordset( _
    "select * from _TableOptimizer where [Table]=""" & target & """ " & _
    "order by Field_Num" _
)
 
rs2.MoveFirst
Do Until rs2.EOF
    typeName = rs2("Type").Value
    fieldName = rs2("Field").Value
    i = rs2("Field_Num").Value
    Select Case typeName
    Case "Text", "Memo"
        rs2.Edit
        rs2("Shortest_Value").Value = rs1("Min" & i).Value
        rs2("Longest_Value").Value = rs1("Max" & i).Value
        rs2.Update
    End Select
    rs2.MoveNext
Loop
 
rs2.Close
rs1.Close
 
End Sub
 
Public Sub ChangeSizes()
 
Dim target As String, rs As DAO.Recordset, sql
 
target = InputBox("Which table", "TableOptimizer: ChangeSizes()", "")
If target = "" Then Exit Sub
 
Set rs = CurrentDb.OpenRecordset( _
    "select * from _TableOptimizer where [table]=""" & target & """" _
)
 
rs.MoveFirst
Do Until rs.EOF
    Debug.Print rs("type").Value, rs("new_size").Value
    If rs("type").Value = "Text" And Not IsNull(rs("new_size").Value) Then
        sql = _
            "alter table [" & target & "] " & _
            "alter column [" & rs("field").Value & "] text (" & rs("new_size").Value & ")"
        CurrentDb.Execute sql
    End If
    rs.MoveNext
Loop
rs.Close
 
End Sub
 
' FieldTypeName
' by Allen Browne, allen@allenbrowne.com. Updated June 2006.
' copied from http://allenbrowne.com/func-06.html
' (No license information found at that URL.)
Function FieldTypeName(fld As DAO.Field) As String
    'Purpose: Converts the numeric results of DAO Field.Type to text.
    Dim strReturn As String    'Name to return

    Select Case CLng(fld.Type) 'fld.Type is Integer, but constants are Long.
        Case dbBoolean: strReturn = "Yes/No"            ' 1
        Case dbByte: strReturn = "Byte"                 ' 2
        Case dbInteger: strReturn = "Integer"           ' 3
        Case dbLong                                     ' 4
            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
            Else
                strReturn = "AutoNumber"
            End If
        Case dbCurrency: strReturn = "Currency"         ' 5
        Case dbSingle: strReturn = "Single"             ' 6
        Case dbDouble: strReturn = "Double"             ' 7
        Case dbDate: strReturn = "Date/Time"            ' 8
        Case dbBinary: strReturn = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
            Else
                strReturn = "Text (fixed width)"        '(no interface)
            End If
        Case dbLongBinary: strReturn = "OLE Object"     '11
        Case dbMemo                                     '12
            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
            Else
                strReturn = "Hyperlink"
            End If
        Case dbGUID: strReturn = "GUID"                 '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt: strReturn = "Big Integer"        '16
        Case dbVarBinary: strReturn = "VarBinary"       '17
        Case dbChar: strReturn = "Char"                 '18
        Case dbNumeric: strReturn = "Numeric"           '19
        Case dbDecimal: strReturn = "Decimal"           '20
        Case dbFloat: strReturn = "Float"               '21
        Case dbTime: strReturn = "Time"                 '22
        Case dbTimeStamp: strReturn = "Time Stamp"      '23

        'Constants for complex types don't work prior to Access 2007.
        Case 101&: strReturn = "Attachment"         'dbAttachment
        Case 102&: strReturn = "Complex Byte"       'dbComplexByte
        Case 103&: strReturn = "Complex Integer"    'dbComplexInteger
        Case 104&: strReturn = "Complex Long"       'dbComplexLong
        Case 105&: strReturn = "Complex Single"     'dbComplexSingle
        Case 106&: strReturn = "Complex Double"     'dbComplexDouble
        Case 107&: strReturn = "Complex GUID"       'dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"    'dbComplexDecimal
        Case 109&: strReturn = "Complex Text"       'dbComplexText
        Case Else: strReturn = "Field type " & fld.Type & " unknown"
    End Select
 
    FieldTypeName = strReturn
End Function

