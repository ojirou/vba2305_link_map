Attribute VB_Name = "Module1"
'
'#############################################################################
' 【EXL】 マップにリンクを張る'
'
'　link_map
'#############################################################################
'Option Explicit
Sub LinkMap()
Dim SheetName As String
Workbooks.Open ("C:\Users\user\git\excel_vba\data_table\list_map_sample.xlsx")
Set wb = ActiveWorkbook
Sheets("map").Activate
For Each a In Range("c3:g12")
    Sign = a.Value
    If Sign <> "" Then
        For Each Rng In wb.Worksheets("list").Range("d3:d20")
            If Rng.Value = Sign Then
                row_num = Rng.Row
                Set H = a.Hyperlinks.Add(a, "")
                With H
                    i_str = CStr(row_num)
                    rangestr = "#list!B" & i_str & ":J" & i_str
                    .SubAddress = rangestr
                    .TextToDisplay = wb.Worksheets("list").Cells(row_num, 13)
                End With
                With a.CurrentRegion
                    .Columns(3).ShrinkToFit = True
                    .Columns(2).WrapText = True
                    .Columns(3).WrapText = True
                    .Columns(4).WrapText = True
                    .Columns(5).WrapText = True
                    .Columns(6).WrapText = True
                    .Columns(2).ShrinkToFit = True
                    .Columns(3).ShrinkToFit = True
                    .Columns(4).ShrinkToFit = True
                    .Columns(5).ShrinkToFit = True
                    .Columns(6).ShrinkToFit = True
'                    .Font.Size
                End With
                If wb.Worksheets("list").Cells(row_num, 14) = "A" Then
                    wb.Worksheets("map").Activate
                        With a
                            .Font.ColorIndex = 3
                        End With
                ElseIf wb.Worksheets("list").Cells(row_num, 14) = "B" Then
                     wb.Worksheets("map").Activate
                         With a
                            .Font.ColorIndex = 5
                        End With
                Else
                     wb.Worksheets("map").Activate
                         With a
                            .Font.ColorIndex = 1
                         End With
                End If
                
            End If
        Next
    End If
Next
End Sub
