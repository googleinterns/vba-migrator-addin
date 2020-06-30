Imports Microsoft.Vbe.Interop
Imports System.Windows.Forms

Friend Class UserControlToolWindow
    Private _VBE As VBE
    Friend Sub Initialize(ByVal vbe As VBE, ByRef lines As List(Of String))
        _VBE = vbe
        'Write the data to the window.
        Call writeToDataGrid(lines)
    End Sub

    ''' <summary>
    ''' Iterate through the list, Add the row to DataGrid and put the list item value in it.
    ''' </summary>
    ''' <param name="lines">list of data to show in window.</param>
    Friend Sub writeToDataGrid(ByRef lines As List(Of String))
        Dim dataTable As DataGridView
        dataTable = DataGridView1
        dataTable.Rows.Clear()
        For i As Integer = 0 To lines.Count - 1 Step +1
            Dim getWord() As String
            getWord = Split(lines.Item(i), ",")
            dataTable.Rows.Add(getWord(0), getWord(1), CInt(getWord(2)), getWord(3))
        Next i
    End Sub

    ''' <summary>
    ''' Implementation of the event, When the DataGrid cell is clicked.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex > -1 Then
            DataGridView1.Rows(e.RowIndex).Selected = True
            Dim line As Integer
            Dim moduleName As String
            Dim _myVbComponents As VBComponents
            Dim _myVbComponent As VBComponent
            '==========================Activating the module================================================

            If DataGridView1.Rows(e.RowIndex).Cells(1).Value <> "" Then
                line = DataGridView1.Rows(e.RowIndex).Cells(2).Value
                moduleName = DataGridView1.Rows(e.RowIndex).Cells(1).Value
                _myVbComponents = _VBE.ActiveVBProject.VBComponents
                For Each _myVbComponent In _myVbComponents
                    If String.Compare(_myVbComponent.Name.ToLower, moduleName.ToLower) = 0 Then
                        'Module get activated.
                        _myVbComponent.Activate()
                        Exit For
                    End If
                Next
                '============================Passing the control==============================
                _VBE.ActiveCodePane.SetSelection(line, 1, line, 100)
                _VBE.ActiveCodePane.Show()
            End If

        End If
    End Sub

End Class
