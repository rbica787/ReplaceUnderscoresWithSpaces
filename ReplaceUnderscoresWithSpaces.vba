Option Explicit

Public Sub ReplaceUnderscoresWithSpaces()

    Dim cell As Range

    If TypeName(Selection) <> "Range" Then Exit Sub

    Application.ScreenUpdating = False

    For Each cell In Selection
        If Not IsError(cell.Value) Then
            If Len(cell.Value) > 0 Then
                cell.Value = Replace(CStr(cell.Value), "_", " ")
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

End Sub
