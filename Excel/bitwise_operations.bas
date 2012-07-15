' This script adds functions for bitwise AND, OR, and XOR operations.
'
' Use these functions within the sheet to perform bitwise operations
' on the cells in your worksheet
'

Attribute VB_Name = "BitWise"

Public Function BITXOR(ParamArray Args()) As Variant
    Dim arg As Variant
    Dim value As Variant
    Dim transitional As Variant
    For Each arg In Args
        Select Case TypeName(arg)
            Case "Range"
            
            Case "Array"
            
            Case Else
                transitional = arg
        End Select
        If IsEmpty(value) Then
            value = transitional
        Else
            value = value Xor transitional
        End If
    Next arg
        
    BITXOR = value
End Function

Public Function BITAND(ParamArray Args()) As Variant
    Dim arg As Variant
    Dim value As Variant
    Dim transitional As Variant
    For Each arg In Args
        Select Case TypeName(arg)
            Case "Range"
                For Each cell In arg.Cells
                    If IsEmpty(transitional) Then
                        transitional = cell.value
                    Else
                        transitional = transitional And cell.value
                    End If
                Next cell
            Case Else
                transitional = arg
        End Select
        If IsEmpty(value) Then
            value = transitional
        Else
            value = value And transitional
        End If
    Next arg
        
    BITAND = value
End Function

Public Function BITOR(ParamArray Args()) As Variant
    Dim arg As Variant
    Dim value As Variant
    Dim transitional As Variant
    For Each arg In Args
       Select Case TypeName(arg)
            Case "Range"
                For Each cell In arg.Cells
                    If IsEmpty(transitional) Then
                        transitional = cell.value
                    Else
                        transitional = transitional Or cell.value
                    End If
                Next cell
            Case Else
                transitional = arg
        End Select
        If IsEmpty(value) Then
            value = transitional
        Else
            value = value Or transitional
        End If
    Next arg
        
    BITOR = value
End Function
