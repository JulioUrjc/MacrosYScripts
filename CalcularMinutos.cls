VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Segundos()

   Dim n, i, j As Long, c As String, min As String, seg As String
      
   For Each tiempo In Selection
    min = "0"
    seg = "0"
    
    If StrComp(tiempo.Offset(0, -3).Value, "34911620397") = False And Mid(tiempo.Offset(0, -2).Value, 1, 1) = 6 Then
    
      For n = 1 To Len(tiempo.Value)
       c = Mid(tiempo.Value, n, 1)
       Select Case c
       Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
         min = min & c
       Case "m"
         Exit For
       End Select
      Next
   
      For i = n + 1 To Len(tiempo.Value)
       c = Mid(tiempo.Value, i, 1)
       Select Case c
       Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
         seg = seg & c
       Case "s"
        Exit For
       End Select
      Next
            
      If tiempo.Offset(0, 0).EntireRow.Hidden = False Then
        tiempo.Offset(0, 1).Value = (min * 60) + seg
      End If
     End If
   Next
   
End Sub
