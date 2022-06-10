Imports System.Math
Module Module4
    Public c_Distance As Single
    Public Function GetCourse(lat1 As Single, lon1 As Single, lat2 As Single, lon2 As Single) As Single
        Dim cLatDiff, cLonDiff, cSumLatHalf, cCalc, cCalc1 As Single
        cLatDiff = (lat2 - lat1) * 60
        cLonDiff = Lon2 - lon1
        If cLonDiff > 180 Then
            cLonDiff = cLonDiff - 360
        ElseIf cLonDiff < -180 Then
            cLonDiff = 360 + cLonDiff
        End If
        cSumLatHalf = (lat1 + lat2) / 2
        cCalc = cLonDiff * Cos(cSumLatHalf * PI / 180) * 60
        c_Distance = Sqrt(cLatDiff ^ 2 + cCalc ^ 2)
        If cLatDiff <> 0 Then
            cCalc1 = Atan(cCalc / cLatDiff) * 180 / PI
        Else
            If cLonDiff < 0 Then
                cCalc1 = 270
            Else
                cCalc1 = 90
            End If
        End If
        If cLatDiff < 0 Then cCalc1 = cCalc1 + 180
        If cCalc1 < 0 Then cCalc1 = cCalc1 + 360
        Return cCalc1
    End Function


    Public Function LatLonConvert(Str As String) As Single
        Dim c_Gr, cSign As Integer, cMin, cConv As Single
        c_Gr = CInt(Left(Str, InStr(Str, Chr(176)) - 1))
        cMin = Val(Mid(Str, InStr(Str, Chr(176)) + 1, Str.Length - InStr(Str, Chr(176)) - 1))
        If Right(Str, 1) = "S" Or Right(Str, 1) = "W" Then
            cSign = -1
        Else
            cSign = 1
        End If
        cConv = (c_Gr + cMin / 60) * cSign
        Return cConv
    End Function
End Module

