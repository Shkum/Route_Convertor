Module Module1
    Public c_FileType As String
    Public Function CheckFileType(fileName As String) As String
        Dim c_String As String
        If LCase(Right(FileName, 3)) = "xls" Then
            c_FileType = "Transas 3000 export file ('EXCEL')"
            Return "XLS"
        End If
        FileOpen(1, FileName, OpenMode.Input)
        c_String = LineInput(1)
        FileClose(1)

        Select Case Left(c_String, 7)
            Case "<TSH_Ro"
                c_FileType = "Transas 3000 or 4000 Route ('.RT3')"
                Return "RT3"
            Case ";Route "
                c_FileType = "Transas 3000 Route ('.CVT')"
                Return "CVT"
            Case "ROUTE N"
                c_FileType = "eGlobe G2 Route ('.G2TXT')"
                Return "G2TXT"
            Case "<?xml v"
                If Left(c_String, 15) = "<?xml version='" Then
                    c_FileType = "eGlobe G2 Route ('.BVS')"
                    Return "BVS"
                ElseIf Left(c_String, 15) = "<?xml version=""" Then
                    c_FileType = "eGlobe G2 Route ('.RTE')"
                    Return "RTE"
                End If
            Case Else
                c_FileType = "UNKNOWN FILE TYPE OR FILE VERSION"
                Return ""
        End Select

    End Function
End Module
