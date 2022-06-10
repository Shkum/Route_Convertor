Imports Microsoft.Office.Interop
Imports System.Math
Module Module3

    Public Sub GetWayPoints(FileName As String, FileType As String)
        On Error GoTo 1
        Dim c_Lat, c_Lon, c_WpName, c_RlGc As String
        Dim c_WpNr, c_Count As Integer

        Select Case FileType

            Case "XLS"
                Dim ex As Excel.Application
                Dim c_Sheet As Excel.Worksheet
                ex = New Excel.Application
                ex.Workbooks.Open(FileName)
                c_Sheet = ex.Workbooks(1).Worksheets(1)
                If Left(c_Sheet.Cells.Range("A1").Text, 5) <> "Route" Then
                    ex.Workbooks.Close()
                    ex.Quit()
                    Exit Sub
                End If
                c_Count = 6
                Do While c_Sheet.Cells.Range("A" & c_Count).Text <> ""
                    Application.DoEvents()
                    c_WpNr = c_Sheet.Cells.Range("A" & c_Count).Value + 1

                    If c_Sheet.Cells.Range("B" & c_Count).Text = "" Then
                        c_WpName = "WP" & c_Sheet.Cells.Range("A" & c_Count).Value + 1
                    Else
                        c_WpName = c_Sheet.Cells.Range("B" & c_Count).Text
                    End If
                    c_Lat = c_Sheet.Cells.Range("C" & c_Count).Text
                    c_Lat = Replace(c_Lat, " ", "")
                    c_Lat = Replace(c_Lat, "'", "")
                    c_Lon = c_Sheet.Cells.Range("D" & c_Count).Text
                    c_Lon = Replace(c_Lon, " ", "")
                    c_Lon = Replace(c_Lon, "'", "")
                    c_RlGc = c_Sheet.Cells.Range("E" & c_Count).Text
                    If c_RlGc = "XXXX" Then c_RlGc = "RL"

                    Form1.ListBox1.Items.Add(c_WpNr & vbTab & c_WpName & vbTab & c_Lat & vbTab & c_Lon & vbTab & c_RlGc)
                    c_Count = c_Count + 1
                Loop
                ex.Quit()
               

            Case "RT3"
                Dim c_Text, c_txtLine, c_NSEW As String
                Dim c_txtArr As Array, c_Len As Integer
                FileOpen(1, FileName, OpenMode.Input)
                While Not EOF(1)
                    c_txtLine = LineInput(1)
                    c_Text = c_Text & c_txtLine
                End While
                FileClose(1)


                c_Text = Right(c_Text, c_Text.Length - InStr(c_Text, "<WayPoint "))
                c_Text = Left(c_Text, InStr(c_Text, "</WayPoints>") - 1)
                c_txtArr = Split(c_Text, "<WayPoint ")

                For c_Count = 0 To UBound(c_txtArr)
                    c_WpNr = c_Count + 1

                    c_Len = (InStr(InStr(c_txtArr(c_Count), "WPName=") + 8, c_txtArr(c_Count), """")) - (InStr(c_txtArr(c_Count), "WPName=") + 8)
                    c_WpName = Mid(c_txtArr(c_Count), InStr(c_txtArr(c_Count), "WPName=") + 8, c_Len)
                    If c_WpName = "" Then c_WpName = "WP" & c_Count + 1

                    c_Len = (InStr(InStr(c_txtArr(c_Count), "Lat=") + 5, c_txtArr(c_Count), """")) - (InStr(c_txtArr(c_Count), "Lat=") + 5)
                    c_Lat = Mid(c_txtArr(c_Count), InStr(c_txtArr(c_Count), "Lat=") + 5, c_Len)
                    If Left(c_Lat, 1) = "-" Then
                        c_NSEW = "S"
                    Else
                        c_NSEW = "N"
                    End If
                    c_Lat = Abs(Val(c_Lat))
                    c_Lat = c_Lat \ 60 & Chr(176) & Format(Round(c_Lat Mod 60, 3), "00.000") & c_NSEW
                    c_Lat = Replace(c_Lat, ",", ".")

                    c_Len = (InStr(InStr(c_txtArr(c_Count), "Lon=") + 5, c_txtArr(c_Count), """")) - (InStr(c_txtArr(c_Count), "Lon=") + 5)
                    c_Lon = Mid(c_txtArr(c_Count), InStr(c_txtArr(c_Count), "Lon=") + 5, c_Len)

                    c_Lon = Val(c_Lon)
                    If Val(c_Lon) > 10800 Then c_Lon = CStr(CDbl(c_Lon) - 21600)
                    If Val(c_Lon) < -10800 Then c_Lon = CStr(CDbl(c_Lon) + 21600)

                    If Left(c_Lon, 1) = "-" Then
                        c_NSEW = "W"
                    Else
                        c_NSEW = "E"
                    End If

                    c_Lon = Abs(CDbl(c_Lon))
                    c_Lon = c_Lon \ 60 & Chr(176) & Format(Round(c_Lon Mod 60, 3), "00.000") & c_NSEW
                    c_Lon = Replace(c_Lon, ",", ".")

                    c_Len = (InStr(InStr(c_txtArr(c_Count), "LegType=") + 9, c_txtArr(c_Count), """")) - (InStr(c_txtArr(c_Count), "LegType=") + 9)
                    c_RlGc = Mid(c_txtArr(c_Count), InStr(c_txtArr(c_Count), "LegType=") + 9, c_Len)
                    If c_RlGc = "0" Then
                        c_RlGc = "RL"
                    Else
                        c_RlGc = "GC"
                    End If

                    Form1.ListBox1.Items.Add(c_WpNr & vbTab & c_WpName & vbTab & c_Lat & vbTab & c_Lon & vbTab & c_RlGc)
                Next



            Case "CVT"
                Dim c_Text, c_txtLine As String
                Dim c_txtArr As Array, c_Len As Integer
                FileOpen(1, FileName, OpenMode.Input)
                While Not EOF(1)
                    c_txtLine = LineInput(1)
                    c_Text = c_Text & c_txtLine
                End While
                FileClose(1)
                c_Text = Right(c_Text, c_Text.Length - InStr(1, c_Text, "WP 001") + 1)
                c_txtArr = Split(c_Text, ";")

                For c_Count = 0 To UBound(c_txtArr)
                    c_WpNr = c_Count + 1
                    c_Len = (InStr(InStr(c_txtArr(c_Count), "Name ") + 5, c_txtArr(c_Count), "Lat ")) - (InStr(c_txtArr(c_Count), "Name ") + 5)
                    c_WpName = Mid(c_txtArr(c_Count), InStr(c_txtArr(c_Count), "Name ") + 5, c_Len)
                    If c_WpName = "" Then c_WpName = "WP" & c_Count + 1

                    c_Len = (InStr(InStr(c_txtArr(c_Count), "Lat ") + 5, c_txtArr(c_Count), "Lon ")) - (InStr(c_txtArr(c_Count), "Lat ") + 5)
                    c_Lat = Mid(c_txtArr(c_Count), InStr(c_txtArr(c_Count), "Lat ") + 5, c_Len)
                    c_Lat = Trim(c_Lat)
                    c_Lat = Replace(c_Lat, ",", ".")

                    c_Len = (InStr(InStr(c_txtArr(c_Count), "Lon ") + 5, c_txtArr(c_Count), "RL ")) - (InStr(c_txtArr(c_Count), "Lon ") + 5)
                    c_Lon = Mid(c_txtArr(c_Count), InStr(c_txtArr(c_Count), "Lon ") + 5, c_Len)
                    c_Lon = Trim(c_Lon)
                    c_Lon = Replace(c_Lon, ",", ".")

                    c_RlGc = "RL"
                    Form1.ListBox1.Items.Add(c_WpNr & vbTab & c_WpName & vbTab & c_Lat & vbTab & c_Lon & vbTab & c_RlGc)
                Next



            Case "G2TXT"
                Dim c_Text, c_txtLine As String
                Dim c_txtArr, c_WpArr As Array
                FileOpen(1, FileName, OpenMode.Input)
                While Not EOF(1)
                    c_txtLine = LineInput(1)
                    c_Text = c_Text & c_txtLine & Chr(13)
                End While
                FileClose(1)
                c_Text = Right(c_Text, c_Text.Length - InStr(1, c_Text, "WP1") + 1)
                c_Text = Left(c_Text, InStr(1, c_Text, "Legend:") - 1)
                Do While Right(c_Text, 1) = Chr(13)
                    c_Text = Left(c_Text, c_Text.Length - 1)
                Loop
                c_txtArr = Split(c_Text, Chr(13))

                For c_Count = 0 To UBound(c_txtArr)
                    c_WpArr = Split(c_txtArr(c_Count), ";")
                    c_WpNr = c_Count + 1
                    c_WpName = c_WpArr(0)
                    c_Lat = c_WpArr(1)
                    c_Lat = Replace(c_Lat, "?", Chr(176))
                    c_Lat = Replace(c_Lat, " ", "")
                    c_Lat = Replace(c_Lat, "'", "")
                    c_Lon = c_WpArr(2)
                    c_Lon = Replace(c_Lon, "?", Chr(176))
                    c_Lon = Replace(c_Lon, " ", "")
                    c_Lon = Replace(c_Lon, "'", "")
                    c_RlGc = c_WpArr(5)
                    If c_RlGc = "RhumbLine" Then
                        c_RlGc = "RL"
                    Else
                        c_RlGc = "GC"
                    End If

                    Form1.ListBox1.Items.Add(c_WpNr & vbTab & c_WpName & vbTab & c_Lat & vbTab & c_Lon & vbTab & c_RlGc)
                Next


            Case "BVS"
                Dim c_Text, c_txtLine, c_NSEW As String
                Dim c_txtArr As Array, c_Len As Integer
                FileOpen(1, FileName, OpenMode.Input)
                While Not EOF(1)
                    c_txtLine = LineInput(1)
                    c_Text = c_Text & c_txtLine
                End While
                FileClose(1)
                c_Text = Right(c_Text, c_Text.Length - InStr(1, c_Text, "<Position") + 1)
                c_Text = Left(c_Text, InStr(1, c_Text, "</TrackInfo>") - 1)
                c_txtArr = Split(c_Text, "/>")
                For c_Count = 0 To UBound(c_txtArr) - 1
                    c_WpNr = c_Count + 1
                    If InStr(c_txtArr(c_Count), "Name=") <> 0 Then
                        c_Len = (InStr(InStr(c_txtArr(c_Count), "Name=") + 6, c_txtArr(c_Count), """")) - (InStr(c_txtArr(c_Count), "Name=") + 6)
                        c_WpName = Mid(c_txtArr(c_Count), InStr(c_txtArr(c_Count), "Name=") + 6, c_Len)
                    Else
                        c_WpName = "WP" & c_Count + 1
                    End If
                    c_Len = (InStr(InStr(c_txtArr(c_Count), "Lat=") + 5, c_txtArr(c_Count), """")) - (InStr(c_txtArr(c_Count), "Lat=") + 5)
                    c_Lat = Mid(c_txtArr(c_Count), InStr(c_txtArr(c_Count), "Lat=") + 5, c_Len)
                    If Left(c_Lat, 1) = "-" Then
                        c_NSEW = "S"
                        c_Lat = Right(c_Lat, c_Lat.Length - 1)
                    Else
                        c_NSEW = "N"
                    End If
                    c_Lat = Int(Val(c_Lat)) & Chr(176) & Format(Round((Val(c_Lat) - Int(Val(c_Lat))) * 60, 3), "00.000") & c_NSEW
                    c_Lat = Replace(c_Lat, ",", ".")

                    
                    c_Len = (InStr(InStr(c_txtArr(c_Count), "Lon=") + 5, c_txtArr(c_Count), """")) - (InStr(c_txtArr(c_Count), "Lon=") + 5)
                    c_Lon = Mid(c_txtArr(c_Count), InStr(c_txtArr(c_Count), "Lon=") + 5, c_Len)
                    If Left(c_Lon, 1) = "-" Then
                        c_NSEW = "W"
                        c_Lon = Right(c_Lon, c_Lon.Length - 1)
                    Else
                        c_NSEW = "E"
                    End If
                    c_Lon = Int(Val(c_Lon)) & Chr(176) & Format(Round((Val(c_Lon) - Int(Val(c_Lon))) * 60, 3), "00.000") & c_NSEW
                    c_Lon = Replace(c_Lon, ",", ".")


                    c_Len = (InStr(InStr(c_txtArr(c_Count), "Navigation=") + 12, c_txtArr(c_Count), """")) - (InStr(c_txtArr(c_Count), "Navigation=") + 12)
                    c_RlGc = Mid(c_txtArr(c_Count), InStr(c_txtArr(c_Count), "Navigation=") + 12, c_Len)


                    Form1.ListBox1.Items.Add(c_WpNr & vbTab & c_WpName & vbTab & c_Lat & vbTab & c_Lon & vbTab & c_RlGc)
                Next


            Case "RTE"
                Dim c_Text, c_txtLine, c_NSEW As String
                Dim c_txtArr As Array, c_Len As Integer
                FileOpen(1, FileName, OpenMode.Input)
                While Not EOF(1)
                    c_txtLine = LineInput(1)
                    c_Text = c_Text & c_txtLine
                End While
                FileClose(1)
                c_Text = Right(c_Text, c_Text.Length - InStr(1, c_Text, "<waypoint") + 1)
                c_Text = Left(c_Text, InStr(1, c_Text, "</waypoints>") - 1)
                c_txtArr = Split(c_Text, "</waypoint>")
                For c_Count = 0 To UBound(c_txtArr) - 1
                    c_WpNr = c_Count + 1
                    If InStr(c_txtArr(c_Count), "waypoint name=") <> 0 Then
                        c_Len = (InStr(InStr(c_txtArr(c_Count), "waypoint name=") + 15, c_txtArr(c_Count), """")) - (InStr(c_txtArr(c_Count), "waypoint name=") + 15)
                        c_WpName = Mid(c_txtArr(c_Count), InStr(c_txtArr(c_Count), "waypoint name=") + 15, c_Len)
                    Else
                        c_WpName = "WP" & c_Count + 1
                    End If
                    
                    c_Len = (InStr(InStr(c_txtArr(c_Count), "latitude=") + 10, c_txtArr(c_Count), """")) - (InStr(c_txtArr(c_Count), "latitude=") + 10)
                    c_Lat = Mid(c_txtArr(c_Count), InStr(c_txtArr(c_Count), "latitude=") + 10, c_Len)
                    If Left(c_Lat, 1) = "-" Then
                        c_NSEW = "S"
                        c_Lat = Right(c_Lat, c_Lat.Length - 1)
                    Else
                        c_NSEW = "N"
                    End If
                    c_Lat = Int(Val(c_Lat)) & Chr(176) & Format(Round((Val(c_Lat) - Int(Val(c_Lat))) * 60, 3), "00.000") & c_NSEW
                    c_Lat = Replace(c_Lat, ",", ".")
                    
                    c_Len = (InStr(InStr(c_txtArr(c_Count), "longitude=") + 11, c_txtArr(c_Count), """")) - (InStr(c_txtArr(c_Count), "longitude=") + 11)
                    c_Lon = Mid(c_txtArr(c_Count), InStr(c_txtArr(c_Count), "longitude=") + 11, c_Len)

                    If Left(c_Lon, 1) = "-" Then
                        c_NSEW = "W"
                        c_Lon = Right(c_Lon, c_Lon.Length - 1)
                    Else
                        c_NSEW = "E"
                    End If

                    c_Lon = Int(Val(c_Lon)) & Chr(176) & Format(Round((Val(c_Lon) - Int(Val(c_Lon))) * 60, 3), "00.000") & c_NSEW
                    c_Lon = Replace(c_Lon, ",", ".")


                    c_Len = (InStr(InStr(c_txtArr(c_Count), "<legType>") + 9, c_txtArr(c_Count), "</legType>")) - (InStr(c_txtArr(c_Count), "<legType>") + 9)
                    c_RlGc = Mid(c_txtArr(c_Count), InStr(c_txtArr(c_Count), "<legType>") + 9, c_Len)
                    If c_RlGc = "rhumbline" Then
                        c_RlGc = "RL"
                    Else
                        c_RlGc = "GC"
                    End If

                    Form1.ListBox1.Items.Add(c_WpNr & vbTab & c_WpName & vbTab & c_Lat & vbTab & c_Lon & vbTab & c_RlGc)
                Next

        End Select
1:
        Select Case Err.Number
            Case 0
            Case Else
                MsgBox("File opening error." & vbCrLf & "Error " & Err.Number & " - " & Err.Description, vbExclamation, "Error saving file")
        End Select
    End Sub
End Module
