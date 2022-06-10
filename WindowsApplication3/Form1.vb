Imports Microsoft.Office.Interop
Imports Microsoft.Office
Imports System.Math
Imports Microsoft.Win32
Public Class Form1
    Dim cFix As Boolean = False

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim c_fileName As String
        OpenFileDialog1.FileName = ""
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            c_fileName = OpenFileDialog1.FileName
            ListBox1.Items.Clear()
            Open_File(c_fileName)
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Application.Exit()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListBox1.CustomTabOffsets.Add(110)
       End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        MsgBox("Created by Sergey Shkumat (pepelnici@meta.ua)" & vbCrLf & vbCrLf & "Tested by Sergey Korolyov" & _
               vbCrLf & vbCrLf & "Convert routes between eGlobe G2 and Transas 3000/4000" & vbCrLf & vbCrLf & "Can open:" & vbCrLf & "eGlobe 2G - .rte, .g2txt, .bvs" & _
               vbCrLf & "Transas 3000/4000 - .cvt, .rt3, .xls (exported routes from" & vbCrLf & "Transas 3000/4000 to Excel)" & _
               vbCrLf & vbCrLf & "Can export route to BonVoyage (Weather) system" & _
               vbCrLf & vbCrLf & "Remark: in case of problems with export to Excel," & _
               vbCrLf & "install 'MS Office Primary Interop Assemblies Package'" & _
               vbCrLf & "(PIARedist.exe) for your version of MS Office.", vbInformation, "About...")
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Button3.Text = "Export to eGlobe G2: .RTE"
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Button3.Text = "Export to Transas 3000/4000: .RT3"
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        On Error GoTo 1
        Dim c_TXT, C_FileName, c_Str(), c_FN As String

        If RadioButton1.Checked = True Then 'RTE
            c_TXT = WPExt("RTE")
            SaveFileDialog1.DefaultExt = "rte"
            SaveFileDialog1.Filter = "eGlobe files (*.rte)|*.rte"

        ElseIf RadioButton2.Checked = True Then 'RT3
            c_TXT = WPExt("RT3")
            SaveFileDialog1.DefaultExt = "rt3"
            SaveFileDialog1.Filter = "Transas 3000 and 4000 files (*.rt3)|*.rt3"


        ElseIf RadioButton4.Checked = True Then 'BVS
            c_TXT = WPExt("BVS")
            SaveFileDialog1.DefaultExt = "bvs"
            SaveFileDialog1.Filter = "BonVoyage files (*.bvs)|*.bvs"


        ElseIf RadioButton3.Checked = True Then 'EXCEL
            Dim ex As Excel.Application, c_Sheet As Excel.Worksheet, c_WorkBook As Excel.Workbook, c_Count As Integer

            SaveFileDialog1.DefaultExt = "xls"
            SaveFileDialog1.Filter = "Excel files|*.xls"
            SaveFileDialog1.FileName = ""
            SaveFileDialog1.AddExtension = True
            If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
                C_FileName = SaveFileDialog1.FileName
                

                ex = New Excel.Application
                c_WorkBook = ex.Workbooks.Add()
                c_Sheet = c_WorkBook.Worksheets.Add()

                
                c_Sheet.Cells.Range("A1").Value = "WP"
                c_Sheet.Cells.Range("B1").Value = "Name"
                c_Sheet.Cells.Range("C1").Value = "Lat"
                c_Sheet.Cells.Range("D1").Value = "Lon"
                c_Sheet.Cells.Range("E1").Value = "Distance"
                c_Sheet.Cells.Range("F1").Value = "Course"
                c_Sheet.Cells.Range("a1:f1").Font.Bold = True
                c_Sheet.Cells.Range("a1:f500").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                For c_Count = 0 To ListBox1.Items.Count - 1
                    c_Str = Split(ListBox1.Items.Item(c_Count), vbTab)
                    c_Sheet.Cells.Range("A" & c_Count + 2).Value = c_Str(0)
                    c_Sheet.Cells.Range("B" & c_Count + 2).Value = c_Str(1)
                    c_Sheet.Cells.Range("C" & c_Count + 2).Value = c_Str(2)
                    c_Sheet.Cells.Range("D" & c_Count + 2).Value = c_Str(3)
                    If c_Count < ListBox1.Items.Count - 1 Then
                        Dim lat1, lon1, lat2, lon2 As String
                        lat1 = c_Str(2)
                        lon1 = c_Str(3)
                        c_Str = Split(ListBox1.Items.Item(c_Count + 1), vbTab)
                        lat2 = c_Str(2)
                        lon2 = c_Str(3)
                        c_Sheet.Cells.Range("F" & c_Count + 2).Value = Round(GetCourse(LatLonConvert(lat1), LatLonConvert(lon1), LatLonConvert(lat2), LatLonConvert(lon2)), 1)
                        c_Sheet.Cells.Range("E" & c_Count + 2).Value = Round(c_Distance, 1)
                    End If
                Next
                c_Sheet.Columns("A:F").AutoFit()
                c_WorkBook.SaveAs(C_FileName)
                ex.Workbooks.Close()
                ex.Quit()

                MsgBox("File '" & C_FileName & "' saved.", vbInformation)
                '
                '
            Else
                MsgBox("File not saved.", vbInformation)
            End If
            Exit Sub
            Else
            MsgBox("Please select export format: 'eGlobe G2', 'Transas 3000/4000', 'BonVoyage' or 'EXCEL'!", vbInformation)
                Exit Sub
            End If
        
        SaveFileDialog1.FileName = ""
        SaveFileDialog1.AddExtension = True
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            C_FileName = SaveFileDialog1.FileName
            c_Str = Split(C_FileName, "\")
            c_FN = c_Str(UBound(c_Str))
            c_FN = Strings.Left(c_FN, c_FN.Length - 4)
            c_TXT = Replace(c_TXT, "?????", c_FN)
            FileOpen(1, C_FileName, OpenMode.Output)
            PrintLine(1, c_TXT)
            FileClose(1)
            MsgBox("File '" & C_FileName & "' saved.", vbInformation)
        Else
            MsgBox("File not saved.", vbInformation)
        End If

1:
        Select Case Err.Number
            Case 0
            Case 13
                If MsgBox("Error exporting file" & vbCrLf & "Wrong settings for Excel detected" & vbCrLf & "Do you want to fix it?", vbQuestion + vbYesNo, "Error...") = DialogResult.Yes Then
                    ExcelFix()
                End If
            Case Else
                MsgBox("File not saved." & vbCrLf & "Error " & Err.Number & " - " & Err.Description, vbExclamation, "Error saving file")
        End Select
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        Button3.Text = "Export to EXCEL"
    End Sub


    Private Sub ListBox1_DragDrop1(sender As Object, e As DragEventArgs) Handles ListBox1.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        'For Each path In files
        'MsgBox(path)
        'Next
        Open_File(files(0))
    End Sub

    Private Sub ListBox1_DragEnter(sender As Object, e As DragEventArgs) Handles ListBox1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub
    Private Sub Open_File(cFileName As String)
        If FileLen(cFileName) < 200000 Then
            If CheckFileType(cFileName) <> "" Then
                Label1.Text = "File: " & cFileName
                Label2.Text = "File type: " & c_FileType
                ListBox1.Items.Clear()
                GetWayPoints(cFileName, CheckFileType(cFileName))
                If ListBox1.Items.Count <> 0 Then MsgBox("File '" & cFileName & "' loaded.", vbInformation)
                Button3.Enabled = True
            Else
                Label1.Text = "File: " & cFileName
                Label2.Text = "File type: " & c_FileType
                Button3.Enabled = False
            End If
        Else
            Label1.Text = "File: " & cFileName
            Label2.Text = "UNKNOWN FILE TYPE OR FILE VERSION"
            MsgBox("File is too big." & vbCrLf & "Please select another file.", vbInformation, "File opening error")
            Button3.Enabled = False
            Exit Sub
        End If
        If ListBox1.Items.Count = 0 Then
            Label1.Text = "File: " & cFileName
            Label2.Text = "UNKNOWN FILE TYPE OR FILE VERSION"
            MsgBox("Wrong file selected or file loading eror.", vbInformation)
            Button3.Enabled = False
        End If

    End Sub
    Private Sub ExcelFix()
        Try
            Dim key As RegistryKey = Registry.ClassesRoot.OpenSubKey("TypeLib\{00020813-0000-0000-C000-000000000046}", True)
            Dim skey As RegistryKey
            For Each subkey In key.GetSubKeyNames
                ListBox1.Items.Add(subkey.ToString())
                skey = My.Computer.Registry.ClassesRoot.OpenSubKey("TypeLib\{00020813-0000-0000-C000-000000000046}\" & subkey.ToString())
                If skey.SubKeyCount = 0 Then
                    key.DeleteSubKey(subkey, True)
                    cFix = True
                End If
                skey.Close()
            Next
            key.Close()
        Catch key As Exception
            MsgBox(key.Message & vbCrLf & vbCrLf & "Program must run as Admistrator", vbInformation)
            Exit Sub
        End Try
        If cFix = True Then
            MsgBox("Defective setting for Excel fixed", vbInformation, "Successfully fixed")
        Else
            MsgBox("Defective setting cannot be fixed", vbExclamation)
        End If
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        Button3.Text = "Export to BonVoyage: .BVS"
    End Sub
End Class
