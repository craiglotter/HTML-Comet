Imports System.IO
Imports System.Globalization

Public Class Main_Screen
    Dim titlecase As Boolean = True
    Private WithEvents PD As Page_Display

    Private Const WM_NCHITTEST As Integer = &H84
    Private Const HTCLIENT As Integer = &H1
    Private Const HTCAPTION As Integer = &H2

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Select Case m.Msg
            Case WM_NCHITTEST
                MyBase.WndProc(m)
                If (m.Result.ToInt32 = HTCLIENT) Then
                    m.Result = IntPtr.op_Explicit(HTCAPTION)
                End If
                Exit Sub
        End Select

        MyBase.WndProc(m)
    End Sub


    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If My.Computer.FileSystem.FileExists((Application.StartupPath & "\Sounds\UHOH.WAV").Replace("\\", "\")) = True Then
                My.Computer.Audio.Play((Application.StartupPath & "\Sounds\UHOH.WAV").Replace("\\", "\"), AudioPlayMode.Background)
            End If
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message()
                Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ":" & ex.Message.ToString
                Display_Message1.Timer1.Interval = 1000
                Display_Message1.ShowDialog()
                Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                dir = Nothing
                Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ":" & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
                filewriter = Nothing
            End If
        Catch exc As Exception
            MsgBox("An error occurred in the application's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub


    Private Sub Activity_Handler(ByVal Message As String)
        Try
            Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            dir = Nothing
            Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & Message)
            filewriter.Flush()
            filewriter.Close()
            filewriter = Nothing
        Catch ex As Exception
            Error_Handler(ex, "Activity_Logger")
        End Try
    End Sub

    Private Sub ListBox1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox1.DragEnter, Me.DragEnter
        Try
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                e.Effect = DragDropEffects.All
            End If
        Catch ex As Exception
            Error_Handler(ex, "ListBox1_DragEnter")
        End Try
    End Sub

    Private Sub ListBox1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox1.DragDrop, Me.DragDrop
        Try
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                Dim myfiles() As String
                Dim i As Integer
                ListBox1.Items.Clear()
                ListBox1.Sorted = True
                myfiles = e.Data.GetData(DataFormats.FileDrop)
                For i = 0 To myfiles.Length - 1
                    If myfiles.Length > 0 Then
                        Dim finfo As FileInfo = New FileInfo(myfiles(i))
                        If finfo.Exists = True Then
                            ListBox1.Items.Add(myfiles(i))
                        End If
                        Dim dinfo As DirectoryInfo = New DirectoryInfo(myfiles(i))
                        If dinfo.Exists = True Then
                            ListBox1.Items.Add(myfiles(i))
                        End If
                        finfo = Nothing
                        dinfo = Nothing
                    End If
                Next
                Show_Results()
            End If
        Catch ex As Exception
            Error_Handler(ex, "ListBox1_DragDrop")
        End Try
    End Sub

    Private Sub Show_Results()
        Try
            Dim URL As String = (Application.StartupPath & "\").Replace("\\", "\") & "Display.htm"
            If My.Computer.FileSystem.FileExists(URL) = True Then
                My.Computer.FileSystem.DeleteFile(URL, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter(URL, True)
            filewriter.WriteLine("<html><body>")
            filewriter.WriteLine("<h3>Files</h3>")
            filewriter.WriteLine("<p><b>Files:</b></p>")
            filewriter.WriteLine("<ul>")
            For Each item As String In ListBox1.Items
                Dim finfo As FileInfo = New FileInfo(item)
                If finfo.Exists() = True Then
                    Dim cinfo As CultureInfo = New CultureInfo("en-US")
                    If titlecase = True Then
                        filewriter.WriteLine("<li>" & cinfo.TextInfo.ToTitleCase(UpperCourseCode(finfo.Name.Remove(finfo.Name.Length - finfo.Extension.Length, finfo.Extension.Length).ToLower)) & " - Click <a target=""_blank"" href=""" & finfo.Name & """>here</a> <font size=""1"" color=""#008080"">(" & finfo.Extension.ToUpper & ")</font>" & filesize(finfo.Length) & "<font size=""1"" color=""#808080""> (" & Format(Now(), "dd/MM/yyyy") & ") </font><font size=""1"" color=""#FF0000""> <i>NEW</i></font></li>")
                    Else
                        filewriter.WriteLine("<li>" & cinfo.TextInfo.ToTitleCase(UpperCourseCode(finfo.Name.Remove(finfo.Name.Length - finfo.Extension.Length, finfo.Extension.Length))) & " - Click <a target=""_blank"" href=""" & finfo.Name & """>here</a> <font size=""1"" color=""#008080"">(" & finfo.Extension.ToUpper & ")</font>" & filesize(finfo.Length) & "<font size=""1"" color=""#808080""> (" & Format(Now(), "dd/MM/yyyy") & ") </font><font size=""1"" color=""#FF0000""> <i>NEW</i></font></li>")
                    End If

                    cinfo = Nothing
                Else
                    Dim dinfo As DirectoryInfo = New DirectoryInfo(item)
                    If dinfo.Exists = True Then
                        FolderWalker("", dinfo.FullName, filewriter)
                        filewriter.WriteLine("</ul><p>&nbsp;</p><ul>")
                    End If
                    dinfo = Nothing
                End If



            Next
            filewriter.WriteLine("</ul>")
            filewriter.WriteLine("</body></html>")

            filewriter.Flush()
            filewriter.Close()
            filewriter = Nothing


            ShowPage(URL)
            titlecase = True
        Catch ex As Exception
            Error_Handler(ex, "Show_Results")
        End Try
    End Sub

    Private Sub FolderWalker(ByVal rootdir As String, ByVal targetdir As String, ByVal filewriter As StreamWriter)
        Try
            Dim dinfo As DirectoryInfo = New DirectoryInfo(targetdir)
            Dim finfo As FileInfo
            rootdir = (rootdir & "/" & dinfo.Name).Replace("//", "/")
            filewriter.WriteLine("</ul>" & vbCrLf & "<p><b>" & UpperCourseCode(rootdir.Remove(0, 1)) & "</b></p>" & vbCrLf & "<ul>" & vbCrLf)
            For Each finfo In dinfo.GetFiles()
                Dim cinfo As CultureInfo = New CultureInfo("en-US")
                If titlecase = True Then
                    filewriter.WriteLine("<li>" & cinfo.TextInfo.ToTitleCase(UpperCourseCode(finfo.Name.Remove(finfo.Name.Length - finfo.Extension.Length, finfo.Extension.Length).ToLower)) & " - Click <a target=""_blank"" href=""" & (rootdir & "/" & finfo.Name).Replace("//", "/") & """>here</a> <font size=""1"" color=""#008080"">(" & finfo.Extension.ToUpper & ")</font>" & filesize(finfo.Length) & "<font size=""1"" color=""#808080""> (" & Format(Now(), "dd/MM/yyyy") & ") </font><font size=""1"" color=""#FF0000""> <i>NEW</i></font></li>")
                Else
                    filewriter.WriteLine("<li>" & cinfo.TextInfo.ToTitleCase(UpperCourseCode(finfo.Name.Remove(finfo.Name.Length - finfo.Extension.Length, finfo.Extension.Length))) & " - Click <a target=""_blank"" href=""" & (rootdir & "/" & finfo.Name).Replace("//", "/") & """>here</a> <font size=""1"" color=""#008080"">(" & finfo.Extension.ToUpper & ")</font>" & filesize(finfo.Length) & "<font size=""1"" color=""#808080""> (" & Format(Now(), "dd/MM/yyyy") & ") </font><font size=""1"" color=""#FF0000""> <i>NEW</i></font></li>")
                End If

                cinfo = Nothing
            Next
            finfo = Nothing
            Dim d2info As DirectoryInfo
            For Each d2info In dinfo.GetDirectories()
                FolderWalker(rootdir, d2info.FullName, filewriter)
            Next
            d2info = Nothing
        Catch ex As Exception
            Error_Handler(ex, "FolderWalker")
        End Try
    End Sub

    Private Sub ShowPage(byval URL as string)

        '  If the instance still exists... (ie. it's Not Nothing)
        If Not IsNothing(PD) Then
            ' MsgBox("If the instance still exists... (ie. it's Not Nothing)")
            '  and if it hasn't been disposed yet
            If Not PD.IsDisposed Then
                '    MsgBox("then it must already be instantiated - maybe it's minimized or hidden behind other forms ?")
                '  then it must already be instantiated - maybe it's
                '  minimized or hidden behind other forms ?
                PD.WindowState = FormWindowState.Normal  ' Optional
                PD.BringToFront()  '  Optional
                PD.WebBrowser1.Navigate(URL)
            Else
                '  MsgBox("else it has already been disposed, so you can instantiate a new form and show it")
                '  else it has already been disposed, so you can
                '  instantiate a new form and show it
                PD = New Page_Display()
                PD.WebBrowser1.Navigate(URL)
                PD.Show()
                PD.WindowState = FormWindowState.Normal
                PD.BringToFront()
                PD.Refresh()
            End If
        Else
            '  MsgBox("else the form = nothing, so you can safely instantiate a new form and show it")
            '  else the form = nothing, so you can safely
            '  instantiate a new form and show it
            PD = New Page_Display()
            PD.WebBrowser1.Navigate(URL)
            PD.Show()
            PD.WindowState = FormWindowState.Normal
            PD.BringToFront()
            PD.Refresh()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim URL As String = (Application.StartupPath & "\").Replace("\\", "\") & "Display.htm"
            If My.Computer.FileSystem.FileExists(URL) = True Then
                My.Computer.FileSystem.DeleteFile(URL, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter(URL, True)
            filewriter.WriteLine("<html><body>")
            filewriter.WriteLine("<ul>")

            filewriter.WriteLine("<li><b>" & Format(Now(), "dd/MM/yyyy") & " - Heading <font color=""#808080"">(Lecturer) </font><font color=""#FF0000""><i>NEW</i></font></b><br><br>message<br></li>")


            filewriter.WriteLine("</ul>")
            filewriter.WriteLine("</body></html>")

            filewriter.Flush()
            filewriter.Close()
            filewriter = Nothing

            ShowPage(URL)

        Catch ex As Exception
            Error_Handler(ex, "Show_Results")
        End Try
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        titlecase = False
        Show_Results()
    End Sub

    Private Sub mainscreen_windowstatechange(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.SizeChanged
        Try
            Me.WindowState = FormWindowState.Normal
        Catch ex As Exception

        End Try
    End Sub

    Private Sub mainscreen_doubleclick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDoubleClick
        Show_Results()
    End Sub

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'PD = New Page_Display
        If My.Application.CommandLineArgs.Count > 0 Then
            CommandLineArgs_ShowResults()
        End If
    End Sub

    Private Sub CommandLineArgs_ShowResults()
        Try

            Dim i As Integer
            ListBox1.Items.Clear()
            ListBox1.Sorted = True

            For Each fil As String In My.Application.CommandLineArgs
                If fil.Length > 0 Then
                    Dim finfo As FileInfo = New FileInfo(fil)
                    If finfo.Exists = True Then
                        ListBox1.Items.Add(finfo.FullName)
                    End If
                    Dim dinfo As DirectoryInfo = New DirectoryInfo(fil)
                    If dinfo.Exists = True Then
                        ListBox1.Items.Add(dinfo.FullName)
                    End If
                    finfo = Nothing
                    dinfo = Nothing
                End If
            Next

            Show_Results()
        Catch ex As Exception
            Error_Handler(ex, "CommandLineArgs_ShowResults")
        End Try
    End Sub

    Private Function UpperCourseCode(ByVal input As String) As String
        Dim result As String = ""
        Dim res As String() = input.Split(" ")
        For Each token As String In res
            If token.Length = 8 Then
                If IsNumeric(token.Substring(3, 4)) = True Then
                    If Not IsNumeric(token.Substring(0, 3)) Then
                        If Not IsNumeric(token.Substring(7, 1)) Then
                            token = token.ToUpper
                        End If
                    End If
                End If
            End If
            result = result & " " & token
        Next
        Return result
    End Function

    Private Function filesize(ByVal size As Long) As String
        Try
            Dim result As String
            Dim switch As String = "bytes"
            If size >= 1024 Then switch = "KB"
            If size >= 1048576 Then switch = "MB"
            If size >= 1073741824 Then switch = "GB"

            Select Case switch
                Case "KB"
                    result = Math.Round((size / 1024), 1).ToString & " KB"
                Case "MB"
                    result = Math.Round((size / 1048576), 1).ToString & " MB"
                Case "GB"
                    result = Math.Round((size / 1073741824), 1).ToString & " GB"
                Case Else
                    result = size.ToString & " Bytes"
            End Select



            result = "<font size=""1"" color=""#5378C3""> (" & result & ")</font>"
            Return result
        Catch ex As Exception
            Error_Handler(ex, "Displaying File Size")
            Return ""
        End Try
    End Function



    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Try
            Me.Close()
        Catch ex As Exception
            Error_Handler(ex, "Exiting Application")
        End Try
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        titlecase = False
        Show_Results()
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Show_Results()
    End Sub
End Class
