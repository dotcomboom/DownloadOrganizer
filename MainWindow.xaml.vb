Class MainWindow
    Private Shared EmptyDelegate As Action = Function()
                                             End Function
    Private Function Organize(real As Boolean, directory As String)
        PreviewBtn.IsEnabled = False
        StartBtn.IsEnabled = False
        Dim pfiles = 0
        Dim files = 0
        Dim ignored = 0
        Dim moved = 0
        Dim starttime = TimeOfDay.Ticks
        'These are definitions of temporary/partial files
        'It will also be noted that if there is a file with the same name it will be ignored
        Dim tempdef = {"!ut", "crdownload", "opdownload", "part", "partial", "temp", "tmp", "lck"}


        Log.Items.Clear()

        Progress.Value = 0
        If Not real Then
            Log.Items.Add("(PREVIEW)")
            Log.Items.Add("")
        End If
        Log.Items.Add("Counting files")
        Log.Items.Add("")
        For Each file As String In My.Computer.FileSystem.GetFiles(directory)
            pfiles = pfiles + 1
        Next
        Progress.Maximum = pfiles

        For Each file As String In My.Computer.FileSystem.GetFiles(directory)
            Progress.Value += 1
            Dispatcher.Invoke(Threading.DispatcherPriority.Render, EmptyDelegate)
            Progress.UpdateLayout()
            files = files + 1
            Dim halt = False
            If IO.File.GetAttributes(file).ToString.Contains("Hidden") Then
                ignored += 1
                Log.Items.Add("Ignored hidden file " & file)
            ElseIf IO.File.GetAttributes(file).ToString.Contains("System") Then
                ignored += 1
                Log.Items.Add("Ignored system file " & file)
            ElseIf IO.File.GetAttributes(file).ToString.Contains("ReadOnly") Then
                ignored += 1
                Log.Items.Add("Ignored read-only file " & file)
            ElseIf tempdef.Contains(IO.Path.GetExtension(file).Replace(".", "")) Then
                ignored += 1
                Log.Items.Add("Ignored temp/partial file " & file)
            Else
                'This for routine checks if the file has a temp file associated with it
                'This is because some programs like Firefox will have the file then
                ' a .part file with the same name.
                For Each ext As String In tempdef
                    If My.Computer.FileSystem.FileExists(file & "." & ext) Then
                        halt = True
                    End If
                Next
                If halt Then
                    ignored += 1
                    Log.Items.Add("Ignored temp/partial file " & file)
                Else
                    moved += 1
                    Dim newdir = directory & "\" & IO.Path.GetExtension(file).Replace(".", "")
                    Dim newpath = newdir & "\" & IO.Path.GetFileName(file)
                    If real Then
                        If My.Computer.FileSystem.DirectoryExists(newdir) Then
                            My.Computer.FileSystem.MoveFile(file, newpath)
                        Else
                            My.Computer.FileSystem.CreateDirectory(newdir)
                            My.Computer.FileSystem.MoveFile(file, newpath)
                        End If
                    End If
                    Log.Items.Add("Moved file " & file & " to " & newpath)
                End If
            End If
        Next

        Log.Items.Add("")
        Log.Items.Add("-- Summary --")
        Log.Items.Add(moved & " files moved")
        Log.Items.Add(ignored & " files ignored")
        'Log.Items.Add("Process completed in " & (TimeOfDay.Ticks / 10000) - (starttime / 10000) & "ms")
        If Not real Then
            Log.Items.Add("")
            Log.Items.Add("(PREVIEW)")
        End If
        Progress.Value = 0
        PreviewBtn.IsEnabled = True
        StartBtn.IsEnabled = True
        Return True
    End Function

    Private Sub PreviewBtn_Click(sender As Object, e As RoutedEventArgs) Handles PreviewBtn.Click
        Organize(False, My.Computer.FileSystem.SpecialDirectories.Desktop.Replace("Desktop", "Downloads"))
    End Sub

    Private Sub StartBtn_Click(sender As Object, e As RoutedEventArgs) Handles StartBtn.Click
        Organize(True, My.Computer.FileSystem.SpecialDirectories.Desktop.Replace("Desktop", "Downloads"))
    End Sub

    Private Sub Progress_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles Progress.ValueChanged
        If Progress.Value = 0 Then
            Progress.Visibility = Visibility.Hidden
        Else
            Progress.Visibility = Visibility.Visible
        End If
    End Sub
End Class
