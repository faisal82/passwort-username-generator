Public Class Form1
    Private OpenName As String
    Private saveSW As IO.StreamWriter
    Private SaveName As String
    Dim t As Threading.Thread
    Delegate Sub UpdateStatDel(ByVal Percent As Integer)
    Dim Invoker As New UpdateStatDel(AddressOf UpdateStat)
    Sub UpdateStat(Optional ByVal Percent As Integer = -1)
        Label2.Text = ""
        If OpenName <> "" Then
            Label2.Text &= "Loaded : " & IO.Path.GetFileName(OpenName) & vbNewLine
        End If
        If SaveName <> "" Then
            Label2.Text &= "Destination : " & IO.Path.GetFileName(SaveName) & vbNewLine
        End If
        If Percent >= 0 Then
            If Percent = 100 Then
                Label2.Text &= "Done"
            Else
                Label2.Text &= "Work in progress : " & Percent
            End If
        End If
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, Button8.Click
        Dim ofd As New OpenFileDialog
        ofd.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
            OpenName = ofd.FileName
            UpdateStat()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click, Button7.Click
        Dim sfd As New SaveFileDialog
        sfd.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        If sfd.ShowDialog = Windows.Forms.DialogResult.OK Then
            Try
                If saveSW IsNot Nothing Then
                    saveSW.Dispose()
                    saveSW = Nothing
                End If
                saveSW = New IO.StreamWriter(sfd.FileName, True)
                SaveName = sfd.FileName
                UpdateStat()
            Catch ex As Exception
                saveSW = Nothing
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error Open File To Write")
            End Try
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        ComboBox1.Enabled = CheckBox1.Checked
        If ComboBox1.SelectedIndex = -1 Then ComboBox1.SelectedIndex = 0
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        ComboBox2.Enabled = CheckBox2.Checked
        If ComboBox2.SelectedIndex = -1 Then ComboBox2.SelectedIndex = 0
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        ComboBox3.Enabled = CheckBox3.Checked
        If ComboBox3.SelectedIndex = -1 Then ComboBox3.SelectedIndex = 0
    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        MTextBox1.Enabled = CheckBox5.Checked
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If saveSW IsNot Nothing Then
            saveSW.Dispose()
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CheckBox1_CheckedChanged(Nothing, Nothing)
        CheckBox2_CheckedChanged(Nothing, Nothing)
        CheckBox3_CheckedChanged(Nothing, Nothing)
        CheckBox5_CheckedChanged(Nothing, Nothing)

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If OpenName = "" Then
            MsgBox("Select a file to open", MsgBoxStyle.Critical, "No File Selected")
            Exit Sub
        End If

        If SaveName = "" Then
            MsgBox("Select a file to save", MsgBoxStyle.Critical, "No File Selected")
            Exit Sub
       
        End If
         
        If CheckBox1.Checked OrElse CheckBox2.Checked OrElse CheckBox3.Checked OrElse CheckBox4.Checked OrElse CheckBox5.Checked Then
            If t IsNot Nothing AndAlso t.IsAlive Then
                Try
                    If MsgBox("Work In Progress, Would you like to abort?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, "Abort?") = MsgBoxResult.Yes Then
                        t.Abort()
                    Else
                        Exit Sub
                    End If
                Catch
                End Try
                Try
                    saveSW.Close()
                Catch
                End Try
            End If
            Try
                If Not (saveSW IsNot Nothing AndAlso saveSW.BaseStream.CanWrite) Then
                    saveSW = New IO.StreamWriter(SaveName, True)
                End If
            Catch ex As Exception
                MsgBox("Cannot open file to write : " & SaveName & vbNewLine & ex.Message, MsgBoxStyle.Critical, "Error Save File")
                Exit Sub
            End Try
            Try
                Lines = IO.File.ReadAllLines(OpenName)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error Open File")
                Exit Sub
            End Try
            If Lines.Length > 0 Then
                FullCount = Lines.Length
                If CheckBox1.Checked Then
                    FullCount *= 10 ^ (ComboBox1.SelectedIndex + 1)
                End If
                If CheckBox2.Checked Then
                    FullCount *= 10 ^ (ComboBox2.SelectedIndex + 1)
                End If
                If CheckBox4.Checked Then
                    FullCount *= 3
                End If
                count = 0
                t = New Threading.Thread(New Threading.ParameterizedThreadStart(AddressOf DoWork))
                t.IsBackground = True
                t.Start(New Object() {If(CheckBox1.Checked, ComboBox1.SelectedIndex, -1), If(CheckBox2.Checked, ComboBox2.SelectedIndex, -1), If(CheckBox3.Checked, ComboBox3.SelectedIndex, -1), CheckBox4.Checked, If(CheckBox5.Checked, MTextBox1.Text, "")})
            Else
                MsgBox("The File Is Empty", MsgBoxStyle.Critical, "No Data")
            End If

        Else
            MsgBox("Select at least one operation", MsgBoxStyle.Critical, "No Operation Selected")
        End If
    End Sub

    Dim Lines() As String
    Dim FullCount As UInt64 = 0
    Dim count As UInt64 = 0
    Sub DoWork(ByVal obj() As Object)
        Try
            Dim Appender As Integer = obj(0)
            Dim Prepender As Integer = obj(1)
            Dim Replacemet As Integer = obj(2)
            Dim Special As Boolean = obj(3)
            Dim Domain As String = obj(4)
            If Appender <> -1 Then Appender = 10 ^ (Appender + 1)
            If Prepender <> -1 Then Prepender = 10 ^ (Prepender + 1)
            For i = 0 To Lines.Length - 1
                append(Lines(i), Appender, Prepender, Replacemet, Special, Domain)
            Next

            MsgBox("Done!")
        Catch ex As Exception
            MsgBox("An Error Occured, Operation was interrupted." & vbNewLine & ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
        Try
            saveSW.Close()
            saveSW.Dispose()
        Catch
        End Try
        saveSW = Nothing
    End Sub
    Sub append(ByVal s As String, ByVal Appender As Integer, ByVal prepender As Integer, ByVal replacement As Integer, ByVal spec As Boolean, ByVal Domain As String)
        If Appender = -1 Then
            prepend(s, prepender, replacement, spec, Domain)
        Else
            For a = 0 To Appender - 1
                prepend(s & a, prepender, replacement, spec, Domain)
            Next
        End If
    End Sub
    Sub prepend(ByVal s As String, ByVal prepender As Integer, ByVal replacement As Integer, ByVal spec As Boolean, ByVal Domain As String)
        If prepender = -1 Then
            Specialize(s, replacement, spec, Domain)
        Else
            For p = 0 To prepender - 1
                Specialize(p & s, replacement, spec, Domain)
            Next
        End If
    End Sub
    Dim lastperc As Integer = -1
    Sub Specialize(ByVal s As String, ByVal replacement As Integer, ByVal spec As Boolean, ByVal Domain As String)
        s = replace(s, replacement)
        If Not spec Then
            saveSW.WriteLine(s & Domain)
            count += 1
        Else
            saveSW.WriteLine(s & "!" & Domain)
            saveSW.WriteLine(s & "?" & Domain)
            saveSW.WriteLine(s & "$" & Domain)
            count += 3
        End If

        If lastperc <> CInt((count / FullCount) * 100) Then
            lastperc = CInt((count / FullCount) * 100)
            Me.Invoke(Invoker, lastperc)
            saveSW.Flush()
        End If

    End Sub
    Function replace(ByVal s As String, ByVal Replacement As Integer) As String
       Select Case Replacement
            Case 0
                Return s.Replace("o", "0")
            Case 1
                Return s.Replace("e", "3")
            Case 2
                Return s.Replace("l", "1")
            Case 3
                Return s.Replace("o", "0").Replace("e", "3").Replace("l", "1")
            Case Else
                Return s
        End Select
    End Function

   
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If t IsNot Nothing AndAlso t.IsAlive Then
            If MsgBox("Work In Progress, Would you like to abort?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, "Exit?") = MsgBoxResult.Yes Then
                Me.Close()
            End If
        Else
            Me.Close()
        End If
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click

    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        MsgBox("open an existing text file with usernames and select the according strings to add. F.ex. a user could be peter. To create a user peter 33 append 2 digits numbers.")
    End Sub
End Class
