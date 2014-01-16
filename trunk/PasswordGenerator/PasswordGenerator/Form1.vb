Public Class Form1

#Region "Business Variables"
    Private OpenName As String
    Private saveSW As IO.StreamWriter
    Private saveNewSW As IO.StreamWriter
    Private SaveName As String
    Private SaveNewFileName As String
    Dim t As Threading.Thread
    Dim tNewFile As Threading.Thread
    Dim Lines() As String
    Dim FullCount As UInt64 = 0
    Dim count As UInt64 = 0
    Dim lastperc As Integer = -1
#End Region

#Region "Delegate & Event"
    Delegate Sub UpdateStatDel(ByVal Percent As Integer)
    Dim Invoker As New UpdateStatDel(AddressOf UpdateStat)
    Dim InvokerNewFile As New UpdateStatDel(AddressOf UpdateStatNewFile)
#End Region

#Region "CheckBox Check Changed Event"

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

    Private Sub chkOnlyAlphabetic_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOnlyAlphabetic.CheckedChanged
        cmbAplhabetic.Enabled = chkOnlyAlphabetic.Checked
        MTextBox2.Enabled = chkOnlyAlphabetic.Checked
        If cmbAplhabetic.SelectedIndex = -1 Then cmbAplhabetic.SelectedIndex = 0
    End Sub

    Private Sub chkOnlyNumeric_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOnlyNumeric.CheckedChanged
        cmbNumerical.Enabled = chkOnlyNumeric.Checked
        MTextBox3.Enabled = chkOnlyNumeric.Checked
        If cmbNumerical.SelectedIndex = -1 Then cmbNumerical.SelectedIndex = 0
    End Sub

    Private Sub chkAlphanumeric_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAlphanumeric.CheckedChanged
        cmbAlphanumericalLetter.Enabled = chkAlphanumeric.Checked
        cmbAlphanumericalDigit.Enabled = chkAlphanumeric.Checked
        MTextBox4.Enabled = chkAlphanumeric.Checked
        MTextBox5.Enabled = chkAlphanumeric.Checked
        If cmbAlphanumericalLetter.SelectedIndex = -1 Then cmbAlphanumericalLetter.SelectedIndex = 0
        If cmbAlphanumericalDigit.SelectedIndex = -1 Then cmbAlphanumericalDigit.SelectedIndex = 0
    End Sub

#End Region

#Region "Form Load and Closing Event"

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If saveSW IsNot Nothing Then
            saveSW.Dispose()
        End If
        If saveNewSW IsNot Nothing Then
            saveNewSW.Dispose()
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CheckBox1_CheckedChanged(Nothing, Nothing)
        CheckBox2_CheckedChanged(Nothing, Nothing)
        CheckBox3_CheckedChanged(Nothing, Nothing)
        CheckBox5_CheckedChanged(Nothing, Nothing)
        chkOnlyAlphabetic_CheckedChanged(Nothing, Nothing)
        chkOnlyNumeric_CheckedChanged(Nothing, Nothing)
        chkAlphanumeric_CheckedChanged(Nothing, Nothing)
    End Sub

#End Region

#Region "Button Click Events"

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

                'If chkOnlyAlphabetic.Checked Then
                '    FullCount *= 26 ^ (cmbAplhabetic.SelectedIndex + 1)
                'End If

                'If chkOnlyAlphabetic.Checked Then
                '    FullCount *= 26 ^ (cmbNumerical.SelectedIndex + 1)
                'End If

                count = 0
                t = New Threading.Thread(New Threading.ParameterizedThreadStart(AddressOf DoWork))
                t.IsBackground = True
                t.Start(New Object() {If(CheckBox1.Checked, ComboBox1.SelectedIndex, -1), If(CheckBox2.Checked, ComboBox2.SelectedIndex, -1), If(CheckBox3.Checked, ComboBox3.SelectedIndex, -1), CheckBox4.Checked, If(CheckBox5.Checked, MTextBox1.Text, "")}) ', If(chkOnlyAlphabetic.Checked, cmbAplhabetic.SelectedIndex - 1, -1), If(chkOnlyNumeric.Checked, cmbNumerical.SelectedIndex - 1, -1)})
            Else
                MsgBox("The File Is Empty", MsgBoxStyle.Critical, "No Data")
            End If

        Else
            MsgBox("Select at least one operation", MsgBoxStyle.Critical, "No Operation Selected")
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ofd As New OpenFileDialog
        ofd.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
            OpenName = ofd.FileName
            UpdateStat()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
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

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If t IsNot Nothing AndAlso t.IsAlive Then
            If MsgBox("Work In Progress, Would you like to abort?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, "Exit?") = MsgBoxResult.Yes Then
                Me.Close()
            End If
        Else
            Me.Close()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        MsgBox("open an existing text file with usernames and select the according strings to add. F.ex. a user could be peter. To create a user peter 33 append 2 digits numbers.")
    End Sub

    Private Sub btnCreateNewSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateNewSave.Click
        Dim sfd As New SaveFileDialog
        sfd.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        If sfd.ShowDialog = Windows.Forms.DialogResult.OK Then
            Try
                If saveNewSW IsNot Nothing Then
                    saveNewSW.Dispose()
                    saveNewSW = Nothing
                End If
                saveNewSW = New IO.StreamWriter(sfd.FileName, True)
                SaveNewFileName = sfd.FileName
                UpdateStatNewFile()
            Catch ex As Exception
                saveNewSW = Nothing
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error Open File To Write")
            End Try
        End If
    End Sub

    Private Sub btnCreateNewFileStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateNewFileStart.Click
        If SaveNewFileName = "" Then
            MsgBox("Select a file to save", MsgBoxStyle.Critical, "No File Selected")
            Exit Sub
        End If

        If chkOnlyAlphabetic.Checked OrElse chkOnlyNumeric.Checked OrElse chkAlphanumeric.Checked Then
            If tNewFile IsNot Nothing AndAlso tNewFile.IsAlive Then
                Try
                    If MsgBox("Work In Progress, Would you like to abort?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, "Abort?") = MsgBoxResult.Yes Then
                        tNewFile.Abort()
                    Else
                        Exit Sub
                    End If
                Catch
                End Try
                Try
                    saveNewSW.Close()
                Catch
                End Try
            End If
            Try
                If Not (saveNewSW IsNot Nothing AndAlso saveNewSW.BaseStream.CanWrite) Then
                    saveNewSW = New IO.StreamWriter(SaveNewFileName, True)
                End If
            Catch ex As Exception
                MsgBox("Cannot open file to write : " & SaveNewFileName & vbNewLine & ex.Message, MsgBoxStyle.Critical, "Error Save File")
                Exit Sub
            End Try
            tNewFile = New Threading.Thread(New Threading.ParameterizedThreadStart(AddressOf DoWorkCreateNewFile))
            tNewFile.IsBackground = True
            tNewFile.Start(New Object() {If(chkOnlyAlphabetic.Checked, cmbAplhabetic.SelectedIndex, -1), If(chkOnlyNumeric.Checked, cmbNumerical.SelectedIndex, -1), If(chkAlphanumeric.Checked, cmbAlphanumericalLetter.SelectedIndex, -1), If(chkAlphanumeric.Checked, cmbAlphanumericalDigit.SelectedIndex, -1), MTextBox2.Text, MTextBox3.Text, MTextBox4.Text, MTextBox5.Text})

        Else
            MsgBox("Select at least one operation", MsgBoxStyle.Critical, "No Operation Selected")
        End If
    End Sub

    Private Sub btnCreateNewFileExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateNewFileExit.Click
        If tNewFile IsNot Nothing AndAlso tNewFile.IsAlive Then
            If MsgBox("Work In Progress, Would you like to abort?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, "Exit?") = MsgBoxResult.Yes Then
                Me.Close()
            End If
        Else
            Me.Close()
        End If
    End Sub

#End Region

#Region "Private Procedures"

    Sub DoWork(ByVal obj() As Object)
        Try
            Dim Appender As Integer = obj(0)
            Dim Prepender As Integer = obj(1)
            Dim Replacemet As Integer = obj(2)
            Dim Special As Boolean = obj(3)
            Dim Domain As String = obj(4)
            'Dim LetterAppender As Integer = obj(5)
            'Dim LetterPrepender As Integer = obj(6)

            If Appender <> -1 Then Appender = 10 ^ (Appender + 1)
            If Prepender <> -1 Then Prepender = 10 ^ (Prepender + 1)
            'If LetterAppender <> -1 Then LetterAppender = 26 ^ (LetterAppender + 1)
            'If LetterPrepender <> -1 Then LetterPrepender = 26 ^ (LetterPrepender + 1)

            For i = 0 To Lines.Length - 1
                append(Lines(i), Appender, Prepender, Replacemet, Special, Domain) ', LetterAppender, LetterPrepender)
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

    Sub append(ByVal s As String, ByVal Appender As Integer, ByVal prepender As Integer, ByVal replacement As Integer, ByVal spec As Boolean, ByVal Domain As String) ', ByVal LetterAppender As Integer, ByVal LetterPrepender As Integer)
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

    Sub DoWorkCreateNewFile(ByVal obj() As Object)
        Try
            Dim onlyAlphabets As Integer = obj(0)
            Dim selectedAlphabets As String = obj(4)

            Dim onlyNumerical As Integer = obj(1)
            Dim selectedNumber As String = obj(5)

            Dim alphanumericalLetter As Integer = obj(2)
            Dim alphanumericalDigit As Integer = obj(3)
            Dim selectedAlphanumericLetter As String = obj(6)
            Dim selectedAlphanumericNumber As String = obj(7)

            Select Case onlyAlphabets
                Case 0
                    onlyAlphabets = 2
                Case 1
                    onlyAlphabets = 3
            End Select

            If onlyAlphabets > 0 Then
                Dim letters As List(Of String) = GenerateLetterCombination(onlyAlphabets, selectedAlphabets)
                For Each Str As String In letters
                    saveNewSW.WriteLine(Str)
                Next
            End If

            Select Case onlyNumerical
                Case 0
                    onlyNumerical = 2
                Case 1
                    onlyNumerical = 3
                Case 2
                    onlyNumerical = 4
            End Select

            If onlyNumerical > 0 Then
                Dim numbers As List(Of String) = GenerateNumericCombination(onlyNumerical, selectedNumber)
                For Each Str As String In numbers
                    saveNewSW.WriteLine(Str)
                Next
            End If

            '''' If Alphanumeric Option is seleted''''''''''''''''''''''''''''''''''''''''''''''''''''
            If chkAlphanumeric.Checked Then
                Dim letters As List(Of String) = New List(Of String)
                Dim numbers As List(Of String) = New List(Of String)
                Select Case alphanumericalLetter
                    Case 0
                        alphanumericalLetter = 2
                    Case 1
                        alphanumericalLetter = 3
                End Select

                If alphanumericalLetter > 0 Then
                    letters = GenerateLetterCombination(alphanumericalLetter, selectedAlphanumericLetter)
                End If

                Select Case alphanumericalDigit
                    Case 0
                        alphanumericalDigit = 2
                    Case 1
                        alphanumericalDigit = 3
                    Case 2
                        alphanumericalDigit = 4
                End Select

                If alphanumericalDigit > 0 Then
                    numbers = GenerateNumericCombination(alphanumericalDigit, selectedAlphanumericNumber)
                End If

                For Each letter As String In letters
                    For Each number As String In numbers
                        saveNewSW.WriteLine(letter + number)
                    Next
                Next
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Me.Invoke(InvokerNewFile, 100)
            MsgBox("Done!")

        Catch ex As Exception

        End Try
        Try
            saveNewSW.Close()
            saveNewSW.Dispose()
        Catch
        End Try
        saveNewSW = Nothing
    End Sub

    Function GenerateLetterCombination(ByVal numberOfLetters As Integer, ByVal selectedLetter As String) As List(Of String)
        Dim selectedLetters() As String = New String() {}
        Dim letter As List(Of String) = New List(Of String)
        If selectedLetter.Length > 0 Then
            selectedLetters = Split(selectedLetter, ",")
        End If

        If selectedLetters.Length > 0 Then
            For Each i As String In selectedLetters
                Dim l As Char = Convert.ToChar(i)
                letter.Add(l.ToString())
            Next
            For index As Integer = 1 To numberOfLetters - 1
                Dim newLetter As List(Of String) = New List(Of String)
                For Each s As String In letter
                    For Each i As String In selectedLetters
                        Dim l As Char = Convert.ToChar(i)
                        newLetter.Add(s + l)
                    Next
                Next
                letter = newLetter
            Next
        Else
            For i As Int16 = Convert.ToInt16("a"c) To Convert.ToInt16("z"c)
                Dim l As Char = Convert.ToChar(i)
                letter.Add(l.ToString())
            Next

            For index As Integer = 1 To numberOfLetters - 1
                Dim newLetter As List(Of String) = New List(Of String)
                For Each s As String In letter
                    For i As Int16 = Convert.ToInt16("a"c) To Convert.ToInt16("z"c)
                        Dim l As Char = Convert.ToChar(i)
                        newLetter.Add(s + l)
                    Next

                Next
                letter = newLetter
            Next
        End If
        Return letter
        'For Each Str As String In letter
        '    saveNewSW.WriteLine(Str)
        'Next
        'Me.Invoke(InvokerNewFile, 100)
        'saveNewSW.Dispose()
        'saveNewSW.Flush()

    End Function

    Function GenerateNumericCombination(ByVal numberOfLetters As Integer, ByVal selectedNumber As String) As List(Of String)

        Dim letter As List(Of String) = New List(Of String)
        Dim selectedNumbers() As String = New String() {}
        If selectedNumber.Length > 0 Then
            selectedNumbers = Split(selectedNumber, ",")
        End If
        If selectedNumbers.Length > 0 Then
            For Each i As String In selectedNumbers
                letter.Add(i.ToString())
            Next

            For index As Integer = 1 To numberOfLetters - 1
                Dim newLetter As List(Of String) = New List(Of String)
                For Each s As String In letter
                    For Each i As String In selectedNumbers
                        newLetter.Add(s + i.ToString())
                    Next

                Next
                letter = newLetter
            Next
        Else
            For i As Int16 = 0 To 9
                letter.Add(i.ToString())
            Next

            For index As Integer = 1 To numberOfLetters - 1
                Dim newLetter As List(Of String) = New List(Of String)
                For Each s As String In letter
                    For i As Int16 = 0 To 9
                        newLetter.Add(s + i.ToString())
                    Next

                Next
                letter = newLetter
            Next
        End If
        Return letter

        'For Each Str As String In letter
        '    saveNewSW.WriteLine(Str)
        'Next

        'Me.Invoke(InvokerNewFile, 100)
        'saveNewSW.Dispose()
        'saveNewSW.Flush()

    End Function

    Sub UpdateStatNewFile(Optional ByVal Percent As Integer = -1)
        Label4.Text = ""
        If SaveNewFileName <> "" Then
            Label4.Text &= "Destination : " & IO.Path.GetFileName(SaveNewFileName) & vbNewLine
        End If
        If Percent >= 0 Then
            If Percent = 100 Then
                Label4.Text &= "Done"
            Else
                Label4.Text &= "Work in progress : " & Percent
            End If
        End If
    End Sub

#End Region

End Class
