Imports System
Imports System.IO
Imports System.Text
Imports System.Drawing
Imports System.IO.Ports
Imports System.Windows.Forms

Public Class frmMain
    ' objeto filewriter para log
    Private fileLog As System.IO.StreamWriter
    Private isConected As Boolean = False
    Private urlFileLog As String = ""

    Private Function StringToHexString(ByVal strInput As String) As String
        strInput = "00" & strInput
        Return Strings.Right(strInput, 2)
    End Function

    Private Sub frmMain_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        SerialPort1.Close()
        fileLog.Close()
    End Sub

    Private Sub frmMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'Dim urlFileLog As String
        ' formata onome do arquivo de log para a data atual
        urlFileLog = DateTime.Now.ToString("dd-MM-yyyy") + ".log"
        ' abre o arquivo para gravar
        fileLog = My.Computer.FileSystem.OpenTextFileWriter(urlFileLog, True)
        ' exibe o nome do arquivo de log na parte inferior da tela
        If Strings.Len(My.Computer.FileSystem.CurrentDirectory) > 60 Then
            LinkFileLog.Text = Strings.Left(My.Computer.FileSystem.CurrentDirectory, 60) + "..\" + urlFileLog
        Else
            LinkFileLog.Text = My.Computer.FileSystem.CurrentDirectory + "\" + urlFileLog
        End If
        ' habilita o link para abrir o arquivo de log
        If My.Computer.FileSystem.FileExists(My.Computer.FileSystem.CurrentDirectory + "\" + urlFileLog) Then
            LinkFileLog.Enabled = True
        End If
        'lista as portas seriais disponiveis
        Dim portas = SerialPort.GetPortNames()
        'carrega na combo as portas disponiveis
        For i = 0 To UBound(portas)
            cmbPorta.Items.Add(portas(i))
        Next
        AddHandler SerialPort1.DataReceived, AddressOf SerialPort1_DataReceived
    End Sub

    Private Sub frmMain_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize
        'redimensiona o rich textbox
        rbtReceived.Width = Me.ClientSize.Width - 28
        rbtReceived.Height = Me.ClientSize.Height - 57
        LinkFileLog.Top = Me.ClientSize.Height - 20
    End Sub

    Private Sub cmbPorta_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles cmbPorta.SelectedValueChanged
        'muda a porta serial e habilita o botão
        SerialPort1.PortName = cmbPorta.Text()
        btnConecta.Enabled = True
    End Sub

    Private Sub btnConecta_Click(sender As Object, e As System.EventArgs) Handles btnConecta.Click
        Dim msg As String = ""
        'muda a ação e o texto do botão dependendo do estado da porta
        If (isConected) Then
            SerialPort1.Close()
            btnConecta.Text = "Conectar"
            msg = vbNewLine + "Fechada porta " + SerialPort1.PortName + " às " + DateTime.Now.ToString("HH:mm:ss tt") + vbNewLine
        Else
            SerialPort1.Open()
            btnConecta.Text = "Desconectar"
            msg = "Iniciada sessão de comunicação em " + DateTime.Now.ToString("dd/MM/yyyy") + vbNewLine
            msg &= "Aberta porta " + SerialPort1.PortName + " às " + DateTime.Now.ToString("HH:mm:ss tt") + vbNewLine
        End If
        PrintInfo(msg)
        'inverte o flag
        isConected = Not isConected
        ' grava o log
        fileLog.WriteLine(msg)

    End Sub

    Delegate Sub DataDelegate(ByVal sdata As String)

    Private Sub SerialPort1_DataReceived(sender As Object, e As System.IO.Ports.SerialDataReceivedEventArgs)
        Dim ReceivedData As Integer
        Dim adreData As New DataDelegate(AddressOf PrintData)
        Dim adreCmd As New DataDelegate(AddressOf PrintCmd)
        Dim adreInfo As New DataDelegate(AddressOf PrintInfo)
        Dim adreError As New DataDelegate(AddressOf PrintError)

        Try
            ReceivedData = SerialPort1.ReadByte
        Catch ex As Exception
            ReceivedData = ex.Message
            Me.Invoke(adreError, ReceivedData + vbNewLine)
        End Try
        Me.Invoke(adreData, StringToHexString(ReceivedData.ToString))
        Dim cmd As String = ""
        If ReceivedData = 5 Then
            Me.Invoke(adreInfo, "<ENQ>")
            cmd = "21 12 34 56 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 80 AA"
            Me.Invoke(adreCmd, cmd)
            SerialPort1.Write(cmd)
        ElseIf ReceivedData = 21 Then
            Me.Invoke(adreInfo, "<OK>")
        ElseIf ReceivedData = 40 Then
            Me.Invoke(adreInfo, "<PILHA>")
            cmd = "37 12 34 56 77 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 68 B0"
            SerialPort1.Write(cmd)
        ElseIf ReceivedData = 37 Then
            Me.Invoke(adreInfo, "<PILHA_OK>")
            cmd = "37 12 34 56 77 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 68 B0"
            SerialPort1.Write(cmd)
        Else
            Me.Invoke(adreError, "<FALHA_COMUNICAÇÃO>")
        End If

    End Sub

    Private Sub PrintDataMessage(ByVal sData As String, sType As String)
        Dim mColor As Color
        Select Case sType
            Case "cmd"
                mColor = Color.Blue
            Case "info"
                mColor = Color.Green
            Case "message"
                mColor = Color.Black
            Case "error"
                mColor = Color.Red
            Case Else
                mColor = Color.DarkGray
        End Select
        rbtReceived.SelectedText = String.Empty
        rbtReceived.SelectionColor = mColor

        rbtReceived.AppendText(sData)
        rbtReceived.ScrollToCaret()
        If rbtReceived.Lines.Length > 1000 Then
            rbtReceived.Text = String.Empty
        End If
    End Sub

    Private Sub PrintData(ByVal sdata As String)
        PrintDataMessage(DateTime.Now.ToLongTimeString + " - " + sdata + " ", "message")
        fileLog.Write(sdata + " ")
    End Sub

    Private Sub PrintError(ByVal sdata As String)
        PrintDataMessage(DateTime.Now.ToLongTimeString + " - " + sdata + " ", "error")
        fileLog.Write(sdata + " ")
    End Sub

    Private Sub PrintCmd(ByVal sdata As String)
        PrintDataMessage(DateTime.Now.ToLongTimeString + " - " + sdata + " ", "cmd")
        fileLog.Write(sdata + " ")
    End Sub
    Private Sub PrintInfo(ByVal sdata As String)
        PrintDataMessage(sdata + " ", "info")
        fileLog.Write(sdata + " ")
    End Sub
    Private Sub LinkFileLog_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkFileLog.LinkClicked
        ' abre o bloco de notas com o arquivo de log
        System.Diagnostics.Process.Start("notepad.exe", urlFileLog)
    End Sub

End Class
