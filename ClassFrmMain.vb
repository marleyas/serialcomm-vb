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

    Private Function StringToHexString(ByVal strInput As String) As String
        strInput = "00" & strInput
        Return Strings.Right(strInput, 2)
    End Function

    Private Sub frmMain_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        SerialPort1.Close()
    End Sub

    Private Sub frmMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fileLog = My.Computer.FileSystem.OpenTextFileWriter(DateTime.Now.ToString("dd-MM-yyyy") + ".log", True)
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
        rbtReceived.Height = Me.ClientSize.Height - 47
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
            msg = vbNewLine + "Fechada porta " + SerialPort1.PortName + " " + DateTime.Now.ToString("HH:mm:ss tt") + vbNewLine

        Else
            SerialPort1.Open()
            btnConecta.Text = "Desconectar"
            msg = "Iniciada sessão de comunicação em " + DateTime.Now.ToString("dd/MM/yyyy") + vbNewLine
            msg &= "Aberta porta " + SerialPort1.PortName + " " + DateTime.Now.ToString("HH:mm:ss tt") + vbNewLine
        End If
        'inverte o flag
        isConected = Not isConected
        rbtReceived.Text &= msg
        fileLog.WriteLine(msg)
    End Sub

    Delegate Sub DataDelegate(ByVal sdata As String)

    Private Sub SerialPort1_DataReceived(sender As Object, e As System.IO.Ports.SerialDataReceivedEventArgs)
        If isConected Then
            Dim ReceivedData As Integer
            Try
                ReceivedData = SerialPort1.ReadByte
            Catch ex As Exception
                ReceivedData = ex.Message
            End Try

            Dim adre As New DataDelegate(AddressOf PrintData)
            Me.Invoke(adre, ReceivedData.ToString)
        End If
    End Sub

    Private Sub PrintData(ByVal sdata As String)
        rbtReceived.Text &= StringToHexString(sdata) + " "
        fileLog.Write(sdata + " ")
    End Sub

    Protected Overrides Sub Finalize()
        fileLog.Close()
        MyBase.Finalize()
    End Sub
End Class
