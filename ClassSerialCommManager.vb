'*************************************************************************
'   Classe para gerenciar comunicação serial com medidores de energia
'
'   Marley Adriano - marleyas@gmail.com
'   Baseado no código de Richard L. McCutchen - richard@psychocoder.net
'
'************************************************************8*************
Imports System
Imports System.Text
Imports System.Drawing
Imports System.IO.Ports
Imports System.Windows.Forms

Public Class SerialCommManager

    ' Enumerador para definir o tipo de transmissão (padrão Hexa)
    Public Enum TransmissionType
        Text
        Hex
    End Enum
    ' Enumerador para definir o tipo de mensagem a ser exibida
    Public Enum MessageType
        Incoming
        Outgoing
        Normal
        Warning
        [Error]
    End Enum

    ' variaveis internas para configuração da comunicação
    Private _baudRate As String = String.Empty
    Private _parity As String = String.Empty
    Private _stopBits As String = String.Empty
    Private _dataBits As String = String.Empty
    Private _portName As String = String.Empty
    Private _transType As TransmissionType
    Private _displayWindow As RichTextBox
    Private _msg As String
    Private _type As MessageType

    Private MessageColor As Color() = {Color.Blue, Color.Green, Color.Black, Color.Orange, Color.Red}
    Private comPort As New SerialPort()
    Private write As Boolean = True

    'definição do Baudrate
    Public Property BaudRate() As String
        Get
            Return _baudRate
        End Get
        Set(ByVal value As String)
            _baudRate = value
        End Set
    End Property

    ' definição da paridade
    Public Property Parity() As String
        Get
            Return _parity
        End Get
        Set(ByVal value As String)
            _parity = value
        End Set
    End Property

    ' definição do bit de stop
    Public Property StopBits() As String
        Get
            Return _stopBits
        End Get
        Set(ByVal value As String)
            _stopBits = value
        End Set
    End Property

    ' definição do bit de dados
    Public Property DataBits() As String
        Get
            Return _dataBits
        End Get
        Set(ByVal value As String)
            _dataBits = value
        End Set
    End Property

    ' definição do nome da porta
    Public Property PortName() As String
        Get
            Return _portName
        End Get
        Set(ByVal value As String)
            _portName = value
        End Set
    End Property

    ' definição do tipo de transmissão
    Public Property CurrentTransmissionType() As TransmissionType
        Get
            Return _transType
        End Get
        Set(ByVal value As TransmissionType)
            _transType = value
        End Set
    End Property

    ' janela de exibição das mensagens, será num rich textbox
    Public Property DisplayWindow() As RichTextBox
        Get
            Return _displayWindow
        End Get
        Set(ByVal value As RichTextBox)
            _displayWindow = value
        End Set
    End Property

    ' definição da mensagem
    Public Property Message() As String
        Get
            Return _msg
        End Get
        Set(ByVal value As String)
            _msg = value
        End Set
    End Property

    ' tipo da mensagem
    Public Property Type() As MessageType
        Get
            Return _type
        End Get
        Set(ByVal value As MessageType)
            _type = value
        End Set
    End Property

    ' construtor da classe com a definição das propriedades
    Public Sub New(ByVal baud As String, ByVal par As String, ByVal sBits As String, ByVal dBits As String, ByVal name As String, ByVal rtb As RichTextBox)
        _baudRate = baud
        _parity = par
        _stopBits = sBits
        _dataBits = dBits
        _portName = name
        _displayWindow = rtb
        ' adicionando um event handler
        AddHandler comPort.DataReceived, AddressOf comPort_DataReceived
    End Sub

    ' construtor "limpo"
    Public Sub New()
        _baudRate = String.Empty
        _parity = String.Empty
        _stopBits = String.Empty
        _dataBits = String.Empty
        _portName = "COM1"
        _displayWindow = Nothing
        'adicionando um event handler
        AddHandler comPort.DataReceived, AddressOf comPort_DataReceived
    End Sub

    ' enviando dados para a serial
    Public Sub WriteData(ByVal msg As String)
        Select Case CurrentTransmissionType
            Case TransmissionType.Text
                ' primeiro ver se a porta está aberta, ou abri-la
                If Not (comPort.IsOpen = True) Then
                    comPort.Open()
                End If
                ' enviar a mensagem para a porta
                comPort.Write(msg)
                ' exibir a mensagem
                _type = MessageType.Outgoing
                _msg = msg + "" + Environment.NewLine + ""
                DisplayData(_type, _msg)
                Exit Select
            Case TransmissionType.Hex
                Try
                    ' converter a mensagem para um array de bytes
                    Dim newMsg As Byte() = HexToByte(msg)
                    If Not write Then
                        DisplayData(_type, _msg)
                        Exit Sub
                    End If
                    ' enviar a mensagem para a porta
                    comPort.Write(newMsg, 0, newMsg.Length)
                    ' converter novamente para hexa e exibir
                    _type = MessageType.Outgoing
                    _msg = ByteToHex(newMsg) + "" + Environment.NewLine + ""
                    DisplayData(_type, _msg)
                Catch ex As FormatException
                    ' exibir uma mensagem de erro
                    _type = MessageType.Error
                    _msg = ex.Message + "" + Environment.NewLine + ""
                    DisplayData(_type, _msg)
                Finally
                    _displayWindow.SelectAll()
                End Try
                Exit Select
            Case Else
                ' primeiro ver se a porta está aberta, ou abri-la                '
                If Not (comPort.IsOpen = True) Then
                    comPort.Open()
                End If
                ' enviar a mensagem para a porta
                comPort.Write(msg)
                ' exibir a mensagem
                _type = MessageType.Outgoing
                _msg = msg + "" + Environment.NewLine + ""
                DisplayData(MessageType.Outgoing, msg + "" + Environment.NewLine + "")
                Exit Select
        End Select
    End Sub

    Private Function HexToByte(ByVal msg As String) As Byte()
        If msg.Length Mod 2 = 0 Then
            ' remover espaços na string
            _msg = msg
            _msg = msg.Replace(" ", "")
            ' criar um array de bytes de comprimento
            'dividido por 2 (Hex possui 2 caracteres no comprimento)
            Dim comBuffer As Byte() = New Byte(_msg.Length / 2 - 1) {}
            For i As Integer = 0 To _msg.Length - 1 Step 2
                comBuffer(i / 2) = CByte(Convert.ToByte(_msg.Substring(i, 2), 16))
            Next
            write = True
            'percorreu o loop do comprimento da string
            'converter todo par de 2 caracteres em um byte
            'e adicionar para o array
            'returnar o array
            Return comBuffer
        Else
            _msg = "Formato inválido"
            _type = MessageType.Error
            DisplayData(_type, _msg)
            write = False
            Return Nothing
        End If
    End Function

    Private Function ByteToHex(ByVal comByte As Byte()) As String
        ' cria um novo objeto StringBuilder
        Dim builder As New StringBuilder(comByte.Length * 3)
        ' percorre em loop cada byte desse array
        For Each data As Byte In comByte
            builder.Append(Convert.ToString(data, 16).PadLeft(2, "0"c).PadRight(3, " "c))
            ' converte o byte para uma string e adiciona ao stringbuilder
        Next
        ' returna o valor convertido
        Return builder.ToString().ToUpper()
    End Function

    ' ajusta a thread para exibir a mensagem (alternativa elegante ao DoEvents)
    <STAThread()> _
    Private Sub DisplayData(ByVal type As MessageType, ByVal msg As String)
        _displayWindow.Invoke(New EventHandler(AddressOf DoDisplay))
    End Sub

    Public Function OpenPort() As Boolean
        Try
            ' primeiro ver se a porta está aberta 
            ' se estiver aberta, fecha-la
            If comPort.IsOpen = True Then
                comPort.Close()
            End If

            ' definir as propriedades do objeto SerialPort
            comPort.BaudRate = Integer.Parse(_baudRate)
            'BaudRate
            comPort.DataBits = Integer.Parse(_dataBits)
            'DataBits
            comPort.StopBits = DirectCast([Enum].Parse(GetType(StopBits), _stopBits), StopBits)
            'StopBits
            comPort.Parity = DirectCast([Enum].Parse(GetType(Parity), _parity), Parity)
            'Parity
            comPort.PortName = _portName
            'PortName
            ' agora sim, abrir a porta
            comPort.Open()
            ' exibir a mensagem de porta aberta
            _type = MessageType.Normal
            _msg = "Porta " + _portName + " aberta em " + DateTime.Now + "" + Environment.NewLine + ""
            DisplayData(_type, _msg)
            'returna true
            Return True
        Catch ex As Exception
            ' exibe uma mensagem de erro ao abrir porta
            DisplayData(MessageType.[Error], ex.Message)
            Return False
        End Try
    End Function

    Public Sub ClosePort()
        If comPort.IsOpen Then
            _msg = "Porta " + _portName + " fechada em " + DateTime.Now + "" + Environment.NewLine + ""
            _type = MessageType.Normal
            ' exibir a mensagem de porta fechada
            DisplayData(_type, _msg)
            comPort.Close()
        End If
    End Sub
    '**************************************************************************
    '   REFORMULAR GERAL
    '
    Public Sub SetParityValues(ByVal obj As Object)
        For Each str As String In [Enum].GetNames(GetType(Parity))
            DirectCast(obj, ComboBox).Items.Add(str)
        Next
    End Sub

    Public Sub SetStopBitValues(ByVal obj As Object)
        For Each str As String In [Enum].GetNames(GetType(StopBits))
            DirectCast(obj, ComboBox).Items.Add(str)
        Next
    End Sub

    Public Sub SetPortNameValues(ByVal obj As Object)
        For Each str As String In SerialPort.GetPortNames()
            DirectCast(obj, ComboBox).Items.Add(str)
        Next
    End Sub
    '
    '   ATÉ AQUI
    '**************************************************************************

    Private Sub comPort_DataReceived(ByVal sender As Object, ByVal e As SerialDataReceivedEventArgs)
        ' é determinada pelo modo de recepção de dados que o usuário selecionou (binary/string)
        Select Case CurrentTransmissionType
            Case TransmissionType.Text
                ' o usuário selecionou string
                ' lê o dado no buffer
                Dim msg As String = comPort.ReadExisting()
                ' exibe o dado para o usuário
                _type = MessageType.Incoming
                _msg = msg
                DisplayData(MessageType.Incoming, msg + "" + Environment.NewLine + "")
                Exit Select
            Case TransmissionType.Hex
                ' o usuário selecionou binário
                ' retorna um número de bytes do buffer
                Dim bytes As Integer = comPort.BytesToRead
                ' cria um array de bytes para armazenar o resultado a ser lido
                Dim comBuffer As Byte() = New Byte(bytes - 1) {}
                ' lê os dados e armazenam no buffer interno
                comPort.Read(comBuffer, 0, bytes)
                ' exibe o dado para o usuário
                _type = MessageType.Incoming
                _msg = ByteToHex(comBuffer) + "" + Environment.NewLine + ""
                DisplayData(MessageType.Incoming, ByteToHex(comBuffer) + "" + Environment.NewLine + "")
                Exit Select
            Case Else
                ' lê os dados a espera no buffer
                Dim str As String = comPort.ReadExisting()
                ' exibe o dado para o usuário
                _type = MessageType.Incoming
                _msg = str + "" + Environment.NewLine + ""
                DisplayData(MessageType.Incoming, str + "" + Environment.NewLine + "")
                Exit Select
        End Select
    End Sub

    ' configura o objeto de exibição dos dados 
    Private Sub DoDisplay(ByVal sender As Object, ByVal e As EventArgs)

        _displayWindow.SelectedText = String.Empty
        _displayWindow.SelectionFont = New Font(_displayWindow.SelectionFont, FontStyle.Bold)
        _displayWindow.SelectionColor = MessageColor(CType(_type, Integer))
        '_msg = _msg + ""
        _displayWindow.AppendText(_msg)
        _displayWindow.ScrollToCaret()
    End Sub

End Class
