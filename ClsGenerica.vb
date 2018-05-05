Imports System.Data.OleDb
Imports System.IO

Public Class ClsGenerica
    Private dbcon As New OleDbConnection
    Private DBCmd As OleDbCommand
    Public DBDA As OleDbDataAdapter
    Public dbdt As DataTable
    Public strPath As String
    Public blnTestaCarregamento As Boolean = False

    Public Sub Executquery(ByVal query As String)

        Try
            dbcon.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strPath & "nomes.accdb;Persist Security Info=False;"
            dbcon.Open()
            DBCmd = New OleDbCommand(query, dbcon)
            dbdt = New DataTable
            DBDA = New OleDbDataAdapter(DBCmd)
            DBDA.Fill(dbdt)

            blnTestaCarregamento = True
        Catch ex As Exception
            MessageBox.Show("error")
        End Try

    End Sub

    Public Sub InsertRow(ByVal connectionString As String, ByVal insertSQL As String)

        Using connection As New OleDbConnection(connectionString)
            ' The insertSQL string contains a SQL statement that
            ' inserts a new row in the source table.
            ' Set the Connection to the new OleDbConnection.
            Dim command As New OleDbCommand(insertSQL) With {
                .Connection = connection
            }

            ' Open the connection and execute the insert command.
            Try
                connection.Open()
                command.ExecuteNonQuery()
            Catch ex As Exception
                MessageBox.Show("deu erro")
            Finally
                connection.Close()
            End Try

            ' The connection is automatically closed when the
            ' code exits the Using block. But...
        End Using

    End Sub

    ''' <summary>
    ''' Serve para limpar todas as TextBox e a Grid do Form Passado.
    ''' Used to clear all TextBox and Grid from the Last Form.
    ''' </summary>
    ''' <param name="frm">Objeto Form</param>
    Public Sub LimpaTextEGrid(frm As Form, blnLimpatextBox As Boolean)

        'Procura cada controle dentro do formulário.
        ''Search for each control within the form.
        For Each obj In frm.Controls

            If TypeOf obj Is TextBox Then
                obj.AutoCompleteMode = AutoCompleteMode.None
                obj.AutoCompleteSource = AutoCompleteSource.None

                If blnLimpatextBox Then obj.Text = ""

            ElseIf TypeOf obj Is DataGridView Then

                If obj.Rows.Count > 0 Then     'obj.Rows.Item(0).Cells.Item(1).Value.
                    obj.DataSource = Nothing   'Remove datasource.
                    obj.Columns.Clear()        'Remove colunas.
                    obj.Rows.Clear()           'Remove linhas.
                    obj.Refresh()              'Faz a grid atualizar-se (It makes the grid update itself).
                End If

            End If

        Next

    End Sub

    ''' <summary>
    ''' Preenche as TextBox com uma lista pronta em 'hardcode' (chumbada).
    ''' Fill in the TextBox with a ready list in 'hardcode' (leaded)
    ''' </summary>
    ''' <param name="frm">Objeto Form</param>
    Public Sub ListaPronta(frm As Form)
        Dim lst As New List(Of String)
        Dim MySource As New AutoCompleteStringCollection()
        Dim colTextBox As New Collection
        Dim Valores() As ClsVetor = {}

        'Uma lista 'hardcode' (chumbada)
        'A list 'hardcode' (leaded)
        lst.Add("Abrãao")
        lst.Add("Alceu")
        lst.Add("Alencar")
        lst.Add("Bernardo")
        lst.Add("Carlos")
        lst.Add("Gustavo")
        lst.Add("Renato")
        lst.Add("Sueli")

        For Each obj In frm.Controls

            If TypeOf obj Is TextBox Then colTextBox.Add(obj)

            If TypeOf obj Is DataGridView Then

                ReDim Valores(lst.Count)

                For index = 0 To lst.Count - 1
                    Valores(index) = New ClsVetor(lst.Item(index).ToString())
                Next

                obj.DataSource = Valores
            End If

        Next

        MySource.AddRange(lst.ToArray)

        For u = 1 To colTextBox.Count
            colTextBox(u).AutoCompleteCustomSource = MySource
            colTextBox(u).AutoCompleteMode = AutoCompleteMode.SuggestAppend
            colTextBox(u).AutoCompleteSource = AutoCompleteSource.CustomSource
        Next

    End Sub

    ''' <summary>
    ''' Serve somente para pegar na pasta "..\DB\" o Banco de Dados e o arquivo texto.
    ''' It only serves to retrieve the ".. \ DB \" folder from the Database and the text file.
    ''' Outras formas de fazer:
    ''' Other ways to do:
    ''' O Caminho é passado pelo o usuário ou recuperado de um "Settings.vb".
    ''' The Path is passed by the user or retrieved from a "Settings.vb".
    ''' </summary>
    ''' <returns>Retorna o caminho do BD e do Arquivo Texto (Returns the path of the DB and the Text File)</returns>
    Public Function PegaOCaminhoDoBancoETexto() As String
        Dim intSegundaBarra As Integer

        strPath = Application.StartupPath

        intSegundaBarra = InStrRev(strPath, "\", -1)
        intSegundaBarra = InStrRev(strPath, "\", intSegundaBarra - 1)

        strPath = strPath.Substring(0, intSegundaBarra) & "BD\"

        PegaOCaminhoDoBancoETexto = strPath
    End Function

    ''' <summary>
    ''' Essa Sub irá varrer todos os textBox dentro do form e preencher com Autocomplete com o Arquivo Texto.
    ''' This Sub will scan all textBoxes inside the form and fill with Autocomplete with the Text File.
    ''' </summary>
    ''' <param name="frm">Objeto Formulário (Form Object)</param>
    ''' <param name="strPath">Caminho do arquivo Texto de Comparação (File Path Comparison Text)</param>
    Public Sub ProcuraTextBox(frm As Form, strPath As String)
        Dim lst As New List(Of String)
        Dim colTextBox As New Collection
        Dim MySource As New AutoCompleteStringCollection()
        Dim Caminho As String
        Dim Linha As Array = {""}
        Dim Valores() As ClsVetor = {}

        Caminho = strPath & "nomes_28.txt"

        'Pega o arquivo
        'Get the file
        If File.Exists(Caminho) Then

            'Percorre o arquivo para gerar uma lista de nomes
            'Scrolls through the file to generate a list of names
            Using sr As StreamReader = File.OpenText(Caminho)

                If sr.Peek() >= 0 Then Linha = sr.ReadLine().Split

            End Using

        End If

        For u = 0 To Linha.Length - 1
            lst.Add(Linha.GetValue(u))
        Next

        For Each obj In frm.Controls

            If TypeOf obj Is TextBox Then colTextBox.Add(obj)

            If TypeOf obj Is DataGridView Then

                ReDim Valores(lst.Count)

                For index = 0 To lst.Count - 1
                    Valores(index) = New ClsVetor(lst.Item(index).ToString())
                Next

                obj.DataSource = Valores
            End If

        Next

        MySource.AddRange(lst.ToArray)

        For u = 1 To colTextBox.Count
            colTextBox(u).AutoCompleteCustomSource = MySource
            colTextBox(u).AutoCompleteMode = AutoCompleteMode.SuggestAppend
            colTextBox(u).AutoCompleteSource = AutoCompleteSource.CustomSource
        Next

    End Sub

    ''' <summary>
    ''' Procura registros no Banco de Dados e preenche as AutoComplete das TextBox e a DataGridView.
    ''' Searches for records in the Database and populates the AutoComplete of the TextBox and the DataGridView.
    ''' </summary>
    ''' <param name="frm">Objeto form (Form Object)</param>
    Public Sub ProcuraBancoDeDados(frm As Form)
        Dim lst As New List(Of String)
        Dim colTextBox As New Collection
        Dim MySource As New AutoCompleteStringCollection()

        If Not blnTestaCarregamento Then
            Executquery("select * from tblNomes")
        End If

        For Each obj In frm.Controls

            If TypeOf obj Is TextBox Then colTextBox.Add(obj)

            If TypeOf obj Is DataGridView Then obj.DataSource = dbdt

        Next

        For u = 0 To dbdt.Rows.Count - 1
            lst.Add(dbdt.Rows(u).Item(1))
        Next

        MySource.AddRange(lst.ToArray)

        For u = 1 To colTextBox.Count
            colTextBox(u).AutoCompleteCustomSource = MySource
            colTextBox(u).AutoCompleteMode = AutoCompleteMode.SuggestAppend
            colTextBox(u).AutoCompleteSource = AutoCompleteSource.CustomSource
        Next

    End Sub

    ''' <summary>
    ''' Essa Sub irá varrer o textBox dentro do form e preencher com Autocomplete com o Arquivo Texto.
    ''' This Sub will scan the textBox inside the form and fill with Autocomplete with the Text File.
    ''' </summary>
    ''' <param name="frm">Objeto Formulário (Form Object)</param>
    ''' <param name="strPath">Caminho do arquivo Texto de Comparação (File Path Comparison Text)</param>
    Public Sub InserirTextBox(frm As Form, strPath As String, txtBox As TextBox)
        Dim Caminho As String
        Dim Linha As Array = {""}
        Dim strNome As String = ""

        Caminho = strPath & "nomes_28.txt"

        'Verifica se existe a pasta - File.Exists(Caminho) Then
        'Check if the folder exists - File.Exists(Caminho) Then
        'Cria a pasta caso não exista
        'Creates the folder if it does not exist
        If Not File.Exists(Caminho) Then

            Using sw As StreamWriter = File.CreateText(Caminho)
                sw.WriteLine("")
            End Using

        End If

        'Usa a pasta que existe
        If File.Exists(Caminho) Then
            'Percorre o arquivo para gerar uma lista de nomes.
            'Scrolls through the file to generate a list of names
            Using sr As StreamReader = File.OpenText(Caminho)

                If sr.Peek() >= 0 Then Linha = sr.ReadLine().Split

            End Using

            If Not txtBox.Text = "" Then
                strNome = txtBox.Text

                'Verifica se as informacoes sao diferentes das existentes
                'Check if information is different from existing
                For Each NomeExistente As String In Linha

                    If NomeExistente IsNot String.Empty Then

                        If NomeExistente.ToUpper = strNome.ToUpper Then
                            strNome = ""
                            Exit For
                        End If

                    End If

                Next

            End If

            'Adiciona as informações ao arquivo já existente
            'Adds the information to the existing file
            If Not strNome = "" Then

                Using sw As StreamWriter = File.AppendText(Caminho)
                    sw.Write(" " & strNome)
                End Using

            End If

        End If

    End Sub

    ''' <summary>
    ''' Essa Sub irá varrer todos os textBox dentro do form e preencher com Autocomplete com o Arquivo Texto.
    ''' This Sub will scan all textBoxes within the form and populate with Autocomplete with the Textfile.
    ''' </summary>
    ''' <param name="frm">Objeto Formulário (Form Object)</param>
    ''' <param name="strPath">Caminho do arquivo Texto de Comparação (File Path Comparison Text)</param>
    Public Sub InserirTextBoxTodos(frm As Form, strPath As String)
        Dim Caminho As String
        Dim Linha As Array = {""}
        Dim strNome As String = ""

        Caminho = strPath & "nomes_28.txt"

        'Verifica se existe a pasta - File.Exists(Caminho) Then
        'Check if the folder exists - File.Exists(Caminho) Then
        'Cria a pasta caso não exista
        'Creates the folder if it does not exist
        If Not File.Exists(Caminho) Then

            Using sw As StreamWriter = File.CreateText(Caminho)
                sw.WriteLine("")
            End Using

            'Usa a pasta que existe.
            'Use the folder that exists.
        Else
            'Percorre o arquivo para gerar uma lista de nomes.
            'Scrolls through the file to generate a list of names.
            Using sr As StreamReader = File.OpenText(Caminho)

                If sr.Peek() >= 0 Then Linha = sr.ReadLine().Split

            End Using

            For Each obj In frm.Controls

                If TypeOf obj Is TextBox Then

                    If Not obj.Text = "" Then
                        strNome = obj.Text

                        'Verifica se as informações são diferentes das existentes.
                        'Check If information Is different from existing.
                        For Each NomeExistente As String In Linha

                            If NomeExistente IsNot String.Empty Then

                                If NomeExistente.ToUpper = strNome.ToUpper Then
                                    strNome = ""
                                    Exit For
                                End If

                            End If

                        Next

                        'Adiciona as informacoes ao arquivo já existente
                        'Adds information to the existing file
                        If Not strNome = "" Then

                            Using sw As StreamWriter = File.AppendText(Caminho)
                                sw.Write(" " & strNome)
                            End Using

                            strNome = ""
                        End If

                    End If

                End If

            Next

        End If

    End Sub

    ''' <summary>
    ''' Essa Sub irá varrer todos os textBox dentro do form e preencher com Autocomplete com o Banco de Dados.
    ''' This Sub will scan all textBoxes within the form and populate with Autocomplete with the Database.
    ''' </summary>
    ''' <param name="frm">Objeto Formulário (Form Object)</param>
    ''' <param name="strPath">Caminho do arquivo Texto de Comparação (File Path Comparison Text)</param>
    Public Sub InserirBancoDeDadosTodos(frm As Form, strPath As String)
        Dim Linha As Array = {""}
        Dim strNome As String = ""

        Dim lst As New List(Of String)
        Dim colTextBox As New Collection
        Dim MySource As New AutoCompleteStringCollection()
        Dim strCon As String
        Dim strSQL As String

        'Executquery("select * from tblNomes")

        For u = 0 To dbdt.Rows.Count - 1
            lst.Add(dbdt.Rows(u).Item(1))
        Next

        For Each obj In frm.Controls

            If TypeOf obj Is TextBox Then

                If Not obj.Text = "" Then
                    strNome = obj.Text

                    'Verifica se as informações são diferentes das existentes.
                    'Check If information Is different from existing.
                    For Each NomeExistente As String In lst

                        If NomeExistente IsNot String.Empty Then

                            If NomeExistente.ToUpper = strNome.ToUpper Then
                                strNome = ""
                                Exit For
                            End If

                        End If

                    Next

                    'Adiciona as informacoes ao Banco de Dados já existente
                    'Adds information to the existing Database
                    If Not strNome = "" Then

                        'Modifique de acordo com o nome de seu BD e a sua tabela no Access:
                        strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strPath & "nomes.accdb;Persist Security Info=False;"
                        strSQL = "INSERT INTO tblNomes ( Nome ) VALUES ( " & strNome & " )"

                        InsertRow(strCon, strSQL)

                        strNome = ""
                    End If

                End If

            End If

        Next

    End Sub

End Class
