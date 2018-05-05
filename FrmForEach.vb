Public Class FrmForEach
    Private clsGen As New ClsGenerica
    Private strPath As String

    Private Sub FrmForEach_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Somente para mostrar onde está o Banco de Dados e o Texto.
        'Only to show where Database and Text are.
        strPath = clsGen.PegaOCaminhoDoBancoETexto()
    End Sub

    ''' <summary>
    ''' Limpa tudo.
    ''' Clean everything.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnLimpaTudo_Click(sender As Object, e As EventArgs) Handles BtnLimpaTudo.Click
        clsGen.LimpaTextEGrid(Me, True)
    End Sub

    ''' <summary>
    ''' Ele preencherá TODAS as TextBox que encontrar com a lista abaixo
    ''' It will populate ALL the TextBox you find with the list below.
    ''' </summary>
    Private Sub BtnCompletaComListaPronta_Click(sender As Object, e As EventArgs) Handles BtnCompletaComListaPronta.Click
        clsGen.LimpaTextEGrid(Me, True)
        clsGen.ListaPronta(Me)
    End Sub

    ''' <summary>
    ''' Procura registros no Banco de Dados e preenche as AutoComplete das TextBox e a DataGridView.
    ''' Searches for records in the Database and populates the AutoComplete of the TextBox and the DataGridView.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnProcuraNoBancoDeDados_Click(sender As Object, e As EventArgs) Handles BtnProcuraNoBancoDeDados.Click
        clsGen.LimpaTextEGrid(Me, True)
        clsGen.ProcuraBancoDeDados(Me)
    End Sub

    ''' <summary>
    ''' Procura registros no Arquivo Texto e preenche as AutoComplete das TextBox e a DataGridView.
    ''' Searches for records in the Text File and populates the AutoComplete of the TextBox and the DataGridView.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnProcuraTextBox_Click(sender As Object, e As EventArgs) Handles BtnProcuraTextBox.Click
        clsGen.LimpaTextEGrid(Me, True)
        clsGen.ProcuraTextBox(Me, strPath)
    End Sub

    ''' <summary>
    ''' Somente no "Leave" do TextBox1 coloquei uma inserção automática.
    ''' Only in the "Leave" of TextBox1 I put an automatic insertion.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TextBox1_Leave(sender As Object, e As EventArgs) Handles TextBox1.Leave
        clsGen.InserirTextBox(Me, strPath, TextBox1)
        clsGen.LimpaTextEGrid(Me, False)
        clsGen.ProcuraTextBox(Me, strPath)
    End Sub

    ''' <summary>
    ''' Neste caso ele irá procurar nomes em TODOS os TextBox, os que não estiverem na lista ele adicionará automaticamente.
    ''' In this case it will look for names in ALL TextBox, those that are not in the list will add it automatically.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnInsereNoTextBoxTodos_Click(sender As Object, e As EventArgs) Handles BtnInsereNoTextBoxTodos.Click
        clsGen.InserirTextBoxTodos(Me, strPath)
        clsGen.LimpaTextEGrid(Me, False)
        clsGen.ProcuraTextBox(Me, strPath)
    End Sub

    Private Sub BtnInsereNoBancoDeDados_Click(sender As Object, e As EventArgs) Handles BtnInsereNoBancoDeDados.Click
        clsGen.InserirBancoDeDadosTodos(Me, strPath)
        clsGen.LimpaTextEGrid(Me, False)
        clsGen.ProcuraBancoDeDados(Me)
    End Sub

    '''' <summary>
    '''' Não consegui criar um LEAVE para todos de uma só vez, não necessitando colocar individualmente... mas acho que não é muito útil, afinal...
    '''' I couldn't create a LEAVE for everyone at once, not needing to put it individually ... but I think it isn't very useful, after all ...
    '''' ESSA ROTINA NÃO FUNCIONA.
    '''' THIS ROUTINE DOES NOT WORK.
    '''' </summary>
    '''' <param name="sender"></param>
    '''' <param name="e"></param>
    'Private Sub TextBox_Leave(sender As Object, e As EventArgs) Handles TextBox1.Leave, TextBox2.Leave, TextBox3.Leave, TextBox4.Leave, TextBox5.Leave, TextBox6.Leave, TextBox7.Leave, TextBox8.Leave, TextBox9.Leave, TextBox10.Leave, TextBox11.Leave, TextBox12.Leave, TextBox13.Leave, TextBox14.Leave, TextBox15.Leave, TextBox16.Leave, TextBox17.Leave, TextBox18.Leave, TextBox19.Leave, TextBox20.Leave, TextBox21.Leave, TextBox22.Leave, TextBox23.Leave, TextBox24.Leave, TextBox25.Leave, TextBox26.Leave, TextBox27.Leave, TextBox28.Leave, TextBox29.Leave, TextBox30.Leave
    '    clsGen.InserirTextBoxTodos(Me, strPath)
    'End Sub

End Class
