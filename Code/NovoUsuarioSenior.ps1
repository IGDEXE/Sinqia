# GUI - Novo Usuario Senior
# Ivo Dias

# Formulario principal
Add-Type -assembly System.Windows.Forms # Recebe a biblioteca do formulario
Add-Type -AssemblyName PresentationFramework # Recebea biblioteca das mensagens
Import-Module ActiveDirectory # Carrega o modulo
$GUI = New-Object System.Windows.Forms.Form # Cria o formulario principal
$GUI.Text ='TI - Novo Usuario Senior' # Titulo
$GUI.AutoSize = $true # Configura para aumentar caso necessario
$GUI.StartPosition = 'CenterScreen' # Inicializa no centro da tela

# Credenciais
$userADM = $env:UserName # Recebe o usuario atual
$userADM = get-aduser -identity $useradm # Pega as informacoes dele
$userfirst = $userADM.givenName # Recebe o primeiro nome
$userlast = $userADM.Surname # Recebe o sobrenome
$Domain = (Get-ADDomain).DNSRoot # Recebe o AD atual
$DomainDC = (Get-ADDomain).DistinguishedName # Recebe o nome do AD
$userAdm = "$Domain\adm.$userfirst$userlast" # Configura o ADM com o usuario atual
$CredDomain = Get-Credential -Message "Informe as credenciais de Administrador do AD" -UserName $userAdm # Recebe as credenciais 

# Nome
$lblNome = New-Object System.Windows.Forms.Label # Cria a label de indentificacao
$lblNome.Text = "Nome:" # Configura o texto
$lblNome.Location  = New-Object System.Drawing.Point(0,10) # Define a localizacao
$lblNome.AutoSize = $true # Configura o tamanho automatico
$GUI.Controls.Add($lblNome) # Adiciona ao formulario principal
$txtNome = New-Object System.Windows.Forms.TextBox # Cria a caixa de texto para o nome
$txtNome.Width = 300 # Configura o tamanho
$txtNome.Location  = New-Object System.Drawing.Point(80,10) # Define a localizacao
$GUI.Controls.Add($txtNome) # Adiciona ao formulario principal

# Sobrenome
$lblSobrenome = New-Object System.Windows.Forms.Label # Cria a label de indentificacao
$lblSobrenome.Text = "Sobrenome:" # Configura o texto
$lblSobrenome.Location  = New-Object System.Drawing.Point(0,30) # Define a localizacao
$lblSobrenome.AutoSize = $true # Configura o tamanho automatico
$GUI.Controls.Add($lblSobrenome) # Adiciona ao formulario principal
$txtSobrenome = New-Object System.Windows.Forms.TextBox # Cria a caixa de texto para o sobrenome
$txtSobrenome.Width = 300 # Configura o tamanho
$txtSobrenome.Location  = New-Object System.Drawing.Point(80,30) # Define a localizacao
$GUI.Controls.Add($txtSobrenome) # Adiciona ao formulario principal

# Setor
$lblSetor = New-Object System.Windows.Forms.Label # Cria a label de indentificacao
$lblSetor.Text = "Setor:" # Configura o texto
$lblSetor.Location  = New-Object System.Drawing.Point(0,50) # Define a localizacao
$lblSetor.AutoSize = $true # Configura o tamanho automatico
$GUI.Controls.Add($lblSetor) # Adiciona ao formulario principal
$cbxSetores = New-Object System.Windows.Forms.ComboBox # Cria a Combobox dos setores
$cbxSetores.Width = 300 # Configura o tamanho
$cbxSetores.Location  = New-Object System.Drawing.Point(80,50) # Define a localizacao
$Setores = get-content "\\SERVIDOR\Scripts\SRE\Ref\ouList.txt" # Recebe os setores da lista no caminho informado
foreach ($setor in $Setores) {
    $cbxSetores.Items.Add($setor) # Adiciona como opcao cada um dos setores
}
$GUI.Controls.Add($cbxSetores) # Adiciona ao formulario principal

# Usuario de referencia
$lblReferencia = New-Object System.Windows.Forms.Label # Cria a label de indentificacao
$lblReferencia.Text = "Referencia:" # Configura o texto
$lblReferencia.Location  = New-Object System.Drawing.Point(0,70) # Define a localizacao
$lblReferencia.AutoSize = $true # Configura o tamanho automatico
$GUI.Controls.Add($lblReferencia) # Adiciona ao formulario principal
$txtReferencia = New-Object System.Windows.Forms.TextBox # Cria a caixa de texto para o usuario de referencia
$txtReferencia.Width = 300 # Configura o tamanho
$txtReferencia.Location  = New-Object System.Drawing.Point(80,70) # Define a localizacao
$GUI.Controls.Add($txtReferencia) # Adiciona ao formulario principal

# Grupos de referencia
$lblGrupos = New-Object System.Windows.Forms.Label # Cria a label de indentificacao
$lblGrupos.Text = "Grupos Ref:" # Configura o texto
$lblGrupos.Location  = New-Object System.Drawing.Point(0,90) # Define a localizacao
$lblGrupos.AutoSize = $true # Configura o tamanho automatico
$GUI.Controls.Add($lblGrupos) # Adiciona ao formulario principal
$cbxGrupos = New-Object System.Windows.Forms.CheckedListBox # Cria a caixa de texto para as opcoes dos grupos
$cbxGrupos.Width = 300 # Configura o tamanho
$cbxGrupos.Location  = New-Object System.Drawing.Point(80,90) # Define a localizacao
$cbxGrupos.Enabled = $false # Deixa desativado ao iniciar
$GUI.Controls.Add($cbxGrupos) # Adiciona ao formulario principal

# Grupos padroes
$lblReferencia2 = New-Object System.Windows.Forms.Label # Cria a label de indentificacao
$lblReferencia2.Text = "Padrao:" # Configura o texto
$lblReferencia2.Location  = New-Object System.Drawing.Point(380,90) # Define a localizacao
$lblReferencia2.AutoSize = $true # Configura o tamanho automatico
$GUI.Controls.Add($lblReferencia2) # Adiciona ao formulario principal
$cbxReferencia = New-Object System.Windows.Forms.CheckedListBox # Cria a caixa de texto para as opcoes dos grupos de referencia
$cbxReferencia.Width = 300 # Configura o tamanho
$cbxReferencia.Location  = New-Object System.Drawing.Point(430,90) # Define a localizacao
$cbxReferencia.Enabled = $false # Deixa desativado ao iniciar
$GUI.Controls.Add($cbxReferencia) # Adiciona ao formulario principal

# Botao de Criar
$Button = New-Object System.Windows.Forms.Button # Cria o botao
$Button.Location = New-Object System.Drawing.Size(550,27) # Define a localizacao
$Button.Size = New-Object System.Drawing.Size(120,50) # Configura o tamanho
$Button.Text = "Criar" # Configura o texto
$Button.Enabled = $false # Deixa desativado ao iniciar
$GUI.Controls.Add($Button) # Adiciona ao formulario principal

# Botao de limpeza
$Limpar = New-Object System.Windows.Forms.Button # Cria o botao
$Limpar.Location = New-Object System.Drawing.Size(400,27) # Define a localizacao
$Limpar.Size = New-Object System.Drawing.Size(120,20) # Configura o tamanho
$Limpar.Text = "Limpar" # Configura o texto
$GUI.Controls.Add($Limpar) # Adiciona ao formulario principal

# Botao de verificar
$btnVerificar = New-Object System.Windows.Forms.Button # Cria o botao
$btnVerificar.Location = New-Object System.Drawing.Size(400,57) # Define a localizacao
$btnVerificar.Size = New-Object System.Drawing.Size(120,20) # Configura o tamanho
$btnVerificar.Text = "Verificar" # Configura o texto
$GUI.Controls.Add($btnVerificar) # Adiciona ao formulario principal

# Campo de retorno
$lblResposta = New-Object System.Windows.Forms.Label # Cria a label para receber o retorno
$lblResposta.Text = "" # Coloca o texto inicial vazio
$lblResposta.Location  = New-Object System.Drawing.Point(0,190) # Define a localizacao
$lblResposta.AutoSize = $true # Configura o tamanho automatico
$GUI.Controls.Add($lblResposta) # Adiciona ao formulario principal

# Label de retorno 2 linha
$lbl2linha = New-Object System.Windows.Forms.Label # Cria a label para receber o retorno
$lbl2linha.Text = "" # Coloca o texto inicial vazio
$lbl2linha.Location  = New-Object System.Drawing.Point(0,210) # Define a localizacao
$lbl2linha.AutoSize = $true # Configura o tamanho automatico
$GUI.Controls.Add($lbl2linha) # Adiciona ao formulario principal

# Botao para verificar 
$btnVerificar.Add_Click(
    {
        # Verifica o usuario de referencia
        try {
            # Configura as referencias de grupo
            $referencia = $txtReferencia.Text # Recebe o usuario de referencia
            $cbxGrupos.Enabled = $true # Deixa o campo ativo
            $cbxReferencia.Enabled = $true # Deixa o campo ativo
            # Adiciona os grupos principais
            $gruposPrincipais = get-content "\\SERVIDOR\Scripts\SRE\Ref\grupoList.txt" # Recebe os grupos de uma lista
            foreach ($grupo in $gruposPrincipais) {
                $cbxReferencia.Items.Add($grupo) # Adiciona como opcao cada um dos setores
            }
            $userRef = Get-ADUser -Identity $referencia -Properties * # Recebe os dados do usuario de referencia
            $Grupos = $userRef.MemberOf # Recebe os grupos do usuario de referencia
            foreach ($Grupo in $Grupos) {
                # Adiciona como opcao cada um dos setores
                $nomeGrupo = Get-ADGroup -Identity $Grupo # Recebe o nome do grupo
                $nomeGrupo = $nomeGrupo.SamAccountName # Usa ele para verificar o identificador
                $cbxGrupos.Items.Add($nomeGrupo) # Adiciona como campo no Checkbox
            }
            $Button.Enabled = $true # Ativa o botao
        }
        catch {
            # Reporta o erro, caso ocorra
            $ErrorMessage = $_.Exception.Message
            $resposta = "|Ocorreu um verificar o usuario $referencia|"
            $Linha2 = "|Erro: $ErrorMessage|"
            $lblResposta.Text =  $resposta
            $lbl2linha.Text = $Linha2
        }
    }
)

# Botao para criar
$Button.Add_Click(
    {
        # Faz o procedimento
        try {
            # Recebe os dados
            $nome = $txtNome.Text # Recebe o nome
            $sobrenome = $txtSobrenome.Text # Recebe o sobrenome
            $setor = $cbxSetores.selectedItem # Recebe o setor
            # Configura os dados
            $userName = "$nome $sobrenome" # Cria o nome de exibicao
            $userAlias = ("$nome.$sobrenome").ToLower() # Cria o logon
            $userMail = "$userAlias@sinqia.com.br" # Cria o email
            $OU = "OU=$setor,OU=SeniorSolution,$DomainDC" # Configura a OU
            $Password = Get-Date -Format Sin@mmss # Cria uma senha
            # Cria o usuario
            New-ADUser -SamAccountName $userAlias -Name $userName -GivenName $nome -Surname $sobrenome -EmailAddress $userMail -Path "$OU" -Credential $CredDomain
            Set-ADUser -Identity $userAlias -UserPrincipalName $userMail -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Replace @{'msRTCSIP-PrimaryUserAddress'="SIP:$userMail"} -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Replace @{proxyAddresses="SMTP:$userMail","SIP:$userMail"} -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Replace @{targetAddress="SMTP:$userAlias@ATTPS.mail.onmicrosoft.com"} -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Replace @{DisplayName="$userName"} -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Replace @{displayNamePrintable="$userName"} -Credential $CredDomain
            Set-ADUser -Identity $userAlias -Replace @{mailNickname="$userAlias"} -Credential $CredDomain
            Set-ADAccountPassword -Identity $userAlias -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force) -Credential $CredDomain # Configura a senha
            Set-ADUser -Identity $Conta -changepasswordatlogon $true -Credential $CredDomain # Configura para trocar a senha no proximo login
            Set-ADUser -Identity $userAlias -Enabled $true -Credential $CredDomain # Deixa a conta ativa
            # Configura as referencias de grupo
            foreach ($Item in $cbxGrupos.CheckedItems) {
                $grupo = $Item.ToString() 
                Add-ADGroupMember -Identity $grupo -Members $userAlias -Credential $CredDomain # Adiciona os grupos que forma selecionados no checkbox
            }
            foreach ($Item in $cbxReferencia.CheckedItems) {
                $grupo = $Item.ToString()
                Add-ADGroupMember -Identity $grupo -Members $userAlias -Credential $CredDomain # Adiciona os grupos que forma selecionados no checkbox
            }
            # Mostra na tela
            $resposta = "Usuario $userName criado com sucesso"
            $Linha2 = "Senha: $Password"
        }
        catch {
            # Reporta o erro, caso ocorra
            $ErrorMessage = $_.Exception.Message
            $resposta = "|Ocorreu um erro ao criar o usuario $userName|"
            $Linha2 = "|Erro: $ErrorMessage|"
        }
        # Limpa os campos na tela
        $lblResposta.Text =  $resposta
        $lbl2linha.Text = $Linha2
        $txtNome.Text = ""
        $txtSobrenome.Text = ""
        $cbxSetores.Text = ""
        $txtReferencia.Text = ""
        $Button.Enabled = $false
        $cbxGrupos.Enabled = $false
        $cbxGrupos.Items.Clear()
        $cbxReferencia.Enabled = $false
        $cbxReferencia.Items.Clear()
    }
)

$Limpar.Add_Click(
    {
        # Limpa os campos
        $cbxReferencia.Items.Clear()
        $cbxGrupos.Items.Clear()
        $txtNome.Text = ""
        $txtSobrenome.Text = ""
        $cbxSetores.Text = ""
        $txtReferencia.Text = ""
        $lblResposta.Text = ""
        $lbl2linha.Text = ""
    }
)

# Inicia o formulario
$GUI.ShowDialog()