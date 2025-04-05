#encoding: UTF8
Add-Type -AssemblyName PresentationFramework  # Carrega o assembly para funcionalidades de interface gráfica WPF
Add-Type -AssemblyName System.Windows.Forms   # Carrega o assembly para criar formulários Windows Forms

# Caminhos globais para os arquivos CSV que armazenam dados do programa
$global:csvPath = Join-Path -Path $PSScriptRoot -ChildPath "data.csv"          # Caminho do CSV com a lista de documentos
$global:rootFoldersPath = Join-Path -Path $PSScriptRoot -ChildPath "pastas.csv" # Caminho do CSV com as pastas raiz
$global:documentos = @()  # Array global para armazenar os documentos indexados
$global:rootFolders = @() # Array global para armazenar as pastas raiz
$global:listBox = $null   # Variável global para a ListBox da interface principal

# Funções Auxiliares
function CriarToolTip($controle, $texto) {
    # Cria uma dica de ferramenta (tooltip) para um controle específico
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.ToolTipStyle = [System.Windows.Forms.ToolTipStyle]::Standard  # Estilo padrão do tooltip
    $toolTip.UseFading = $true      # Ativa efeito de desvanecimento
    $toolTip.UseAnimation = $true   # Ativa animação
    $toolTip.IsBalloon = $false     # Desativa estilo de balão
    $toolTip.ShowAlways = $true     # Mostra mesmo quando o controle está inativo
    $toolTip.AutoPopDelay = 5000    # Tempo que o tooltip fica visível (5 segundos)
    $toolTip.InitialDelay = 500     # Atraso inicial antes de aparecer (0,5 segundos)
    $toolTip.ReshowDelay = 100      # Atraso para reaparecer (0,1 segundos)
    $toolTip.Font = New-Object System.Drawing.Font("Segoe UI", 8)  # Define a fonte do tooltip
    $toolTip.SetToolTip($controle, $texto) | Out-Null  # Associa o tooltip ao controle
}

function AplicarEstiloBotao($botao) {
    # Aplica um estilo visual consistente aos botões
    $botao.BackColor = [System.Drawing.Color]::LightGray  # Cor de fundo cinza claro
    $botao.ForeColor = [System.Drawing.Color]::Black      # Cor do texto preta
    $botao.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat  # Estilo plano
    $botao.FlatAppearance.BorderSize = 1                  # Tamanho da borda
    $botao.Font = New-Object System.Drawing.Font("Segoe UI", 9)  # Fonte do botão
}

function AplicarEstiloTextBox($textBox) {
    # Aplica estilo visual às caixas de texto
    $textBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)  # Fonte da caixa de texto
    $textBox.ForeColor = [System.Drawing.Color]::Black    # Cor do texto preta
    $textBox.BackColor = [System.Drawing.Color]::White    # Cor de fundo branca
}

function AplicarEstiloListBox($listBox) {
    # Aplica estilo visual às listas (ListBox)
    $listBox.Font = New-Object System.Drawing.Font("Consolas", 10)  # Fonte monoespaçada para melhor legibilidade
    $listBox.ForeColor = [System.Drawing.Color]::Black    # Cor do texto preta
    $listBox.BackColor = [System.Drawing.Color]::White    # Cor de fundo branca
}

function AplicarEstiloJanela($janela) {
    # Aplica estilo visual às janelas (Forms)
    $janela.BackColor = [System.Drawing.Color]::WhiteSmoke  # Cor de fundo cinza claro
    $janela.Font = New-Object System.Drawing.Font("Segoe UI", 9)  # Fonte padrão da janela
}

function Load-RootFolders {
    # Carrega as pastas raiz do arquivo CSV 'pastas.csv'
    $global:rootFolders = @()  # Limpa o array global
    if (Test-Path $global:rootFoldersPath) {  # Verifica se o arquivo existe
        $pastas = Import-Csv -Path $global:rootFoldersPath -Encoding UTF8  # Importa os dados do CSV
        $global:rootFolders = $pastas | Where-Object { $_.Caminho -and $_.Caminho -ne "" }  # Filtra entradas válidas
    }
    return $global:rootFolders  # Retorna as pastas carregadas
}

function Save-RootFolders {
    # Salva as pastas raiz no arquivo CSV 'pastas.csv'
    if ($global:rootFolders.Count -gt 0) {  # Se houver pastas para salvar
        $global:rootFolders | Export-Csv -Path $global:rootFoldersPath -NoTypeInformation -Encoding UTF8  # Exporta para CSV
    } elseif (Test-Path $global:rootFoldersPath) {  # Se não houver pastas e o arquivo existir
        Remove-Item $global:rootFoldersPath -Force  # Remove o arquivo
    }
}

function CarregarDocumentosDoCSV {
    # Carrega os documentos do arquivo CSV 'data.csv'
    $global:documentos = @()  # Limpa o array global
    if (Test-Path $global:csvPath) {  # Verifica se o arquivo existe
        $global:documentos = Import-Csv -Path $global:csvPath -Encoding UTF8  # Importa os dados do CSV
    }
    return $global:documentos  # Retorna os documentos carregados
}

function Save-CSVData {
    # Salva os documentos no arquivo CSV 'data.csv'
    $global:documentos | Export-Csv -Path $global:csvPath -NoTypeInformation -Encoding UTF8  # Exporta para CSV
}

function IndexarArquivos {
    $rootFolders = Load-RootFolders
    $global:documentos = @(
        foreach ($folder in $rootFolders) {
            if (Test-Path $folder.Caminho) {
                Get-ChildItem -Path $folder.Caminho -Recurse -File -ErrorAction SilentlyContinue |
                    ForEach-Object { [PSCustomObject]@{ Nome = $_.Name; Caminho = $_.FullName } }
            }
        }
    )
    Save-CSVData
    CarregarDocumentos
}

function AbrirArquivo($caminho) {
    # Abre um arquivo no sistema operacional
    if ([string]::IsNullOrWhiteSpace($caminho) -or -not (Test-Path $caminho)) {  # Verifica se o caminho é válido
        return
    }
    try {
        Invoke-Item -Path $caminho  # Abre o arquivo com o aplicativo padrão
    } catch {
        Write-Host "Erro ao abrir arquivo: $_"  # Exibe erro no console, se ocorrer
    }
}

function SelecionarPasta {
    # Abre um diálogo para o usuário selecionar uma pasta
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Selecione uma pasta"  # Texto exibido no diálogo
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {  # Se o usuário confirmar
        return $dialog.SelectedPath  # Retorna o caminho selecionado
    }
    return $null  # Retorna nulo se o usuário cancelar
}

function CarregarDocumentos {
    # Carrega os documentos na interface principal
    $textBoxPesquisa.Text = ""  # Limpa a caixa de pesquisa
    FiltrarDocumentos ""  # Filtra os documentos sem texto (mostra todos)
}

function FiltrarDocumentos($filtro) {
    # Filtra os documentos exibidos na ListBox com base no texto de pesquisa
    $global:listBox.Items.Clear()  # Limpa a ListBox
    $documentos = CarregarDocumentosDoCSV  # Carrega os documentos do CSV
    
    if ($filtro -eq "Pesquisar") {  # Se o texto for o placeholder "Pesquisar"
        $filtroNormalizado = ""  # Considera como vazio
    } else {
        $filtroNormalizado = $filtro.ToLower()  # Converte o filtro para minúsculas
    }

    foreach ($doc in $documentos) {  # Para cada documento
        $nome = $doc.Nome.ToLower()  # Nome em minúsculas
        $caminho = $doc.Caminho.ToLower()  # Caminho em minúsculas
        if ($nome.Contains($filtroNormalizado) -or $caminho.Contains($filtroNormalizado)) {  # Se o filtro estiver no nome ou caminho
            $global:listBox.Items.Add("$($doc.Nome) | $($doc.Caminho)")  # Adiciona à ListBox
        }
    }
}

function GerirPastasRaiz {
    # Cria uma janela para gerenciar as pastas raiz
    $janela = New-Object Windows.Forms.Form
    $janela.Text = "Gerir Pastas"  # Título da janela
    $janela.Size = New-Object Drawing.Size(500, 400)  # Tamanho da janela
    $janela.MinimumSize = New-Object Drawing.Size(400, 300)  # Tamanho mínimo
    $janela.StartPosition = "CenterScreen"  # Centraliza na tela
    AplicarEstiloJanela $janela  # Aplica estilo visual

    # Lista de Pastas
    $listPastas = New-Object Windows.Forms.ListBox
    $listPastas.Location = New-Object Drawing.Point(10, 10)  # Posição na janela
    $listPastas.Anchor = 'Top, Left, Right, Bottom'  # Ancora para redimensionamento
    $listPastas.Width = $janela.ClientSize.Width - 20  # Largura ajustada
    $listPastas.Height = 200  # Altura fixa
    AplicarEstiloListBox $listPastas  # Aplica estilo visual
    $listPastas.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended  # Permite seleção múltipla
    $janela.Controls.Add($listPastas)  # Adiciona à janela

    # Botão Adicionar Pasta
    $btnAdicionar = New-Object Windows.Forms.Button
    $btnAdicionar.Text = "Adicionar Pasta"  # Texto do botão
    $btnAdicionar.Location = New-Object Drawing.Point(10, 220)  # Posição
    $btnAdicionar.Font = New-Object Drawing.Font("Arial", 12)  # Fonte
    $btnAdicionar.Width = 140  # Largura
    $btnAdicionar.Anchor = 'Bottom, Left'  # Ancora na parte inferior esquerda
    AplicarEstiloBotao $btnAdicionar  # Aplica estilo visual
    CriarToolTip $btnAdicionar "Adiciona uma nova pasta na listagem de pastas para catalogar todos os arquivos dentro da pasta pretendida incluindo as subpastas."  # Tooltip
    $btnAdicionar.Add_Click({  # Evento de clique
        $novaPasta = SelecionarPasta  # Abre diálogo para selecionar pasta
        if ($novaPasta -and -not $listPastas.Items.Contains($novaPasta)) {  # Se válida e não duplicada
            $listPastas.Items.Add($novaPasta)  # Adiciona à lista
        }
    })
    $janela.Controls.Add($btnAdicionar)  # Adiciona à janela

    # Botão Remover Pasta
    $btnRemover = New-Object Windows.Forms.Button
    $btnRemover.Text = "Remover Pasta"  # Texto do botão
    $btnRemover.Location = New-Object Drawing.Point(160, 220)  # Posição
    $btnRemover.Font = New-Object Drawing.Font("Arial", 12)  # Fonte
    $btnRemover.Width = 140  # Largura
    $btnRemover.Anchor = 'Bottom, Left'  # Ancora na parte inferior esquerda
    AplicarEstiloBotao $btnRemover  # Aplica estilo visual
    CriarToolTip $btnRemover "Remove a pasta selecionada da listagem de pastas."  # Tooltip
    $btnRemover.Add_Click({  # Evento de clique
        $selectedItems = @($listPastas.SelectedItems)  # Obtém itens selecionados como array
        foreach ($sel in $selectedItems) {  # Para cada item selecionado
            $listPastas.Items.Remove($sel)  # Remove da lista
        }
    })
    $janela.Controls.Add($btnRemover)  # Adiciona à janela

    # Botão Guardar Alterações
    $btnGuardar = New-Object Windows.Forms.Button
    $btnGuardar.Text = "Guardar"  # Texto do botão
    $btnGuardar.Location = New-Object Drawing.Point(10, 260)  # Posição
    $btnGuardar.Font = New-Object Drawing.Font("Arial", 12)  # Fonte
    $btnGuardar.Width = $janela.ClientSize.Width - 20  # Largura ajustada
    $btnGuardar.Anchor = 'Bottom, Left, Right'  # Ancora para redimensionamento
    AplicarEstiloBotao $btnGuardar  # Aplica estilo visual
    CriarToolTip $btnGuardar "Guarda a listagem de pastas e cataloga os arquivos contidos dentro das pastas da listagem incluindo as subpastas."  # Tooltip
    $btnGuardar.Add_Click({  # Evento de clique
        $global:rootFolders = @()  # Limpa o array global
```powershell
        foreach ($item in $listPastas.Items) {  # Para cada pasta na lista
            $global:rootFolders += [PSCustomObject]@{ Caminho = $item }  # Adiciona ao array global
        }
        Save-RootFolders  # Salva as pastas no CSV
        IndexarArquivos  # Indexa os arquivos das pastas
        $janela.Close()  # Fecha a janela
    })
    $janela.Controls.Add($btnGuardar)  # Adiciona à janela

    # Evento de redimensionamento da janela
    $janela.Add_Resize({  # Ajusta a largura dos controles ao redimensionar
        $listPastas.Width = $janela.ClientSize.Width - 20
        $btnGuardar.Width = $janela.ClientSize.Width - 20
    })

    # Carrega as pastas raiz na lista
    $listPastas.Items.Clear()  # Limpa a lista
    $global:rootFolders = Load-RootFolders  # Carrega as pastas do CSV
    foreach ($pasta in $global:rootFolders) {  # Para cada pasta
        $listPastas.Items.Add($pasta.Caminho)  # Adiciona à lista
    }

    AplicarEstiloJanela $janela  # Aplica estilo visual (repetido, mas mantido por consistência)
    $janela.ShowDialog()  # Exibe a janela como modal
}

### GUI PRINCIPAL
$form = New-Object Windows.Forms.Form  # Cria a janela principal
$form.Text = "Botica | Catalogo de ficheiros"  # Título da janela
$form.Size = New-Object Drawing.Size(800, 520)  # Tamanho da janela
$form.MinimumSize = New-Object Drawing.Size(600, 400)  # Tamanho mínimo
$form.StartPosition = "CenterScreen"  # Centraliza na tela
AplicarEstiloJanela $form  # Aplica estilo visual

# TextBox de pesquisa
$textBoxPesquisa = New-Object Windows.Forms.TextBox  # Cria a caixa de texto para pesquisa
$textBoxPesquisa.Location = New-Object Drawing.Point(10, 10)  # Posição
$textBoxPesquisa.Width = 760  # Largura
$textBoxPesquisa.Height = 40  # Altura
$textBoxPesquisa.Anchor = 'Top, Left, Right'  # Ancora para redimensionamento
AplicarEstiloTextBox $textBoxPesquisa  # Aplica estilo visual
$textBoxPesquisa.Text = "Pesquisar"  # Texto inicial (placeholder)
$textBoxPesquisa.ForeColor = [System.Drawing.Color]::Gray  # Cor do placeholder
$form.Controls.Add($textBoxPesquisa)  # Adiciona à janela

# Eventos para placeholder
$textBoxPesquisa.Add_GotFocus({  # Quando a caixa ganha foco
    if ($textBoxPesquisa.Text -eq "Pesquisar") {  # Se o texto for o placeholder
        $textBoxPesquisa.Text = ""  # Limpa o texto
        $textBoxPesquisa.ForeColor = [System.Drawing.Color]::Black  # Muda a cor para preto
    }
})

$textBoxPesquisa.Add_LostFocus({  # Quando a caixa perde foco
    if ($textBoxPesquisa.Text -eq "") {  # Se estiver vazia
        $textBoxPesquisa.Text = "Pesquisar"  # Restaura o placeholder
        $textBoxPesquisa.ForeColor = [System.Drawing.Color]::Gray  # Muda a cor para cinza
    }
})

$textBoxPesquisa.Add_TextChanged({  # Quando o texto muda
    FiltrarDocumentos $textBoxPesquisa.Text  # Filtra os documentos com base no texto
})

# ListBox
$listBox = New-Object Windows.Forms.ListBox  # Cria a lista de documentos
$listBox.Location = New-Object Drawing.Point(10, 55)  # Posição
$listBox.Size = New-Object Drawing.Size(760, 350)  # Tamanho
$listBox.Anchor = 'Top, Bottom, Left, Right'  # Ancora para redimensionamento
AplicarEstiloListBox $listBox  # Aplica estilo visual
$listBox.HorizontalScrollbar = $true  # Ativa barra de rolagem horizontal
$listBox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended  # Permite seleção múltipla
$listBox.Add_DoubleClick({  # Evento de duplo clique
    foreach ($selectedItem in $listBox.SelectedItems) {  # Para cada item selecionado
        if ($selectedItem -and $selectedItem -match '\|\s*(.+)$') {  # Se válido e contém caminho
            AbrirArquivo $matches[1]  # Abre o arquivo correspondente
        }
    }
})
$listBox.Add_KeyDown({  # Evento de tecla pressionada
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {  # Se a tecla for Enter
        foreach ($selectedItem in $listBox.SelectedItems) {  # Para cada item selecionado
            if ($selectedItem -and $selectedItem -match '\|\s*(.+)$') {  # Se válido e contém caminho
                AbrirArquivo $matches[1]  # Abre o arquivo correspondente
            }
        }
    }
})
$form.Controls.Add($listBox)  # Adiciona à janela
$global:listBox = $listBox  # Atribui à variável global

# Botões
$btnReindexar = New-Object Windows.Forms.Button  # Botão para atualizar a indexação
$btnReindexar.Text = "Atualizar"  # Texto do botão
$btnReindexar.Size = New-Object Drawing.Size(120, 30)  # Tamanho
$btnReindexar.Location = New-Object Drawing.Point(10, 420)  # Posição
$btnReindexar.Anchor = 'Bottom, Left'  # Ancora na parte inferior esquerda
AplicarEstiloBotao $btnReindexar  # Aplica estilo visual
CriarToolTip $btnReindexar "Atualiza lista de arquivos catalogados conforme a listagem de pastas."  # Tooltip
$btnReindexar.Add_Click({ IndexarArquivos })  # Evento de clique
$form.Controls.Add($btnReindexar)  # Adiciona à janela

$btnPastas = New-Object Windows.Forms.Button  # Botão para gerenciar pastas
$btnPastas.Text = "Pastas"  # Texto do botão
$btnPastas.Size = New-Object Drawing.Size(120, 30)  # Tamanho
$btnPastas.Location = New-Object Drawing.Point(140, 420)  # Posição
$btnPastas.Anchor = 'Bottom, Left'  # Ancora na parte inferior esquerda
AplicarEstiloBotao $btnPastas  # Aplica estilo visual
CriarToolTip $btnPastas "Gerencia a listagem de pastas para catalogar os arquivos de interesse."  # Tooltip
$btnPastas.Add_Click({ GerirPastasRaiz })  # Evento de clique
$form.Controls.Add($btnPastas)  # Adiciona à janela

$btnSair = New-Object Windows.Forms.Button  # Botão para sair
$btnSair.Text = "Sair"  # Texto do botão
$btnSair.Size = New-Object Drawing.Size(120, 30)  # Tamanho
$btnSair.Location = New-Object Drawing.Point(270, 420)  # Posição
$btnSair.Anchor = 'Bottom, Left'  # Ancora na parte inferior esquerda
AplicarEstiloBotao $btnSair  # Aplica estilo visual
CriarToolTip $btnSair "Fecha o aplicativo."  # Tooltip
$btnSair.Add_Click({ $form.Close() })  # Evento de clique
$form.Controls.Add($btnSair)  # Adiciona à janela

# Evento de redimensionamento da janela principal
$form.Add_Resize({  # Ajusta a largura dos controles ao redimensionar
    $textBoxPesquisa.Width = $form.ClientSize.Width - 20
    $listBox.Width = $form.ClientSize.Width - 20
})

CarregarDocumentos  # Carrega os documentos na inicialização
$form.ShowDialog()  # Exibe a janela principal como modal