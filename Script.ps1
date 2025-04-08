#encoding: Windows-1252
Add-Type -AssemblyName PresentationFramework  # Carrega o assembly para funcionalidades de interface gráfica WPF
Add-Type -AssemblyName System.Windows.Forms   # Carrega o assembly para criar formulários Windows Forms

# Caminhos globais para os arquivos JSON que armazenam dados do programa
$global:jsonPath = Join-Path -Path $PSScriptRoot -ChildPath "data.json"          # Caminho do JSON com a lista de documentos
$global:rootFoldersJsonPath = Join-Path -Path $PSScriptRoot -ChildPath "pastas.json" # Caminho do JSON com as pastas raiz
$global:documentos = @()  # Array global para armazenar os documentos indexados
$global:rootFolders = @() # Array global para armazenar as pastas raiz
$global:listBox = $null   # Variável global para a ListBox da interface principal
$global:indexJob = $null  # Variável global para armazenar o job de indexação

# Funções Auxiliares
function CriarToolTip($controle, $texto) {
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.ToolTipStyle = [System.Windows.Forms.ToolTipStyle]::Standard
    $toolTip.UseFading = $true
    $toolTip.UseAnimation = $true
    $toolTip.IsBalloon = $false
    $toolTip.ShowAlways = $true
    $toolTip.AutoPopDelay = 5000
    $toolTip.InitialDelay = 500
    $toolTip.ReshowDelay = 100
    $toolTip.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $toolTip.SetToolTip($controle, $texto) | Out-Null
}

function AplicarEstiloBotao($botao) {
    $botao.BackColor = [System.Drawing.Color]::LightGray
    $botao.ForeColor = [System.Drawing.Color]::Black
    $botao.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $botao.FlatAppearance.BorderSize = 1
    $botao.Font = New-Object System.Drawing.Font("Segoe UI", 9)
}

function AplicarEstiloTextBox($textBox) {
    $textBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $textBox.ForeColor = [System.Drawing.Color]::Black
    $textBox.BackColor = [System.Drawing.Color]::White
}

function AplicarEstiloListBox($listBox) {
    $listBox.Font = New-Object System.Drawing.Font("Consolas", 10)
    $listBox.ForeColor = [System.Drawing.Color]::Black
    $listBox.BackColor = [System.Drawing.Color]::White
}

function AplicarEstiloJanela($janela) {
    $janela.BackColor = [System.Drawing.Color]::WhiteSmoke
    $janela.Font = New-Object System.Drawing.Font("Segoe UI", 9)
}

function Load-RootFolders {
    $global:rootFolders = @()
    if (-not (Test-Path -Path $global:rootFoldersJsonPath -PathType Leaf)) { return @() }
    
    try {
        $global:rootFolders = [System.IO.File]::ReadAllText($global:rootFoldersJsonPath, [System.Text.Encoding]::Default) | 
            ConvertFrom-Json | 
            Where-Object { $_.Caminho -and $_.Caminho -ne "" }
    } catch {
        Write-Host "Erro ao carregar pastas.json: $_"
    }
    return $global:rootFolders
}

function Save-RootFolders {
    if (-not $global:rootFolders) { return }
    
    try {
        if ($global:rootFolders.Count -gt 0) {
            $json = $global:rootFolders | ConvertTo-Json -Depth 10 -Compress
            [System.IO.File]::WriteAllText($global:rootFoldersJsonPath, $json, [System.Text.Encoding]::Default)
        } elseif (Test-Path -Path $global:rootFoldersJsonPath -PathType Leaf) {
            Remove-Item -Path $global:rootFoldersJsonPath -Force
        }
    } catch {
        Write-Host "Erro ao salvar pastas.json: $_"
    }
}

function CarregarDocumentosDoJson {
    $global:documentos = @()
    if (-not (Test-Path -Path $global:jsonPath -PathType Leaf)) { return @() }
    
    try {
        $global:documentos = [System.IO.File]::ReadAllText($global:jsonPath, [System.Text.Encoding]::Default) | 
            ConvertFrom-Json
    } catch {
        Write-Host "Erro ao carregar data.json: $_"
    }
    return $global:documentos
}

function Save-JsonData {
    if (-not $global:documentos) { return }
    
    try {
        $json = $global:documentos | ConvertTo-Json -Depth 10 -Compress
        [System.IO.File]::WriteAllText($global:jsonPath, $json, [System.Text.Encoding]::Default)
    } catch {
        Write-Host "Erro ao salvar data.json: $_"
    }
}

# Função IndexarArquivos com melhoria de desempenho
function IndexarArquivos {
    # Inicia a indexação em segundo plano
    if ($global:indexJob -and $global:indexJob.State -eq "Running") {
        Write-Host "Catalogo em andamento. Aguarde conclir."
        return
    }

    $rootFolders = Load-RootFolders
    $statusLabel.Text = "A indexar arquivos..."
    $progressBar.Visible = $true
    $progressBar.Value = 0
    $btnReindexar.Enabled = $false  # Desativa o botão enquanto indexa

    $global:indexJob = Start-Job -ScriptBlock {
        $docs = @(
            foreach ($folder in $using:rootFolders) {
                if (Test-Path $folder.Caminho) {
                    Get-ChildItem -Path $folder.Caminho -Recurse -File -ErrorAction SilentlyContinue |
                        ForEach-Object { [PSCustomObject]@{ Nome = $_.Name; Caminho = $_.FullName } }
                }
            }
        )
        $docs | ConvertTo-Json -Depth 10 | Set-Content -Path $using:jsonPath -Encoding Default
        return $docs
    }

    # Monitora o job em segundo plano
    MonitorarIndexacao
}

function MonitorarIndexacao {
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 500  # Verifica a cada 0,5 segundos
    $timer.Add_Tick({
        if ($global:indexJob -and $global:indexJob.State -ne "Running") {
            $timer.Stop()
            $timer.Dispose()

            if ($global:indexJob.State -eq "Completed") {
                $global:documentos = Receive-Job -Job $global:indexJob
                CarregarDocumentos
                $statusLabel.Text = "Catalogo feito."
            } else {
                $statusLabel.Text = "Erro ao catalogar: $($global:indexJob.State)"
            }

            $progressBar.Visible = $false
            $btnReindexar.Enabled = $true
            Remove-Job -Job $global:indexJob -Force
            $global:indexJob = $null
        } elseif ($global:indexJob) {
            # Simula progresso (não temos o número exato de arquivos, então usamos um ciclo)
            if ($progressBar.Value -lt 90) {
                $progressBar.Value += 5
            }
        }
    })
    $timer.Start()
}


function AbrirArquivo($caminho) {
    if ([string]::IsNullOrWhiteSpace($caminho) -or -not (Test-Path $caminho)) {
        return
    }
    try {
        Invoke-Item -Path $caminho
    } catch {
        Write-Host "Erro ao abrir arquivo: $_"
    }
}

function SelecionarPasta {
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Selecione uma pasta"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }
    return $null
}

function CarregarDocumentos {
    $textBoxPesquisa.Text = ""
    FiltrarDocumentos ""
}

function FiltrarDocumentos($filtro) {
    $global:listBox.Items.Clear()
    $documentos = CarregarDocumentosDoJson
    
    if ($filtro -eq "Pesquisar") {
        $filtroNormalizado = ""
    } else {
        $filtroNormalizado = $filtro.ToLower()
    }

    foreach ($doc in $documentos) {
        $nome = $doc.Nome.ToLower()
        $caminho = $doc.Caminho.ToLower()
        if ($nome.Contains($filtroNormalizado) -or $caminho.Contains($filtroNormalizado)) {
            $global:listBox.Items.Add("$($doc.Nome) | $($doc.Caminho)")
        }
    }
}


function GerirPastasRaiz {
    $janela = New-Object Windows.Forms.Form
    $janela.Text = "Gerir Pastas"
    $janela.Size = New-Object Drawing.Size(500, 400)
    $janela.MinimumSize = New-Object Drawing.Size(400, 300)
    $janela.StartPosition = "CenterScreen"
    AplicarEstiloJanela $janela

    $listPastas = New-Object Windows.Forms.ListBox
    $listPastas.Location = New-Object Drawing.Point(10, 10)
    $listPastas.Anchor = 'Top, Left, Right, Bottom'
    $listPastas.Width = $janela.ClientSize.Width - 20
    $listPastas.Height = 200
    AplicarEstiloListBox $listPastas
    $listPastas.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
    $janela.Controls.Add($listPastas)

    $btnAdicionar = New-Object Windows.Forms.Button
    $btnAdicionar.Text = "Adicionar Pasta"
    $btnAdicionar.Location = New-Object Drawing.Point(10, 220)
    $btnAdicionar.Font = New-Object System.Drawing.Font("Arial", 12)
    $btnAdicionar.Width = 140
    $btnAdicionar.Anchor = 'Bottom, Left'
    AplicarEstiloBotao $btnAdicionar
    CriarToolTip $btnAdicionar "Adiciona uma nova pasta na listagem de pastas para catalogar todos os arquivos dentro da pasta pretendida incluindo as subpastas."
    $btnAdicionar.Add_Click({
        $novaPasta = SelecionarPasta
        if ($novaPasta -and -not $listPastas.Items.Contains($novaPasta)) {
            $listPastas.Items.Add($novaPasta)
        }
    })
    $janela.Controls.Add($btnAdicionar)

    $btnRemover = New-Object Windows.Forms.Button
    $btnRemover.Text = "Remover Pasta"
    $btnRemover.Location = New-Object Drawing.Point(160, 220)
    $btnRemover.Font = New-Object System.Drawing.Font("Arial", 12)
    $btnRemover.Width = 140
    $btnRemover.Anchor = 'Bottom, Left'
    AplicarEstiloBotao $btnRemover
    CriarToolTip $btnRemover "Remove a pasta selecionada da listagem de pastas."
    $btnRemover.Add_Click({
        $selectedItems = @($listPastas.SelectedItems)
        foreach ($sel in $selectedItems) {
            $listPastas.Items.Remove($sel)
        }
    })
    $janela.Controls.Add($btnRemover)

    $btnGuardar = New-Object Windows.Forms.Button
    $btnGuardar.Text = "Guardar"
    $btnGuardar.Location = New-Object Drawing.Point(10, 260)
    $btnGuardar.Font = New-Object System.Drawing.Font("Arial", 12)
    $btnGuardar.Width = $janela.ClientSize.Width - 20
    $btnGuardar.Anchor = 'Bottom, Left, Right'
    AplicarEstiloBotao $btnGuardar
    CriarToolTip $btnGuardar "Guarda a listagem de pastas e cataloga os arquivos contidos dentro das pastas da listagem incluindo as subpastas."
    $btnGuardar.Add_Click({
        $global:rootFolders = @()
        foreach ($item in $listPastas.Items) {
            $global:rootFolders += [PSCustomObject]@{ Caminho = $item }
        }
        Save-RootFolders

        # Se a lista de pastas estiver vazia, limpa a lista de documentos
        if ($global:rootFolders.Count -eq 0) {
            $global:documentos = @()
            if (Test-Path $global:jsonPath) {
                Remove-Item $global:jsonPath -Force
            }
            CarregarDocumentos  # Atualiza a ListBox para refletir a lista vazia
        } else {
            IndexarArquivos  # Caso contrário, indexa normalmente
        }
        $janela.Close()
    })
    $janela.Controls.Add($btnGuardar)

    $janela.Add_Resize({
        $listPastas.Width = $janela.ClientSize.Width - 20
        $btnGuardar.Width = $janela.ClientSize.Width - 20
    })

    $listPastas.Items.Clear()
    $global:rootFolders = Load-RootFolders
    foreach ($pasta in $global:rootFolders) {
        $listPastas.Items.Add($pasta.Caminho)
    }

    AplicarEstiloJanela $janela
    $janela.ShowDialog()
}

### GUI PRINCIPAL
$form = New-Object Windows.Forms.Form
$form.Text = "Botica | Catalogo de ficheiros"
$form.Size = New-Object Drawing.Size(800, 520)
$form.MinimumSize = New-Object Drawing.Size(600, 400)
$form.StartPosition = "CenterScreen"
AplicarEstiloJanela $form

# TextBox de pesquisa
$textBoxPesquisa = New-Object Windows.Forms.TextBox
$textBoxPesquisa.Location = New-Object Drawing.Point(10, 10)
$textBoxPesquisa.Width = 760
$textBoxPesquisa.Height = 40
$textBoxPesquisa.Anchor = 'Top, Left, Right'
AplicarEstiloTextBox $textBoxPesquisa
$textBoxPesquisa.Text = "Pesquisar"
$textBoxPesquisa.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($textBoxPesquisa)

$textBoxPesquisa.Add_GotFocus({
    if ($textBoxPesquisa.Text -eq "Pesquisar") {
        $textBoxPesquisa.Text = ""
        $textBoxPesquisa.ForeColor = [System.Drawing.Color]::Black
    }
})

$textBoxPesquisa.Add_LostFocus({
    if ($textBoxPesquisa.Text -eq "") {
        $textBoxPesquisa.Text = "Pesquisar"
        $textBoxPesquisa.ForeColor = [System.Drawing.Color]::Gray
    }
})

$textBoxPesquisa.Add_TextChanged({
    FiltrarDocumentos $textBoxPesquisa.Text
})

# ListBox
$listBox = New-Object Windows.Forms.ListBox
$listBox.Location = New-Object Drawing.Point(10, 55)
$listBox.Size = New-Object Drawing.Size(760, 350)
$listBox.Anchor = 'Top, Bottom, Left, Right'
AplicarEstiloListBox $listBox
$listBox.HorizontalScrollbar = $true
$listBox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
$listBox.VirtualMode = $true
$listBox.Add_RetrieveVirtualItem({
    param($sender, $e)
    $index = $e.ItemIndex
    if ($index -lt $global:documentosFiltrados.Count) {
        $doc = $global:documentosFiltrados[$index]
        $e.Item = New-Object System.Windows.Forms.ListViewItem("$($doc.Nome) | $($doc.Caminho)")
    }
})
$listBox.Add_DoubleClick({
    foreach ($selectedItem in $listBox.SelectedItems) {
        if ($selectedItem -and $selectedItem -match '\|\s*(.+)$') {
            AbrirArquivo $matches[1]
        }
    }
})
$listBox.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        foreach ($selectedItem in $listBox.SelectedItems) {
            if ($selectedItem -and $selectedItem -match '\|\s*(.+)$') {
                AbrirArquivo $matches[1]
            }
        }
    }
})
$form.Controls.Add($listBox)
$global:listBox = $listBox

# Botões
$btnReindexar = New-Object Windows.Forms.Button
$btnReindexar.Text = "Atualizar"
$btnReindexar.Size = New-Object Drawing.Size(120, 30)
$btnReindexar.Location = New-Object Drawing.Point(10, 420)
$btnReindexar.Anchor = 'Bottom, Left'
AplicarEstiloBotao $btnReindexar
CriarToolTip $btnReindexar "Atualiza lista de arquivos catalogados conforme a listagem de pastas."
$btnReindexar.Add_Click({ IndexarArquivos })
$form.Controls.Add($btnReindexar)

$btnPastas = New-Object Windows.Forms.Button
$btnPastas.Text = "Pastas"
$btnPastas.Size = New-Object Drawing.Size(120, 30)
$btnPastas.Location = New-Object Drawing.Point(140, 420)
$btnPastas.Anchor = 'Bottom, Left'
AplicarEstiloBotao $btnPastas
CriarToolTip $btnPastas "Gerencia a listagem de pastas para catalogar os arquivos de interesse."
$btnPastas.Add_Click({ GerirPastasRaiz })
$form.Controls.Add($btnPastas)

$btnSair = New-Object Windows.Forms.Button
$btnSair.Text = "Sair"
$btnSair.Size = New-Object Drawing.Size(120, 30)
$btnSair.Location = New-Object Drawing.Point(270, 420)
$btnSair.Anchor = 'Bottom, Left'
AplicarEstiloBotao $btnSair
CriarToolTip $btnSair "Fecha o aplicativo."
$btnSair.Add_Click({ $form.Close() })
$form.Controls.Add($btnSair)

# Barra de Status
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = [System.Drawing.Color]::WhiteSmoke
$statusStrip.SizingGrip = $false
$statusStrip.Dock = 'Bottom'
$form.Controls.Add($statusStrip)

$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Pronto"
$statusLabel.Spring = $true  # Faz o label ocupar o espaço disponível
$statusStrip.Items.Add($statusLabel)

$progressBar = New-Object System.Windows.Forms.ToolStripProgressBar
$progressBar.Width = 200
$progressBar.Visible = $false
$statusStrip.Items.Add($progressBar)

# Evento de redimensionamento da janela principal
$form.Add_Resize({
    $textBoxPesquisa.Width = $form.ClientSize.Width - 20
    $listBox.Width = $form.ClientSize.Width - 20
    $listBox.Height = $form.ClientSize.Height - 110  # Ajusta a altura para a barra de status
})

CarregarDocumentos
$form.ShowDialog()