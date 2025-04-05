Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Variáveis globais
$global:documentos = @()
$global:csvPath = Join-Path -Path $PSScriptRoot -ChildPath "data.csv"
$global:rootFoldersPath = Join-Path -Path $PSScriptRoot -ChildPath "pastas.csv"
$global:listBox = $null
$global:rootFolders = @()

function Load-RootFolders {
    $global:rootFolders = @()
    if (Test-Path $global:rootFoldersPath) {
        $pastas = Import-Csv -Path $global:rootFoldersPath -Encoding UTF8
        $global:rootFolders = $pastas | Where-Object { $_.Caminho -and $_.Caminho -ne "" }
    }
    return $global:rootFolders
}

function Save-RootFolders {
    if ($global:rootFolders.Count -gt 0) {
        $global:rootFolders | Export-Csv -Path $global:rootFoldersPath -NoTypeInformation -Encoding UTF8
    } elseif (Test-Path $global:rootFoldersPath) {
        Remove-Item $global:rootFoldersPath -Force
    }
}

function CarregarDocumentosDoCSV {
    $global:documentos = @()
    if (Test-Path $global:csvPath) {
        $global:documentos = Import-Csv -Path $global:csvPath -Encoding UTF8
    }
    return $global:documentos
}

function Save-CSVData {
    $global:documentos | Export-Csv -Path $global:csvPath -NoTypeInformation -Encoding UTF8
}

function IndexarArquivos {
    $rootFolders = Load-RootFolders
    $existingDocs = CarregarDocumentosDoCSV
    $global:documentos = @()

    if ($rootFolders.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Nenhuma pasta raiz definida.", "Aviso", "OK", "Warning")
        Save-CSVData
        CarregarDocumentos
        return
    }

    foreach ($folder in $rootFolders) {
        if (Test-Path $folder.Caminho) {
            try {
                $files = Get-ChildItem -Path $folder.Caminho -Recurse -File -ErrorAction Stop
                foreach ($file in $files) {
                    $doc = [PSCustomObject]@{ Nome = $file.Name; Caminho = $file.FullName }
                    $global:documentos += $doc
                }
            } catch {
                Write-Host "Erro ao indexar: $($_.Exception.Message)"
            }
        }
    }

    Save-CSVData
    CarregarDocumentos
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
    $dialog.Description = "Selecione uma pasta raiz"
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
    $documentos = CarregarDocumentosDoCSV
    $filtroNormalizado = $filtro.ToLower()

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
    $janela.Text = "Gerir Pastas Raiz"
    $janela.Size = New-Object Drawing.Size(500, 400)
    $janela.StartPosition = "CenterScreen"
    $janela.BackColor = [System.Drawing.Color]::FromName('Window')  # Cor de fundo padrão do Windows

    $listPastas = New-Object Windows.Forms.ListBox
    $listPastas.Size = New-Object Drawing.Size(460, 200)
    $listPastas.Location = New-Object Drawing.Point(10, 10)
    $janela.Controls.Add($listPastas)

    $listPastas.Items.Clear()
    $global:rootFolders = Load-RootFolders
    foreach ($pasta in $global:rootFolders) {
        $listPastas.Items.Add($pasta.Caminho)
    }

    $btnAdicionar = New-Object Windows.Forms.Button
    $btnAdicionar.Text = "Adicionar Pasta"
    $btnAdicionar.Size = New-Object Drawing.Size(140, 30)
    $btnAdicionar.Location = New-Object Drawing.Point(10, 230)
    $btnAdicionar.BackColor = [System.Drawing.Color]::FromName('ButtonFace')  # Cor do botão padrão
    $btnAdicionar.Add_Click({
        $novaPasta = SelecionarPasta
        if ($novaPasta -and -not $listPastas.Items.Contains($novaPasta)) {
            $listPastas.Items.Add($novaPasta)
        }
    })
    $janela.Controls.Add($btnAdicionar)

    $btnRemover = New-Object Windows.Forms.Button
    $btnRemover.Text = "Remover Selecionada"
    $btnRemover.Size = New-Object Drawing.Size(140, 30)
    $btnRemover.Location = New-Object Drawing.Point(170, 230)
    $btnRemover.BackColor = [System.Drawing.Color]::FromName('ButtonFace')  # Cor do botão padrão
    $btnRemover.Add_Click({
        $sel = $listPastas.SelectedItem
        if ($sel) {
            $listPastas.Items.Remove($sel)
        }
    })
    $janela.Controls.Add($btnRemover)

    $btnGuardar = New-Object Windows.Forms.Button
    $btnGuardar.Text = "Guardar Alterações"
    $btnGuardar.Size = New-Object Drawing.Size(460, 30)
    $btnGuardar.Location = New-Object Drawing.Point(10, 280)
    $btnGuardar.BackColor = [System.Drawing.Color]::FromName('ButtonFace')  # Cor do botão padrão
    $btnGuardar.Add_Click({
        $global:rootFolders = @()
        foreach ($item in $listPastas.Items) {
            $global:rootFolders += [PSCustomObject]@{ Caminho = $item }
        }
        Save-RootFolders
        IndexarArquivos
        $janela.Close()
    })
    $janela.Controls.Add($btnGuardar)

    $janela.ShowDialog()
}

# GUI principal
$form = New-Object Windows.Forms.Form
$form.Text = "Indexador de Documentos"
$form.Size = New-Object Drawing.Size(800, 520)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromName('Window')  # Cor de fundo padrão do Windows

# Barra de pesquisa (Aumentada)
$textBoxPesquisa = New-Object Windows.Forms.TextBox
$textBoxPesquisa.Size = New-Object Drawing.Size(760, 40)  # Aumentei a altura da TextBox
$textBoxPesquisa.Location = New-Object Drawing.Point(10, 10)
$textBoxPesquisa.Anchor = "Top, Left, Right"
$textBoxPesquisa.BackColor = [System.Drawing.Color]::FromName('Window')  # Cor de fundo padrão do Windows
$textBoxPesquisa.ForeColor = [System.Drawing.Color]::FromName('WindowText')  # Cor do texto padrão
$textBoxPesquisa.Font = New-Object Drawing.Font("Arial", 14)  # Aumentei o tamanho da fonte
$textBoxPesquisa.Add_TextChanged({
    FiltrarDocumentos $textBoxPesquisa.Text
})
$form.Controls.Add($textBoxPesquisa)

# ListBox
$global:listBox = New-Object Windows.Forms.ListBox
$global:listBox.Size = New-Object Drawing.Size(760, 260)
$global:listBox.Location = New-Object Drawing.Point(10, 55)  # Ajustei para dar espaço maior
$global:listBox.HorizontalScrollbar = $true
$global:listBox.BackColor = [System.Drawing.Color]::FromName('Window')  # Cor de fundo padrão do Windows
$global:listBox.ForeColor = [System.Drawing.Color]::FromName('WindowText')  # Cor do texto padrão

# Duplo clique para abrir documento
$global:listBox.Add_DoubleClick({
    $selecionado = $global:listBox.SelectedItem
    if ($selecionado) {
        $caminho = $selecionado -split '\|', 2
        if ($caminho.Count -eq 2) {
            AbrirArquivo($caminho[1].Trim())
        }
    }
})

$form.Controls.Add($global:listBox)

# Botões
$btnIndexar = New-Object Windows.Forms.Button
$btnIndexar.Text = "Indexar Arquivos"
$btnIndexar.Size = New-Object Drawing.Size(120, 30)
$btnIndexar.Location = New-Object Drawing.Point(10, 360)
$btnIndexar.BackColor = [System.Drawing.Color]::FromName('ButtonFace')  # Cor do botão padrão
$btnIndexar.Add_Click({ IndexarArquivos })
$form.Controls.Add($btnIndexar)

$btnAbrir = New-Object Windows.Forms.Button
$btnAbrir.Text = "Abrir Arquivo"
$btnAbrir.Size = New-Object Drawing.Size(120, 30)
$btnAbrir.Location = New-Object Drawing.Point(140, 360)
$btnAbrir.BackColor = [System.Drawing.Color]::FromName('ButtonFace')  # Cor do botão padrão
$btnAbrir.Add_Click({
    if ($listBox.SelectedItem) {
        $linha = $listBox.SelectedItem
        $partes = $linha -split '\|'
        $caminho = $partes[1].Trim()
        AbrirArquivo $caminho
    }
})
$form.Controls.Add($btnAbrir)

$btnPastas = New-Object Windows.Forms.Button
$btnPastas.Text = "Gerir Pastas"
$btnPastas.Size = New-Object Drawing.Size(120, 30)
$btnPastas.Location = New-Object Drawing.Point(270, 360)
$btnPastas.BackColor = [System.Drawing.Color]::FromName('ButtonFace')  # Cor do botão padrão
$btnPastas.Add_Click({ GerirPastasRaiz })
$form.Controls.Add($btnPastas)

$btnFechar = New-Object Windows.Forms.Button
$btnFechar.Text = "Fechar"
$btnFechar.Size = New-Object Drawing.Size(120, 30)
$btnFechar.Location = New-Object Drawing.Point(400, 360)
$btnFechar.BackColor = [System.Drawing.Color]::FromName('ButtonFace')  # Cor do botão padrão
$btnFechar.Add_Click({ $form.Close() })
$form.Controls.Add($btnFechar)

# Exibir a janela
IndexarArquivos
$form.ShowDialog()
