Add-Type -AssemblyName PresentationFramework

# Função para carregar documentos do arquivo CSV
function CarregarDocumentosDoCSV {
    $documentos = @()
    $csvPath = Join-Path -Path $PSScriptRoot -ChildPath "documentos.csv"
    if (Test-Path $csvPath) {
        $documentos = Import-Csv -Path $csvPath
    }
    return $documentos
}

# Função para salvar documento em arquivo CSV (adicionar)
function AdicionarDocumentoNoCSV {
    param($novoDocumento)
    $csvPath = Join-Path -Path $PSScriptRoot -ChildPath "documentos.csv"
    $novoDocumento | Export-Csv -Path $csvPath -NoTypeInformation -Append
    $global:documentos += $novoDocumento
}


# Função para salvar documentos em arquivo CSV (editar e apagar)
function SalvarDocumentosNoCSV {
    param($documentos)
    $csvPath = Join-Path -Path $PSScriptRoot -ChildPath "documentos.csv"
    $documentos | Export-Csv -Path $csvPath -NoTypeInformation
}

# Função para abrir um arquivo ou diretório com o aplicativo padrão
function AbrirArquivo {
    param($caminho)

    if (-not [string]::IsNullOrWhiteSpace($caminho)) {
        if (Test-Path $caminho) {
            if (Test-Path $caminho -PathType Leaf) {
                Start-Process $caminho
            }
            elseif (Test-Path $caminho -PathType Container) {
                Invoke-Item $caminho
            }
            else {
                Write-Host "O caminho especificado não corresponde a um arquivo ou diretório existente: $caminho"
            }
        }
        else {
            Write-Host "O arquivo ou diretório não existe: $caminho"
        }
    }
    else {
        Write-Host "Caminho do arquivo ou diretório não especificado."
    }
}


# Função para carregar todos os documentos na lista
function CarregarDocumentos {
    $listBox.Items.Clear()
    $global:documentos = CarregarDocumentosDoCSV
    foreach ($doc in $global:documentos) {
        $listBox.Items.Add($doc.Nome)
    }
}


# Função para filtrar documentos com base no texto de pesquisa
function FiltrarDocumentos {
    param($termoPesquisa)

    $listBox.Items.Clear()
    
    if ($documentos.Count -gt 0) {
        if ([string]::IsNullOrWhiteSpace($termoPesquisa) -or $termoPesquisa -eq "Digite aqui o que procura para filtrar o conteúdo") {
            # Se o termo de pesquisa estiver vazio, carrega todos os documentos
            CarregarDocumentos
        }
        else {
            foreach ($doc in $documentos) {
                if ($null -ne $doc -and ($doc.Nome -like "*$termoPesquisa*")) {
                    $listBox.Items.Add($doc.Nome)
                }
            }
        }
    }
}


# Função para adicionar documento
function AdicionarDocumento {
    $novoNome = Show-InputBox "Digite o nome do novo documento:"
    $novoCaminho = Show-InputBox "Digite o caminho do novo documento:"
    if (-not [string]::IsNullOrWhiteSpace($novoNome) -and -not [string]::IsNullOrWhiteSpace($novoCaminho)) {
        $novoDocumento = [PSCustomObject]@{
            Nome    = $novoNome
            Caminho = $novoCaminho
        }
        AdicionarDocumentoNoCSV $novoDocumento
        CarregarDocumentos
    }
    else {
        Write-Host "Nome ou caminho do documento não podem ser vazios."
    }
}

# Função para editar documento
function EditarDocumento {
    $indiceSelecionado = $listBox.SelectedIndex
    if ($indiceSelecionado -ge 0) {
        $nomeSelecionado = $listBox.SelectedItem
        $novoNome = Show-InputBox "Digite o novo nome do documento '$nomeSelecionado':"
        $novoCaminho = Show-InputBox "Digite o novo caminho do documento '$nomeSelecionado':"
        if (-not [string]::IsNullOrWhiteSpace($novoNome) -and -not [string]::IsNullOrWhiteSpace($novoCaminho)) {
            $documentoEditado = [PSCustomObject]@{
                Nome    = $novoNome
                Caminho = $novoCaminho
            }
            $documentos[$indiceSelecionado] = $documentoEditado
            SalvarDocumentosNoCSV $documentos
            CarregarDocumentos
        }
        else {
            Write-Host "Nome ou caminho do documento não podem ser vazios."
        }
    }
    else {
        Write-Host "Nenhum documento selecionado para editar."
    }
}
# Função para apagar documento
function ApagarDocumento {
    $indiceSelecionado = $listBox.SelectedIndex
    if ($indiceSelecionado -ge 0) {
        $documentoRemovido = $documentos[$indiceSelecionado]
        $documentos = $documentos | Where-Object { $_ -ne $documentoRemovido }
        SalvarDocumentosNoCSV $documentos
        CarregarDocumentos
    }
    else {
        Write-Host "Nenhum documento selecionado para apagar."
    }
}

# Função para exibir uma caixa de diálogo de entrada
function Show-InputBox {
    param([string]$message)
    $inputBox = New-Object -TypeName System.Windows.Forms.TextBox
    $form = New-Object -TypeName System.Windows.Forms.Form
    $form.Text = "Input Box"
    $form.Height = 150
    $form.Width = 300
    $label = New-Object -TypeName System.Windows.Forms.Label
    $label.Location = New-Object -TypeName System.Drawing.Point(10, 20)
    $label.Size = New-Object -TypeName System.Drawing.Size(280, 20)
    $label.Text = $message
    $form.Controls.Add($label)
    $inputBox.Location = New-Object -TypeName System.Drawing.Point(10, 50)
    $inputBox.Size = New-Object -TypeName System.Drawing.Size(260, 20)
    $form.Controls.Add($inputBox)
    $okButton = New-Object -TypeName System.Windows.Forms.Button
    $okButton.Location = New-Object -TypeName System.Drawing.Point(180, 80)
    $okButton.Size = New-Object -TypeName System.Drawing.Size(90, 30)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)
    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $form.Dispose()
        return $inputBox.Text
    }
    else {
        $form.Dispose()
        return $null
    }
}

# Lista de documentos
$documentos = CarregarDocumentosDoCSV

# Criar janela WPF
$window = New-Object System.Windows.Window
$window.Title = "Lab Assistant"
$window.Width = 500
$window.Height = 650

# Criar barra de ferramentas (toolbar)
$toolBar = New-Object System.Windows.Controls.ToolBar

# Botão para adicionar documento
$addButton = New-Object System.Windows.Controls.Button
$addButton.Content = "Adicionar"
$addButton.Add_Click({ AdicionarDocumento })
$toolBar.Items.Add($addButton)

# Botão para editar documento
$editButton = New-Object System.Windows.Controls.Button
$editButton.Content = "Editar"
$editButton.Add_Click({ EditarDocumento })
$toolBar.Items.Add($editButton)

# Botão para apagar documento
$deleteButton = New-Object System.Windows.Controls.Button
$deleteButton.Content = "Apagar"
$deleteButton.Add_Click({ ApagarDocumento })
$toolBar.Items.Add($deleteButton)

# Adicionar barra de ferramentas (dockbar) ao stackpanel
$stackPanel = New-Object System.Windows.Controls.StackPanel
$stackPanel.Orientation = [System.Windows.Controls.Orientation]::Vertical
$stackPanel.Children.Add($toolBar)

# Label para a barra de pesquisa
$searchLabel = New-Object System.Windows.Controls.Label
$searchLabel.Content = "Barra de pesquisa"
$searchLabel.FontWeight = "Bold"
$searchLabel.HorizontalContentAlignment = "Center"

# Criar barra de pesquisa
$searchBox = New-Object System.Windows.Controls.TextBox
$searchBox.Width = 300
$searchBox.Height = 25
$searchBox.Margin = "0,0,0,0"
$searchBox.Text = "Digite aqui o que procura para filtrar o conteúdo"  # Placeholder
$searchBox.Add_GotFocus({
        if ($searchBox.Text -eq "Digite aqui o que procura para filtrar o conteúdo") {
            $searchBox.Text = ""
        }
    })
$searchBox.Add_LostFocus({
        if ($searchBox.Text -eq "") {
            $searchBox.Text = "Digite aqui o que procura para filtrar o conteúdo"
        }
    })

# Botão de pesquisa
$searchButton = New-Object System.Windows.Controls.Button
$searchButton.Content = "Filtrar"
$searchButton.Width = 100
$searchButton.Margin = "5,10,0,0"
$searchButton.Add_Click({
        $termoPesquisa = $searchBox.Text.ToLower()
        FiltrarDocumentos $termoPesquisa
    })

# Label para a lista de documentos
$listBoxLabel = New-Object System.Windows.Controls.Label
$listBoxLabel.Content = "Resultados da pesquisa"
$listBoxLabel.FontWeight = "Bold"
$listBoxLabel.HorizontalContentAlignment = "Center"

# Criar lista de documentos
$listBox = New-Object System.Windows.Controls.ListBox
$listBox.Width = 450
$listBox.Height = 200
$listBox.Margin = "0,10,0,0"

# Evento MouseDoubleClick para abrir o arquivo selecionado ao dar duplo clique
$listBox.Add_MouseDoubleClick({
        $indiceSelecionado = $listBox.SelectedIndex
        if ($indiceSelecionado -ge 0) {
            $nomeSelecionado = $listBox.SelectedItem
            $caminho = ($documentos | Where-Object { $_.Nome -eq $nomeSelecionado }).Caminho
            AbrirArquivo $caminho
        }
        else {
            Write-Host "Nenhum arquivo selecionado."
        }
    })

# Evento KeyDown da barra de pesquisa
$searchBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq "Enter") {
            $termoPesquisa = $searchBox.Text.ToLower()
            FiltrarDocumentos $termoPesquisa
        }
    })

# Botão para abrir o arquivo selecionado
$button = New-Object System.Windows.Controls.Button
$button.Content = "Abrir"
$button.Width = 100
$button.Margin = "5,10,0,0"
$button.Add_Click({
        $indiceSelecionado = $listBox.SelectedIndex
        if ($indiceSelecionado -ge 0) {
            $nomeSelecionado = $listBox.SelectedItem
            $caminho = ($documentos | Where-Object { $_.Nome -eq $nomeSelecionado }).Caminho
            AbrirArquivo $caminho
        }
        else {
            Write-Host "Nenhum arquivo selecionado."
        }
    })

# Carregar todos os documentos na lista ao iniciar o programa
CarregarDocumentos

# Adicionar controles à janela
$stackPanel.Children.Add($searchLabel)
$stackPanel.Children.Add($searchBox)
$stackPanel.Children.Add($searchButton)
$stackPanel.Children.Add($listBoxLabel)
$stackPanel.Children.Add($listBox)
$stackPanel.Children.Add($button)

# Adicionar espaço entre o botão "Abrir" e a imagem
$spacer = New-Object System.Windows.Controls.Label
$spacer.Height = 20
$stackPanel.Children.Add($spacer)

# Adicionar imagem
# $image = New-Object System.Windows.Controls.Image
# $image.Source = [System.Windows.Media.Imaging.BitmapImage]::new([System.Uri]::new("H:\TE...

$window.Content = $stackPanel

# Mostrar janela
$window.ShowDialog() | Out-Null
