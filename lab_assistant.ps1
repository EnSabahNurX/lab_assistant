# Adicionar os assemblies necessários
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Função para carregar documentos do arquivo CSV
function CarregarDocumentosDoCSV {
    $documentos = @()
    $csvPath = Join-Path -Path $PSScriptRoot -ChildPath "documentos.csv"
    if (Test-Path $csvPath) {
        # Tentar importar o arquivo CSV com codificação Default
        try {
            $documentos = Import-Csv -Path $csvPath -Encoding Default
        }
        catch {
            Write-Host "Erro ao carregar o arquivo CSV: $_"
        }
    }
    return $documentos
}

# Função para salvar os dados no arquivo CSV com codificação UTF-8
function Save-CSVData {
    $script:global:csvData | Export-Csv -Path documentos.csv -NoTypeInformation -Encoding UTF8
}

# Função para carregar os dados do arquivo CSV com codificação UTF-8
function Load-CSVData {
    $csvData = @()
    $csvPath = Join-Path -Path $PSScriptRoot -ChildPath "documentos.csv"
    if (Test-Path $csvPath) {
        # Tentar importar o arquivo CSV com codificação UTF-8
        try {
            $csvData = Import-Csv -Path $csvPath -Encoding UTF8
        }
        catch {
            Write-Host "Erro ao carregar o arquivo CSV: $_"
        }
    }
    return $csvData
}

# Função para abrir um arquivo ou diretório com o aplicativo padrão
function AbrirArquivo {
    param($caminho)

    if (-not [string]::IsNullOrWhiteSpace($caminho)) {
        try {
            if (Test-Path $caminho -PathType Leaf) {
                Start-Process -FilePath $caminho
            }
            elseif (Test-Path $caminho -PathType Container) {
                Invoke-Item -Path $caminho
            }
            else {
                Write-Host "Caminho inválido ou não encontrado: $caminho"
            }
        }
        catch {
            Write-Host "Erro ao abrir o arquivo ou diretório: $_"
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
        # Converter o nome do documento para UTF-8
        $utf8Nome = [System.Text.Encoding]::UTF8.GetBytes($doc.Nome)
        $utf8Nome = [System.Text.Encoding]::UTF8.GetString($utf8Nome)

        # Adicionar o documento à lista usando a codificação UTF-8
        $listBox.Items.Add($utf8Nome)
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

# Função para selecionar um arquivo ou pasta
function SelecionarArquivoOuPasta {
    param([string]$titulo, [bool]$selecionarArquivo)

    $dialog = New-Object Microsoft.Win32.OpenFileDialog

    if (-not $selecionarArquivo) {
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    }

    $resultado = $dialog.ShowDialog()

    if ($resultado -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }

    return $null
}

# Função para selecionar o arquivo ou pasta e atualizar o DataGrid
function Selecionar-ArquivoOuPasta {    
    $row = $datagrid.SelectedItem
    if ($row -ne $null) {
        if ([string]::IsNullOrEmpty($row.Nome)) {
            [System.Windows.MessageBox]::Show("Por favor, preencha primeiro a coluna 'Nome' e carregue em Enter antes de selecionar o arquivo ou pasta.", "Atenção", "OK", "Warning")
        } else {
            $caminho = SelecionarArquivoOuPasta "Selecionar arquivo ou pasta" $true
            if ($caminho -ne $null) {
                $row.Caminho = $caminho
                $datagrid.Items.Refresh()
            }
        }
    }
}

# Função para criar a GUI do editor de CSV
function Show-GUI {
    $window = New-Object System.Windows.Window
    $window.Title = "Editor de catálogo de documentos"
    
    $toolbar = New-Object System.Windows.Controls.ToolBar

    $buttonAdd = New-Object System.Windows.Controls.Button
    $buttonAdd.Content = "+ Adicionar"
    $buttonAdd.FontSize = 24
    $buttonAdd.VerticalContentAlignment = "Center"
    $buttonAdd.ToolTip = "Adicionar novo item"
    $buttonAdd.Add_Click({
            $newRow = New-Object PSObject -Property @{
                Nome               = ""
                Caminho            = ""
                "Caminho completo" = ""
            }
            $script:global:csvData += $newRow
            $datagrid.ItemsSource = $script:global:csvData
        })
    $buttonAdd.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0) # Adiciona margem à direita

    $buttonDelete = New-Object System.Windows.Controls.Button
    $buttonDelete.Content = "- Apagar"
    $buttonDelete.FontSize = 24
    $buttonDelete.VerticalContentAlignment = "Center"
    $buttonDelete.ToolTip = "Apagar item selecionado"
    $buttonDelete.Add_Click({
            if ($null -ne $datagrid.SelectedItem) {
                $result = [System.Windows.MessageBox]::Show("Tem certeza de que deseja apagar este item?", "Confirmação", "YesNo", "Warning") # Melhorando o tipo de mensagem para 'Warning'
                if ($result -eq "Yes") {
                    $script:global:csvData = $script:global:csvData | Where-Object { $_ -ne $datagrid.SelectedItem }
                    $datagrid.ItemsSource = $script:global:csvData
                    Save-CSVData
                }
            }
        })
    $buttonDelete.Margin = New-Object System.Windows.Thickness(0, 0, 20, 0) # Adiciona margem à direita

    $toolbar.Items.Add($buttonAdd)
    $toolbar.Items.Add($buttonDelete)

    $grid = New-Object System.Windows.Controls.Grid

    $datagrid = New-Object System.Windows.Controls.DataGrid
    $datagrid.AutoGenerateColumns = $false
    $datagrid.ItemsSource = $script:global:csvData

    # Coluna de Nome
    $nomeColumn = New-Object System.Windows.Controls.DataGridTextColumn
    $nomeColumn.Header = "Nome"
    $nomeColumn.Binding = New-Object System.Windows.Data.Binding("Nome")
    $datagrid.Columns.Add($nomeColumn)
    
    # Coluna de Caminho com botão para selecionar arquivo/pasta
    $caminhoColumn = New-Object System.Windows.Controls.DataGridTemplateColumn
    $caminhoColumn.Header = "Caminho"
    $caminhoColumn.CellTemplate = New-Object System.Windows.DataTemplate
    $caminhoFactory = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.Button])
    $caminhoFactory.Name = "SelectButton"
    $caminhoFactory.SetValue([System.Windows.Controls.Button]::ContentProperty, "Selecionar")

    # Manipulador de evento para o botão "Selecionar" na coluna "Caminho"
    $caminhoFactory.AddHandler([System.Windows.Controls.Button]::ClickEvent, [System.Windows.RoutedEventHandler] {
            Selecionar-ArquivoOuPasta
        })

    $caminhoColumn.CellTemplate.VisualTree = $caminhoFactory
    $datagrid.Columns.Add($caminhoColumn)

    # Coluna de Caminho completo (vinculada à coluna "Caminho")
    $nomeColumnLocal = New-Object System.Windows.Controls.DataGridTextColumn
    $nomeColumnLocal.Header = "Caminho completo"
    $nomeColumnLocal.Binding = New-Object System.Windows.Data.Binding("Caminho")
    $datagrid.Columns.Add($nomeColumnLocal)
    
    # Evento para salvar automaticamente ao perder o foco da célula
    $datagrid.Add_LostFocus({
            # Finalizar qualquer edição em andamento no DataGrid
            if ($datagrid.IsEditing) {
                $datagrid.CommitEdit()
            }
            if ($datagrid.IsAddingNew) {
                $datagrid.CommitEdit()
            }

            # Salvar automaticamente as alterações
            Save-CSVData
        })

    # Adicionar evento KeyDown para o DataGrid
    $datagrid.Add_KeyDown({
            param($sender, $e)
            if ($e.Key -eq "Delete") {
                if ($null -ne $datagrid.SelectedItem) {
                    $result = [System.Windows.MessageBox]::Show("Tem certeza de que deseja apagar este item?", "Confirmação", "YesNo", "Warning") # Melhorando o tipo de mensagem para 'Warning'
                    if ($result -eq "Yes") {
                        $script:global:csvData = $script:global:csvData | Where-Object { $_ -ne $datagrid.SelectedItem }
                        $datagrid.ItemsSource = $script:global:csvData
                        Save-CSVData
                    }
                }
            }
        })

    $grid.Children.Add($datagrid)

    $dockPanel = New-Object System.Windows.Controls.DockPanel
    [System.Windows.Controls.DockPanel]::SetDock($toolbar, "Top")
    [System.Windows.Controls.DockPanel]::SetDock($grid, "Bottom")
    $dockPanel.Children.Add($toolbar)
    $dockPanel.Children.Add($grid)

    $window.Content = $dockPanel
    $window.ShowDialog() | Out-Null
}

# Função para abrir a janela de edição de CSV ao clicar em "Ajustes"
function OpenCSVEditor {
    # Carregar os dados do arquivo CSV
    $script:global:csvData = Load-CSVData

    # Mostrar a GUI do editor de CSV
    Show-GUI

    # Salvar os dados no arquivo CSV ao fechar a janela de edição
    Save-CSVData

    # Carregar todos os documentos na lista ao iniciar o programa
    CarregarDocumentos
}

# Lista de documentos
$documentos = CarregarDocumentosDoCSV

# Criar janela WPF
$window = New-Object System.Windows.Window
$window.Title = "Lab Assistant"
$window.Width = 500
$window.Height = 650

# Criar stack panel
$stackPanel = New-Object System.Windows.Controls.StackPanel
$stackPanel.Orientation = [System.Windows.Controls.Orientation]::Vertical

# Criar barra de ferramentas (dockbar)
$toolBar = New-Object System.Windows.Controls.ToolBar
$toolBar.Margin = "0,0,0,5"

# Botão de ajustes
$btnAjustes = New-Object System.Windows.Controls.Button
$btnAjustes.Content = "Ajustes"
$btnAjustes.ToolTip = "Abrir editor de CSV"
$btnAjustes.Add_Click({
        OpenCSVEditor
    })

# Adicionar botão à barra de ferramentas
$toolBar.Items.Add($btnAjustes)

# Adicionar barra de ferramentas ao stack panel
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
$searchButton.ToolTip = "Filtrar documentos"
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
$button.ToolTip = "Abrir documento selecionado"
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

# Adicionar controles à janela
$stackPanel.Children.Add($searchLabel)
$stackPanel.Children.Add($searchBox)
$stackPanel.Children.Add($searchButton)
$stackPanel.Children.Add($listBoxLabel)
$stackPanel.Children.Add($listBox)
$stackPanel.Children.Add($button)

# Carregar todos os documentos na lista ao iniciar o programa
CarregarDocumentos

# Adicionar espaço entre o botão "Abrir" e a imagem
$spacer = New-Object System.Windows.Controls.Label
$spacer.Height = 20
$stackPanel.Children.Add($spacer)

# Adicionar imagem
$image = New-Object System.Windows.Controls.Image
$imagePath = Join-Path -Path $PSScriptRoot -ChildPath "image.jpg"
$image.Source = [System.Windows.Media.Imaging.BitmapImage]::new([System.Uri]::new($imagePath))
$stackPanel.Children.Add($image)

$window.Content = $stackPanel

# Mostrar janela
$window.ShowDialog() | Out-Null
