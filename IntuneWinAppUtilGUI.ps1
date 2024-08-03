# Adicionar montagem de Windows Forms

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web.Extensions

# Adicionando um ícone personalizado ao formulário

$iconPath = "https://vhdxgabrielluiz.blob.core.windows.net/vhdx/gabrielluiz-icone.ico"
$iconTempPath = [System.IO.Path]::GetTempFileName() + ".ico"
$webClient = New-Object System.Net.WebClient
$webClient.DownloadFile($iconPath, $iconTempPath) # Baixe o ícone e salve-o temporariamente
$icon = New-Object System.Drawing.Icon($iconTempPath)

# Função para selecionar uma pasta

function Select-FolderDialog {
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    }
    return $null
}

# Função para selecionar um arquivo

function Select-FileDialog {
    $fileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    if ($fileBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $fileBrowser.FileName
    }
    return $null
}

# Função para exibir a janela de progresso

function Show-ProgressForm {
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Gerando o arquivo .intunewin"
    $progressForm.Size = New-Object System.Drawing.Size(400,100)
    $progressForm.StartPosition = "CenterScreen"
    $progressForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $progressForm.ControlBox = $false

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    $progressBar.Dock = [System.Windows.Forms.DockStyle]::Fill
    $progressForm.Controls.Add($progressBar)

    return $progressForm
}

# Função para salvar os dados em um arquivo JSON

function Save-FormData {
    $formData = @{
        ExePath = $textBoxExe.Text
        SourcePath = $textBoxSource.Text
        FilePath = $textBoxFile.Text
        OutputPath = $textBoxOutput.Text
    }
    $json = [System.Web.Script.Serialization.JavaScriptSerializer]::new().Serialize($formData)
    Set-Content -Path "formData.json" -Value $json
}

# Função para carregar os dados de um arquivo JSON

function Load-FormData {
    if (Test-Path "formData.json") {
        $json = Get-Content -Path "formData.json" -Raw
        $formData = [System.Web.Script.Serialization.JavaScriptSerializer]::new().DeserializeObject($json)
        $textBoxExe.Text = $formData.ExePath
        $textBoxSource.Text = $formData.SourcePath
        $textBoxFile.Text = $formData.FilePath
        $textBoxOutput.Text = $formData.OutputPath
    }
}

# Função para limpar os campos do formulário

function Clear-FormData {
    $textBoxExe.Text = ""
    $textBoxSource.Text = ""
    $textBoxFile.Text = ""
    $textBoxOutput.Text = ""
}

# Criar o formulário principal

$form = New-Object System.Windows.Forms.Form
$form.Text = "IntuneWinAppUtil GUI"
$form.Size = New-Object System.Drawing.Size(620,350)
$form.StartPosition = "CenterScreen"
$form.Icon = $icon

# Caixa de texto para o caminho do IntuneWinAppUtil.exe

$labelExe = New-Object System.Windows.Forms.Label
$labelExe.Text = "IntuneWinAppUtil.exe:"
$labelExe.Location = New-Object System.Drawing.Point(10,20)
$form.Controls.Add($labelExe)

$textBoxExe = New-Object System.Windows.Forms.TextBox
$textBoxExe.Size = New-Object System.Drawing.Size(350,20)
$textBoxExe.Location = New-Object System.Drawing.Point(150,20)
$form.Controls.Add($textBoxExe)

$buttonExe = New-Object System.Windows.Forms.Button
$buttonExe.Text = "Selecionar"
$buttonExe.Location = New-Object System.Drawing.Point(510,18)
$buttonExe.Size = New-Object System.Drawing.Size(75,23)
$buttonExe.Add_Click({
    $exePath = Select-FileDialog
    if ($exePath) {
        $textBoxExe.Text = $exePath
    }
})
$form.Controls.Add($buttonExe)

# Caixa de texto para o caminho da pasta de origem

$labelSource = New-Object System.Windows.Forms.Label
$labelSource.Text = "Pasta de Origem:"
$labelSource.Location = New-Object System.Drawing.Point(10,60)
$form.Controls.Add($labelSource)

$textBoxSource = New-Object System.Windows.Forms.TextBox
$textBoxSource.Size = New-Object System.Drawing.Size(350,20)
$textBoxSource.Location = New-Object System.Drawing.Point(150,60)
$form.Controls.Add($textBoxSource)

$buttonSource = New-Object System.Windows.Forms.Button
$buttonSource.Text = "Selecionar"
$buttonSource.Location = New-Object System.Drawing.Point(510,58)
$buttonSource.Size = New-Object System.Drawing.Size(75,23)
$buttonSource.Add_Click({
    $folderPath = Select-FolderDialog
    if ($folderPath) {
        $textBoxSource.Text = $folderPath
    }
})
$form.Controls.Add($buttonSource)

# Caixa de texto para o caminho do arquivo de origem

$labelFile = New-Object System.Windows.Forms.Label
$labelFile.Text = "Arquivo de instalação:"
$labelFile.Location = New-Object System.Drawing.Point(10,100)
$form.Controls.Add($labelFile)

$textBoxFile = New-Object System.Windows.Forms.TextBox
$textBoxFile.Size = New-Object System.Drawing.Size(350,20)
$textBoxFile.Location = New-Object System.Drawing.Point(150,100)
$form.Controls.Add($textBoxFile)

$buttonFile = New-Object System.Windows.Forms.Button
$buttonFile.Text = "Selecionar"
$buttonFile.Location = New-Object System.Drawing.Point(510,98)
$buttonFile.Size = New-Object System.Drawing.Size(75,23)
$buttonFile.Add_Click({
    $filePath = Select-FileDialog
    if ($filePath) {
        $textBoxFile.Text = $filePath
    }
})
$form.Controls.Add($buttonFile)

# Caixa de texto para o caminho da pasta de saída

$labelOutput = New-Object System.Windows.Forms.Label
$labelOutput.Text = "Pasta de Saída:"
$labelOutput.Location = New-Object System.Drawing.Point(10,140)
$form.Controls.Add($labelOutput)

$textBoxOutput = New-Object System.Windows.Forms.TextBox
$textBoxOutput.Size = New-Object System.Drawing.Size(350,20)
$textBoxOutput.Location = New-Object System.Drawing.Point(150,140)
$form.Controls.Add($textBoxOutput)

$buttonOutput = New-Object System.Windows.Forms.Button
$buttonOutput.Text = "Selecionar"
$buttonOutput.Location = New-Object System.Drawing.Point(510,138)
$buttonOutput.Size = New-Object System.Drawing.Size(75,23)
$buttonOutput.Add_Click({
    $outputPath = Select-FolderDialog
    if ($outputPath) {
        $textBoxOutput.Text = $outputPath
    }
})
$form.Controls.Add($buttonOutput)

# Botão para executar o comando

$buttonExecute = New-Object System.Windows.Forms.Button
$buttonExecute.Text = "Executar"
$buttonExecute.Location = New-Object System.Drawing.Point(200,180)
$buttonExecute.Size = New-Object System.Drawing.Size(75,23)
$buttonExecute.Add_Click({
    $exePath = $textBoxExe.Text
    $sourcePath = $textBoxSource.Text
    $filePath = $textBoxFile.Text
    $outputPath = $textBoxOutput.Text
    if ($exePath -and $sourcePath -and $filePath -and $outputPath) {
        $intunewinFile = Join-Path $outputPath ([System.IO.Path]::GetFileNameWithoutExtension($filePath) + ".intunewin")
        if (Test-Path $intunewinFile) {
            $result = [System.Windows.Forms.MessageBox]::Show("O arquivo $intunewinFile já existe e será deletado.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
                return
            }
            Remove-Item $intunewinFile -ErrorAction Stop
        }

        $progressForm = Show-ProgressForm
        $progressForm.Show()

        $job = Start-Job -ScriptBlock {
            param ($exePath, $sourcePath, $filePath, $outputPath)
            & $exePath -c $sourcePath -s $filePath -o $outputPath
        } -ArgumentList $exePath, $sourcePath, $filePath, $outputPath

        $job | Wait-Job

        $progressForm.Close()

        if (Test-Path $intunewinFile) {
            [System.Windows.Forms.MessageBox]::Show("Arquivo .intunewin gerado com sucesso!", "Sucesso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("Falha ao gerar o arquivo .intunewin.", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
        Save-FormData
    } else {
        [System.Windows.Forms.MessageBox]::Show("Por favor, preencha todos os campos.", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$form.Controls.Add($buttonExecute)


# Botão para limpar os campos do formulário

$buttonClear = New-Object System.Windows.Forms.Button
$buttonClear.Text = "Limpar"
$buttonClear.Location = New-Object System.Drawing.Point(320,180)
$buttonClear.Size = New-Object System.Drawing.Size(75,23)
$buttonClear.Add_Click({
    Clear-FormData
})
$form.Controls.Add($buttonClear)


# Botão para download do IntuneWinAppUtil.exe

$buttonDownload = New-Object System.Windows.Forms.Button
$buttonDownload.Text = "Download IntuneWinAppUtil.exe"
$buttonDownload.Location = New-Object System.Drawing.Point(200,210) # Ajuste a posição conforme necessário
$buttonDownload.Size = New-Object System.Drawing.Size(200,23)
$buttonDownload.Add_Click({
    Start-Process "https://raw.githubusercontent.com/microsoft/Microsoft-Win32-Content-Prep-Tool/master/IntuneWinAppUtil.exe"
})
$form.Controls.Add($buttonDownload)


# Adicionar a imagem abaixo do botão de download

try {
    $imageURL = "https://vhdxgabrielluiz.blob.core.windows.net/vhdx/perfil.png"
    $iconImage = [System.Drawing.Image]::FromStream((New-Object System.Net.WebClient).OpenRead($imageURL))
    $pictureBox = New-Object System.Windows.Forms.PictureBox
    $pictureBox.Image = $iconImage
    $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
    $pictureBox.Size = New-Object System.Drawing.Size(64, 64)
    $pictureBox.Location = New-Object System.Drawing.Point(257, 240) # Ajuste a posição conforme necessário
    $pictureBox.Cursor = [System.Windows.Forms.Cursors]::Hand
    $pictureBox.Add_Click({
        Start-Process "https://gabrielluiz.com"
    })
    $form.Controls.Add($pictureBox)
} catch {
    Write-Host "Não foi possível carregar a imagem: $_"
}

# Carregar os dados do formulário ao iniciar

$form.Add_Shown({
    Load-FormData
    $form.Activate()
})

# Executar o formulário

[System.Windows.Forms.Application]::Run($form)
