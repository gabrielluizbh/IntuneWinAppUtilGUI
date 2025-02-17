# Ensure the script is compatible with PowerShell 7
if ($PSVersionTable.PSVersion.Major -ge 7) {
    # Load necessary assemblies
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
}

# Define language options
$languages = @{
    "en"    = @{
        "Title"            = "IntuneWinAppUtil GUI"
        "ExeLabel"         = "IntuneWinAppUtil.exe:"
        "SourceLabel"      = "Source Folder:"
        "FileLabel"        = "Installation File:"
        "OutputLabel"      = "Output Folder:"
        "Select"           = "Select"
        "Execute"          = "Execute"
        "Clear"            = "Clear"
        "Download"         = "Download IntuneWinAppUtil.exe"
        "SuccessMessage"   = "The .intunewin file was generated successfully!"
        "ErrorMessage"     = "Failed to generate the .intunewin file."
        "WarningMessage"   = "The file already exists and will be deleted."
        "WarningMessage1"  = "Warning"
        "WarningMessage2"  = "Success"
        "ErrorMessage1"    = "Error"
        "AllFieldsWarning" = "Please, enter all fields required."
        "WarningMessage3"  = "It could not load the image:"
    }
    "pt-BR" = @{
        "Title"            = "IntuneWinAppUtil GUI"
        "ExeLabel"         = "IntuneWinAppUtil.exe:"
        "SourceLabel"      = "Pasta de Origem:"
        "FileLabel"        = "Arquivo de instalação:"
        "OutputLabel"      = "Pasta de Saída:"
        "Select"           = "Selecionar"
        "Execute"          = "Executar"
        "Clear"            = "Limpar"
        "Download"         = "Download IntuneWinAppUtil.exe"
        "SuccessMessage"   = "Arquivo .intunewin gerado com sucesso!"
        "ErrorMessage"     = "Falha ao gerar o arquivo .intunewin."
        "WarningMessage"   = "O arquivo já existe e será deletado."
        "WarningMessage1"  = "Aviso"
        "WarningMessage2"  = "Successo"
        "ErrorMessage1"    = "Erro"
        "AllFieldsWarning" = "Por favor, preencha todos os campos."
        "WarningMessage3"  = "Não foi possível carregar a imagem:"
    }
    "fr"    = @{
        "Title"            = "IntuneWinAppUtil GUI"
        "ExeLabel"         = "IntuneWinAppUtil.exe:"
        "SourceLabel"      = "Dossier Source:"
        "FileLabel"        = "Fichier d'installation:"
        "OutputLabel"      = "Dossier de sortie:"
        "Select"           = "Sélectionner"
        "Execute"          = "Exécuter"
        "Clear"            = "Effacer"
        "Download"         = "Télécharger IntuneWinAppUtil.exe"
        "SuccessMessage"   = "Le fichier .intunewin a été généré avec succès!"
        "ErrorMessage"     = "Échec de la génération du fichier .intunewin."
        "WarningMessage"   = "Le fichier existe déjà et sera supprimé."
        "WarningMessage1"  = "Avertissement"
        "WarningMessage2"  = "Succès"
        "ErrorMessage1"    = "Erreur"
        "AllFieldsWarning" = "Veuillez remplir tous les champs requis."
        "WarningMessage3"  = "Impossible de charger l'image:"
    }
    "es"    = @{
        "Title"            = "IntuneWinAppUtil GUI"
        "ExeLabel"         = "IntuneWinAppUtil.exe:"
        "SourceLabel"      = "Carpeta de origen:"
        "FileLabel"        = "Archivo de instalación:"
        "OutputLabel"      = "Carpeta de salida:"
        "Select"           = "Seleccionar"
        "Execute"          = "Ejecutar"
        "Clear"            = "Limpiar"
        "Download"         = "Descargar IntuneWinAppUtil.exe"
        "SuccessMessage"   = "¡El archivo .intunewin se generó correctamente!"
        "ErrorMessage"     = "Error al generar el archivo .intunewin."
        "WarningMessage"   = "El archivo ya existe y se eliminará."
        "WarningMessage1"  = "Advertencia"
        "WarningMessage2"  = "Éxito"
        "ErrorMessage1"    = "Error"
        "AllFieldsWarning" = "Por favor, complete todos los campos requeridos."
        "WarningMessage3"  = "No se pudo cargar la imagen:"
    }
}

# Create the language selection form
$langForm = New-Object System.Windows.Forms.Form
$langForm.Text = "Select Language"
$langForm.Size = New-Object System.Drawing.Size(300, 150)
$langForm.StartPosition = "CenterScreen"

# Add ComboBox for language selection
$comboBoxLanguage = New-Object System.Windows.Forms.ComboBox
$comboBoxLanguage.Size = New-Object System.Drawing.Size(200, 20)
$comboBoxLanguage.Location = New-Object System.Drawing.Point(50, 40)
$comboBoxLanguage.Items.AddRange(@("English", "Português", "Français", "Español"))
$comboBoxLanguage.SelectedIndex = 0  # Default selection (English)
$langForm.Controls.Add($comboBoxLanguage)

# Add OK Button
$buttonOK = New-Object System.Windows.Forms.Button
$buttonOK.Text = "OK"
$buttonOK.Size = New-Object System.Drawing.Size(75, 23)
$buttonOK.Location = New-Object System.Drawing.Point(110, 80)
$langForm.Controls.Add($buttonOK)

# Declare the languageSelected variable globally
$global:languageSelected = $null

# Add Button Click event to capture language selection
$buttonOK.Add_Click({
        # Set the selected language based on ComboBox selection
        if ($comboBoxLanguage.SelectedItem -eq "English") {
            $global:languageSelected = "en"
        }
        elseif ($comboBoxLanguage.SelectedItem -eq "Português") {
            $global:languageSelected = "pt-BR"
        }
        elseif ($comboBoxLanguage.SelectedItem -eq "Français") {
            $global:languageSelected = "fr"
        }
        elseif ($comboBoxLanguage.SelectedItem -eq "Español") {
            $global:languageSelected = "es"
        }

        # Debugging: output the selected language to verify
        Write-Host "Selected language inside button event: $global:languageSelected"

        # Close the language selection form
        $langForm.Close()
    })

# Display the language selection form and block further execution until it's closed
[System.Windows.Forms.Application]::Run($langForm)

# Debugging: check if the language was selected before continuing
Write-Host "Language selected after form closed: $global:languageSelected"

# Ensure the user has selected a language before continuing
if (-not $global:languageSelected) {
    Write-Host "Error: No language selected. Please select a language."
    exit
}

# Now that the language is selected, load the appropriate language for the form
$lang = $languages[$global:languageSelected]

# Define functions to save and load form data
function Save-FormData {
    $formData = @{
        ExePath    = $textBoxExe.Text
        SourcePath = $textBoxSource.Text
        FilePath   = $textBoxFile.Text
        OutputPath = $textBoxOutput.Text
    }
    $json = [System.Text.Json.JsonSerializer]::Serialize($formData)
    Set-Content -Path "formData.json" -Value $json
}

function Load-FormData {
    if (Test-Path "formData.json") {
        $json = Get-Content -Path "formData.json" -Raw
        $formData = [System.Text.Json.JsonSerializer]::Deserialize($json, [Hashtable])
        $textBoxExe.Text = $formData.ExePath
        $textBoxSource.Text = $formData.SourcePath
        $textBoxFile.Text = $formData.FilePath
        $textBoxOutput.Text = $formData.OutputPath
    }
}

# Define function to clear form data
function Clear-FormData {
    $textBoxExe.Text = ""
    $textBoxSource.Text = ""
    $textBoxFile.Text = ""
    $textBoxOutput.Text = ""
}

# Initialize form and controls
$form = New-Object System.Windows.Forms.Form
$form.Text = $lang["Title"]
$form.Size = New-Object System.Drawing.Size(800, 400)  # Increased form width
$form.StartPosition = "CenterScreen"

# Add personalized icon to the form
$iconPath = "https://vhdxgabrielluiz.blob.core.windows.net/vhdx/gabrielluiz-icone.ico"
$iconTempPath = [System.IO.Path]::GetTempFileName() + ".ico"
$webClient = New-Object System.Net.WebClient
$webClient.DownloadFile($iconPath, $iconTempPath)
$icon = New-Object System.Drawing.Icon($iconTempPath)
$form.Icon = $icon

# Add controls to the form
# TextBox for IntuneWinAppUtil.exe path
$labelExe = New-Object System.Windows.Forms.Label
$labelExe.Text = $lang["ExeLabel"]
$labelExe.AutoSize = $true  # Ensure the label adjusts to the text size
$labelExe.Location = New-Object System.Drawing.Point(10, 20)
$form.Controls.Add($labelExe)

$textBoxExe = New-Object System.Windows.Forms.TextBox
$textBoxExe.Size = New-Object System.Drawing.Size(500, 20)  # Increased text box width
$textBoxExe.Location = New-Object System.Drawing.Point(150, 20)
$form.Controls.Add($textBoxExe)

$buttonExe = New-Object System.Windows.Forms.Button
$buttonExe.Text = $lang["Select"]
$buttonExe.Location = New-Object System.Drawing.Point(660, 18)  # Adjusted button location
$buttonExe.Size = New-Object System.Drawing.Size(75, 23)
$buttonExe.Add_Click({
        $exePath = Select-FileDialog
        if ($exePath) {
            $textBoxExe.Text = $exePath
        }
    })
$form.Controls.Add($buttonExe)

# TextBox for source folder path
$labelSource = New-Object System.Windows.Forms.Label
$labelSource.Text = $lang["SourceLabel"]
$labelSource.AutoSize = $true
$labelSource.Location = New-Object System.Drawing.Point(10, 60)
$form.Controls.Add($labelSource)

$textBoxSource = New-Object System.Windows.Forms.TextBox
$textBoxSource.Size = New-Object System.Drawing.Size(500, 20)
$textBoxSource.Location = New-Object System.Drawing.Point(150, 60)
$form.Controls.Add($textBoxSource)

$buttonSource = New-Object System.Windows.Forms.Button
$buttonSource.Text = $lang["Select"]
$buttonSource.Location = New-Object System.Drawing.Point(660, 58)
$buttonSource.Size = New-Object System.Drawing.Size(75, 23)
$buttonSource.Add_Click({
        $folderPath = Select-FolderDialog
        if ($folderPath) {
            $textBoxSource.Text = $folderPath
        }
    })
$form.Controls.Add($buttonSource)

# TextBox for source file path
$labelFile = New-Object System.Windows.Forms.Label
$labelFile.Text = $lang["FileLabel"]
$labelFile.AutoSize = $true
$labelFile.Location = New-Object System.Drawing.Point(10, 100)
$form.Controls.Add($labelFile)

$textBoxFile = New-Object System.Windows.Forms.TextBox
$textBoxFile.Size = New-Object System.Drawing.Size(500, 20)
$textBoxFile.Location = New-Object System.Drawing.Point(150, 100)
$form.Controls.Add($textBoxFile)

$buttonFile = New-Object System.Windows.Forms.Button
$buttonFile.Text = $lang["Select"]
$buttonFile.Location = New-Object System.Drawing.Point(660, 98)
$buttonFile.Size = New-Object System.Drawing.Size(75, 23)
$buttonFile.Add_Click({
        $filePath = Select-FileDialog
        if ($filePath) {
            $textBoxFile.Text = $filePath
        }
    })
$form.Controls.Add($buttonFile)

# TextBox for output folder path
$labelOutput = New-Object System.Windows.Forms.Label
$labelOutput.Text = $lang["OutputLabel"]
$labelOutput.AutoSize = $true
$labelOutput.Location = New-Object System.Drawing.Point(10, 140)
$form.Controls.Add($labelOutput)

$textBoxOutput = New-Object System.Windows.Forms.TextBox
$textBoxOutput.Size = New-Object System.Drawing.Size(500, 20)
$textBoxOutput.Location = New-Object System.Drawing.Point(150, 140)
$form.Controls.Add($textBoxOutput)

$buttonOutput = New-Object System.Windows.Forms.Button
$buttonOutput.Text = $lang["Select"]
$buttonOutput.Location = New-Object System.Drawing.Point(660, 138)
$buttonOutput.Size = New-Object System.Drawing.Size(75, 23)
$buttonOutput.Add_Click({
        $outputPath = Select-FolderDialog
        if ($outputPath) {
            $textBoxOutput.Text = $outputPath
        }
    })
$form.Controls.Add($buttonOutput)

# Button to execute the command
$buttonExecute = New-Object System.Windows.Forms.Button
$buttonExecute.Text = $lang["Execute"]
$buttonExecute.Location = New-Object System.Drawing.Point(200, 180)
$buttonExecute.Size = New-Object System.Drawing.Size(75, 23)
$buttonExecute.Add_Click({
        $exePath = $textBoxExe.Text
        $sourcePath = $textBoxSource.Text
        $filePath = $textBoxFile.Text
        $outputPath = $textBoxOutput.Text
        if ($exePath -and $sourcePath -and $filePath -and $outputPath) {
            $intunewinFile = Join-Path $outputPath ([System.IO.Path]::GetFileNameWithoutExtension($filePath) + ".intunewin")
            if (Test-Path $intunewinFile) {
                $result = [System.Windows.Forms.MessageBox]::Show($lang["WarningMessage"], $lang["WarningMessage1"], [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
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
                [System.Windows.Forms.MessageBox]::Show($lang["WarningMessage"], $lang["WarningMessage2"], [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            else {
                [System.Windows.Forms.MessageBox]::Show($lang["ErrorMessage"], $lang["ErrorMessage1"], [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
            Save-FormData
        }
        else {
            [System.Windows.Forms.MessageBox]::Show($lang["AllFieldsWarning"], $lang["WarningMessage1"], [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
$form.Controls.Add($buttonExecute)

# Button to clear form fields
$buttonClear = New-Object System.Windows.Forms.Button
$buttonClear.Text = $lang["Clear"]
$buttonClear.Location = New-Object System.Drawing.Point(320, 180)
$buttonClear.Size = New-Object System.Drawing.Size(75, 23)
$buttonClear.Add_Click({
        Clear-FormData
    })
$form.Controls.Add($buttonClear)

# Button to download IntuneWinAppUtil.exe
$buttonDownload = New-Object System.Windows.Forms.Button
$buttonDownload.Text = $lang["Download"]
$buttonDownload.Location = New-Object System.Drawing.Point(200, 210)
$buttonDownload.Size = New-Object System.Drawing.Size(200, 23)
$buttonDownload.Add_Click({
        Start-Process "https://raw.githubusercontent.com/microsoft/Microsoft-Win32-Content-Prep-Tool/master/IntuneWinAppUtil.exe"
    })
$form.Controls.Add($buttonDownload)

# Add image below the download button
try {
    $imageURL = "https://vhdxgabrielluiz.blob.core.windows.net/vhdx/perfil.png"
    $iconImage = [System.Drawing.Image]::FromStream((New-Object System.Net.WebClient).OpenRead($imageURL))
    $pictureBox = New-Object System.Windows.Forms.PictureBox
    $pictureBox.Image = $iconImage
    $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
    $pictureBox.Size = New-Object System.Drawing.Size(64, 64)
    $pictureBox.Location = New-Object System.Drawing.Point(257, 240)
    $pictureBox.Cursor = [System.Windows.Forms.Cursors]::Hand
    $pictureBox.Add_Click({
            Start-Process "https://gabrielluiz.com"
        })
    $form.Controls.Add($pictureBox)
}
catch {
    Write-Host "$($lang["WarningMessage3"]) $_"
}

# Load form data on start-up
$form.Add_Shown({
        Load-FormData
        $form.Activate()
    })

# Run the form
[System.Windows.Forms.Application]::Run($form)