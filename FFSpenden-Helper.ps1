#############################################################################
## Filename: FFSpenden-Helper.ps1
## Author:   Michael Daniel
## E-Mail:   michael.daniel.h@gmail.com
## Purpose:  Creates a SZR-Upload File based on an Excel-List of donators.
##           The response SZR Ziparchive can be imported for creating the
##           Finanzonline XML File.
## Thanks:   Uses dfinke's awesome ImportExcel module
##
#############################################################################



#############################################################################
## Functions
#############################################################################

# Download of the ImportExcel files if PS4 is used.
# The provided Install Script is buggy and not functioning, so I adapted it by myself.
Function Get-ModuleImportExcelFiles() {
    [CmdLetBinding()]
    Param (
        [ValidateNotNullOrEmpty()]
        [String]$ModuleName = 'ImportExcel',
        [String]$InstallDirectory,
        [ValidateNotNullOrEmpty()]
        [String]$GitPath = 'https://raw.github.com/dfinke/ImportExcel/master'
    )

    Begin {
        Try {
            Write-Verbose "$ModuleName Modulinstallation wurde gestartet"

            $Files = @(
                'EPPlus.dll',
                'ImportExcel.psd1',
                'ImportExcel.psm1',
                'AddConditionalFormatting.ps1',
                'Charting.ps1',
                'ColorCompletion.ps1',
                'ConvertFromExcelData.ps1',
                'ConvertFromExcelToSQLInsert.ps1',
                'ConvertExcelToImageFile.ps1',
                'ConvertToExcelXlsx.ps1',
                'Copy-ExcelWorkSheet.ps1',
                'Export-charts.ps1',
                'Export-Excel.ps1',
                'Export-ExcelSheet.ps1',
                'formatting.ps1',
                'Get-ExcelColumnName.ps1',
                'Get-ExcelSheetInfo.ps1',
                'Get-ExcelWorkbookInfo.ps1',
                'Get-HtmlTable.ps1',
                'Get-Range.ps1',
                'Get-XYRange.ps1',
                'Import-Html.ps1',
                'InferData.ps1',
                'Invoke-Sum.ps1',
                'New-ConditionalFormattingIconSet.ps1',
                'New-ConditionalText.ps1',
                'New-ExcelChart.ps1',
                'New-PSItem.ps1',
                'Open-ExcelPackage.ps1',
                'Pivot.ps1',
                'plot.ps1',
                'Send-SqlDataToExcel.ps1',
                'Set-CellStyle.ps1',
                'Set-Column.ps1',
                'Set-Row.ps1',
                'SetFormat.ps1',
                'TrackingUtils.ps1',
                'Update-FirstObjectProperties.ps1'
            )
        }
        Catch {
            throw "Fehler beim Installieren des Moduls in das Verzeichnis '$InstallDirectory': $_"
        }
    }

    Process {
        Try {
            if (-not $InstallDirectory) {
                Write-Verbose "$ModuleName Kein Installationsverzeichnis angegeben"

                $PersonalModules = Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath WindowsPowerShell\Modules

                if (($env:PSModulePath -split ';') -notcontains $PersonalModules) {
                    Write-Warning "$ModuleName Persoenlicher Modulpfad '$PersonalModules' nicht gefunden in '`$env:PSModulePath'"
                }

                if (-not (Test-Path $PersonalModules)) {
                    Write-Error "$ModuleName Pfad '$PersonalModules' existiert nicht"
                }

                $InstallDirectory = Join-Path -Path $PersonalModules -ChildPath $ModuleName
                Write-Verbose "$ModuleName Installationsverzeichnis ist '$InstallDirectory'"
            }

            if (-not (Test-Path $InstallDirectory)) {
                New-Item -Path $InstallDirectory -ItemType Directory -EA Stop | Out-Null
                Write-Verbose "$ModuleName Modul-Ordner wurde erstellt: '$InstallDirectory'"
            }

            $WebClient = New-Object System.Net.WebClient
        
            $Files | ForEach-Object {
                $WebClient.DownloadFile("$GitPath/$_","$installDirectory\$_")
                Write-Verbose "$ModuleName Moduldatei erfolgreich installiert: '$_'"
            }

            Write-Verbose "$ModuleName Modulinstallation erfolgreich abgeschlossen"
        }
        Catch {
            throw "Fehler beim Installieren des Moduls in das Verzeichnis '$InstallDirectory': $_"
        }
    }
}

# Get a filename by picking a file from a dialog-box
# filtering is available
Function Get-FileName($initialDirectory, $filter = "Excel (*.xlsx)| *.xlsx")
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = $filter
    $OpenFileDialog.ShowDialog() | Out-Null
    return $OpenFileDialog.FileName
}

# shows a info message box (for informational messages to the user
Function Show-MessageboxInfo($message, $title, $buttons = [System.Windows.MessageBoxButton]::OK)
{
    [System.Windows.MessageBox]::Show($message,$title,$buttons,'Information')
}

# shows a error message box
Function Show-MessageboxError($message, $title, $buttons = [System.Windows.MessageBoxButton]::OK)
{
    [System.Windows.MessageBox]::Show($message,$title,$buttons, 'Error')
}

# writes a XML file for uploading in Finanzonline
Function Write-FoXML($listDonators, $steuernr, $messageRefId, $zeitraum, $xmlFilePath)
{
    $xmlWriter = New-Object System.Xml.XmlTextWriter($xmlFilePath,$null)
    
    #formatting
    $xmlWriter.Formatting = "Indented"
    $xmlWriter.Indentation = 1
    $xmlWriter.IndentChar = "`t"

    # header
    $xmlWriter.WriteStartDocument()

    # create root element
    $xmlWriter.WriteStartElement("SonderausgabenUebermittlung")
    $xmlWriter.WriteAttributeString("xmlns","https://finanzonline.bmf.gv.at/fon/ws/uebermittlungSonderausgaben")

    # create some metadata nodes
    $xmlWriter.WriteStartElement("Info_Daten")
    $xmlWriter.WriteElementString("Fastnr_Fon_Tn",$steuernr)
    $xmlWriter.WriteElementString("Fastnr_Org",$steuernr)
    $xmlWriter.WriteEndElement()

    $xmlWriter.WriteStartElement("MessageSpec")
    $xmlWriter.WriteElementString("MessageRefId", $messageRefId)
    $xmlWriter.WriteElementString("Timestamp", (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss"))
    $xmlWriter.WriteElementString("Uebermittlungsart","FF")
    $xmlWriter.WriteElementString("Zeitraum",$zeitraum)
    $xmlWriter.WriteEndElement()

    $refnr=0
    foreach($donator in $listDonators) {
        if($donator.Vorname.Length -gt 0 -and $donator.Nachname.Length -gt 0 -and $donator.Geburtsdatum.Length -gt 0 -and $donator.vbpkSA.Length -gt 0) {
            $xmlWriter.WriteStartElement("Sonderausgaben")
            $xmlWriter.WriteAttributeString("Uebermittlungs_Typ","E")
            $xmlWriter.WriteElementString("RefNr", ((Get-Date).ToString("yyyyMMdd") + ($refnr.ToString())) )
            $xmlWriter.WriteElementString("Betrag",$donator.'Summe Spenden')
            $xmlWriter.WriteElementString("vbPK", $donator.vbpkSA)
            $xmlWriter.WriteEndElement()
            $refnr++
        }
    }

    # close root node
    $xmlWriter.WriteEndElement()

    # finalize the document
    $xmlWriter.WriteEndDocument()
    $xmlWriter.Flush()
    $xmlWriter.Close()
     
}

# extract valid donators from an excel worksheet
# a donator is valid if the firstname, the surname and the birthday is present
Function Get-ValidDonatorsFromExcelFile($filepath, $sheet, $rowstart)
{
    $ImportedData = Import-Excel -Path $filepath -WorksheetName $sheet -StartRow $rowstart

    $listDonators = @()

    foreach ($donator in $ImportedData){
        # check if firstname, surname, birthdate is present
        if(($donator.Vorname).Length -lt 1){
            continue
        }
        if(($donator.Nachname).Length -lt 1){
            continue
        }
    
        try{
            if(Get-Date($donator.Geburtsdatum)) {}
        }
        catch{
            continue
        }
   
        $listDonators += $donator
    }

    return $listDonators
}

Function Get-SZRVKZ($filepath, $sheet)
{
    $szrheader = Import-Excel -Path $filepath -WorksheetName $sheet -HeaderName "Content"
    foreach ($row in $szrheader) {
        if(($row).Content -ne $null -and ($row).Content.StartsWith("VKZ")) {
            $vkz=($row).Content.Split("=")
            if( $vkz.Count -eq 2) {
                return $vkz[1]
            }
        }
    }
    return ""
}

Function Get-SZRUploadFile($filepath, $sheet, $listDonators)
{
    # Read the header Information
    $szrheader = Import-Excel -Path $filepath -WorksheetName $sheet -HeaderName "Content"
    $szrfilecontent = ""
    $szrfileNrOfCols = 0
    foreach ($row in $szrheader) {
        $szrfilecontent += $row.Content
        $szrfilecontent += "`n"  
        if($row.Content -ne $null -and ($row).Content.StartsWith("LAUFNR")) {
            $szrfileNrOfCols = ($row.Content.Split(";")).Count    
        }
    }

    $laufnr = 0
    foreach ($donator in $listDonators) {
        $laufnr++
        $szrfilecontent += $laufnr.ToString() + ";"
        $szrfilecontent += $donator.Nachname + ";"
        $szrfilecontent += $donator.Vorname + ";"
        $szrfilecontent += (Get-Date($donator.Geburtsdatum)).ToString("dd.MM.yyyy") + ";"
        for($i=4; $i -lt $szrfileNrOfCols; $i++) {
            $szrfilecontent += ";"
        }
        $szrfilecontent += "`n"
    }
    return $szrfilecontent
}

Function Get-FOMetaData($filepath, $sheet, $header) {
    $data = Import-Excel -Path $filepath -WorksheetName $sheet
    return $data.$header
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
Function Unzip
{
    param([string]$zipfile, [string]$outpath)
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile,$outpath)
}

Function Set-vbPKSA($szrAnswerfile, $listDonators) {
    $szrAnswerFileUnzipped = Join-Path ((Get-Item $szrAnswerfile).DirectoryName) -ChildPath (Get-Item $szrAnswerfile).BaseName   
    Unzip $szrAnswerfile $szrAnswerFileUnzipped

    $bpkfile = Get-ChildItem $szrAnswerFileUnzipped | Where-Object {$_.BaseName.Contains("VERSCHL_BPK") }
    $headerPassed = $false
    $NrNachname = 0
    $NrVorname = 0
    $NrGebDatum = 0
    $NrVBPK = 0
    $bpkcontent = Get-Content $bpkfile.FullName
    
    foreach ($row in $bpkcontent) {
        if($row.StartsWith("LAUFNR")){
            $headerPassed = $true
            $NrNachname = $row.Split(";").IndexOf("NACHNAME")
            $NrVorname = $row.Split(";").IndexOf("VORNAME")
            $NrGebDatum = $row.Split(";").IndexOf("GEBDATUM")
            $splitted = $row.Split(";")
            $NrVBPK = $splitted.IndexOf( ($splitted | Where-Object { $_.StartsWith("VBPK") } ) )
            continue       
        }

        if($headerPassed) {
            $splitted = $row.Split(";")
            $nachname = $splitted[$NrNachname]
            $vorname = $splitted[$NrVorname]
            $gebdatum = $splitted[$NrGebDatum]
            $vbpk = $splitted[$NrVBPK]

            foreach ($donator in $listDonators) {
                if(($donator.Vorname -eq $vorname) -and ($donator.Nachname -eq $nachname) -and ($donator.Geburtsdatum -eq $gebdatum)) {
                    $donator | Add-Member -MemberType NoteProperty -Name "vbpkSA" -Value $vbpk
                    break
                }
            }

        }
    }
    
    Remove-Item $szrAnswerFileUnzipped -Force -Recurse
}

#############################################################################
## actual script start
#############################################################################

# Install ImportExcel Module - depending on installed PS Version
if($PSVersionTable.PSVersion.Major -eq 5) {
    if (-NOT (Get-Module -ListAvailable -Name ImportExcel)){
        Install-Module ImportExcel -scope CurrentUser
    }
} elseif($PSVersionTable.PSVersion.Major -eq 4) {
    if (-NOT (Get-Module -ListAvailable -Name ImportExcel)) {
        Get-ModuleImportExcelFiles -Verbose
    }
} else {
    # not supported / tested version of Powershell
}

# Double check if the module is (properly installed)
if (-NOT (Get-Module -ListAvailable -Name ImportExcel)){
        Show-MessageboxError "Das benötigte Modul `"ImportExcel`" konnte nicht installiert werden. Überprüfen Sie die Internetverbindung und beachten Sie die roten Fehlermeldungen." "Fehler bei Modulinstallation" | Out-Null
        exit
}

# style information (can be done with Visual Studio - just paste the xaml content here in $inputXML)
$inputXML = @"
<Window x:Class="FF_Spenden_Helper.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:FF_Spenden_Helper"
        mc:Ignorable="d"
        Title="FF Spenden Helper" Height="615.581" Width="556.336">
    <ScrollViewer>
        <Grid>
            <GroupBox Header="Schritt 1: Stammzahlenregister Upload-Datei erstellen" HorizontalAlignment="Left" Height="175" Margin="10,10,0,0" VerticalAlignment="Top" Width="497">
                <Grid>
                    <Button x:Name="btn_chooseExcelFile" Content="Datei auswählen" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="123"/>
                    <TextBox x:Name="tb_loadedFile" HorizontalAlignment="Left" Height="23" Margin="188,10,0,0" TextWrapping="Wrap" Text="Keine Datei ausgewählt." VerticalAlignment="Top" Width="287"/>
                    <Label Content="Arbeitsmappe (Spender):" HorizontalAlignment="Left" Margin="10,35,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="tb_WorksheetDonators" HorizontalAlignment="Left" Height="23" Margin="188,38,0,0" TextWrapping="Wrap" Text="2017" VerticalAlignment="Top" Width="287"/>
                    <Label Content="Anfangszeile:" HorizontalAlignment="Left" Margin="10,66,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="tb_BeginRow" HorizontalAlignment="Left" Height="23" Margin="188,70,0,0" TextWrapping="Wrap" Text="4" VerticalAlignment="Top" Width="287"/>
                    <Label Content="Arbeitsmappe (SZR-Header):" HorizontalAlignment="Left" Margin="10,98,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="tb_WorksheetSZR" HorizontalAlignment="Left" Height="23" Margin="188,102,0,0" TextWrapping="Wrap" Text="SZR-Header" VerticalAlignment="Top" Width="287"/>
                    <Button x:Name="btn_GenerateSZRFile" Content="Stammzahlenregister Upload-Datei generieren" HorizontalAlignment="Left" Margin="15,129,0,0" VerticalAlignment="Top" Width="460"/>

                </Grid>
            </GroupBox>
            <GroupBox Header="Schritt 2: Stammzahlenregister Upload-Datei hochladen" HorizontalAlignment="Left" Height="73" Margin="10,203,0,0" VerticalAlignment="Top" Width="497">
                <Grid>
                    <TextBlock HorizontalAlignment="Left" Margin="10,10,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="465" Height="31"><Run Text="Die generierte Datei in Finanzonline (Extern -&gt; Stammzahlenregister) hochladen."/><LineBreak/><Run Text="Die Antwort-Datei (ZIP-Archiv) herunterladen und im nächsten Schritt angeben."/></TextBlock>
                </Grid>
            </GroupBox>
            <GroupBox Header="Schritt 3: Finanzonline Upload-Datei erstellen" HorizontalAlignment="Left" Height="112" Margin="10,303,0,0" VerticalAlignment="Top" Width="497">
                <Grid>
                    <Button x:Name="btn_chooseSZRFile" Content="Datei auswählen" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="123"/>
                    <TextBox x:Name="tb_loadedSZRFile" HorizontalAlignment="Left" Height="34" Margin="188,10,0,0" TextWrapping="Wrap" Text="Keine Datei ausgewählt." VerticalAlignment="Top" Width="287"/>
                    <Button x:Name="btn_GenerateFOFile" Content="Finanzonline Upload-Datei generieren" HorizontalAlignment="Left" Margin="10,49,0,0" VerticalAlignment="Top" Width="465"/>
                </Grid>
            </GroupBox>
            <GroupBox Header="Schritt 4: XML-Datei bei Finanzonline hochladen" HorizontalAlignment="Left" Height="119" Margin="10,447,0,0" VerticalAlignment="Top" Width="497">
                <Grid >
                    <TextBlock HorizontalAlignment="Left" Margin="10,10,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="465" Height="77"><Run Text="Auf Finanzonline einsteigen (https://finanzonlinge.bmf.gv.at). Danach auf "/><Run Text=" &quot;Eingaben&quot; und &quot;Übermittlung&quot; gehen. Den Punkt &quot;Sonderausgaben&quot; ankreuzen.  Mit der Schaltfläche &quot;Datei auswählen&quot;, die soeben erstellte Datei auswählen"/><Run Text=" und den Vorgang mit Klick auf &quot;Datei senden&quot; abschließen."/></TextBlock>
                </Grid>
            </GroupBox>
        </Grid>
    </ScrollViewer>
</Window>
"@

$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML

#Read XAML 
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
 
 try {
    $Form=[Windows.Markup.XamlReader]::Load( $reader )
}
catch {
    Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
}
 
#===========================================================================
# Store Form Objects In PowerShell
#===========================================================================
 
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}
 
Function Get-FormVariables{

    if ($global:ReadmeDisplay -ne $true){
        Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true
    }

    write-host "Found the following interactable elements from our form" -ForegroundColor Cyan

    get-variable WPF*

}
 
#Get-FormVariables

#===========================================================================
# Global Variables
#===========================================================================
[string]$SzrVkz = ""
$ListDonators = @()
[string]$Steuernummer = ""
[string]$MessageRefId = ""
 
#===========================================================================
# Actually make the objects work
#===========================================================================

# Button for choosing Excel file with donators
$WPFbtn_chooseExcelFile.Add_Click({
    $WPFtb_loadedFile.Text = Get-FileName ".\"
})

# Button for starting the SZR File generation
$WPFbtn_GenerateSZRFile.Add_Click({
    if( -NOT (Test-Path $WPFtb_loadedFile.Text)) {
        Show-MessageboxError "Es wurde keine gültige Excel-Datei ausgewählt. Bitte nochmals überprüfen." "Keine Datei gewählt."
        return
    }
    try {
    $ListDonators = Get-ValidDonatorsFromExcelFile $WPFtb_loadedFile.Text $WPFtb_WorksheetDonators.Text $WPFtb_BeginRow.Text
    $szrUploadFile = Get-SZRUploadFile $WPFtb_loadedFile.Text $WPFtb_WorksheetSZR.Text $listDonators
    $SzrVkz = Get-SZRVKZ $WPFtb_loadedFile.Text $WPFtb_WorksheetSZR.Text
    if($SzrVkz.Length -gt 0) {
        $szrFilename = "BPK_" + $SzrVkz + "_" + $(Get-Date).ToString("yyyyMMddHHmm") + ".csv"
        $szrUploadFile | Out-File -Encoding utf8 (".\" + $szrFilename)
        Show-MessageboxInfo ("Die Datei " + $szrFilename + " wurde erfolgreich erstellt. Es wurden " + $ListDonators.Count + " Spender ermittelt." ) "SZR-Datei Erstellung erfolgreich." 
    }
    }
    catch {
        Show-MessageboxError ("Ein Fehler ist aufgetreten: " + $_.Exception.Message + "`nBitte die Eingabefelder kontrollieren.") "Fehler bei Spender-Ermittlung."
        return
    }
})

# Button for choosing an SZR response archive
$WPFbtn_chooseSZRFile.Add_Click({
    $WPFtb_loadedSZRFile.Text = Get-FileName ".\" "ZIP-Archiv (*.zip)| *.zip"
})

# Button for generating an Finanzonline XML
$WPFbtn_GenerateFOFile.Add_Click({
    if( -NOT (Test-Path $WPFtb_loadedFile.Text)) {
        Show-MessageboxError "Es wurde keine gültige Excel-Datei ausgewählt. Bitte nochmals überprüfen.", "Keine Excel-Datei gewählt."
        return
    }
    if( -NOT (Test-Path $WPFtb_loadedSZRFile.Text)) {
        Show-MessageboxError "Es wurde keine SZR-Antwort-Datei ausgewählt. Bitte nochmals überprüfen.", "Keine SZR-Datei gewählt."
        return
    }
    
    $Steuernummer = Get-FOMetaData $WPFtb_loadedFile.Text "Finanzonline" "Steuernummer"
    $MessageRefId = Get-FOMetaData $WPFtb_loadedFile.Text "Finanzonline" "MessageRefId"
    $ListDonators = Get-ValidDonatorsFromExcelFile $WPFtb_loadedFile.Text $WPFtb_WorksheetDonators.Text $WPFtb_BeginRow.Text

    Set-vbPKSA $WPFtb_loadedSZRFile.Text $ListDonators

    $outputXMLFilePath =  Join-Path ((Get-Item $WPFtb_loadedSZRFile.Text).Directory) -ChildPath ((Get-Item $WPFtb_loadedSZRFile.Text).BaseName + ".xml")

    Write-FoXML $ListDonators $Steuernummer $MessageRefId $WPFtb_WorksheetDonators.Text $outputXMLFilePath

    Show-MessageboxInfo ("Die Datei " + $outputXMLFilePath + " wurde erfolgreich erstellt. Die Datei kann bei Finanzonline hochgeladen werden.") "Finanzonline-XML Erstellung erfolgreich."
})
 
 
#===========================================================================
# Show the form
#===========================================================================
$Form.ShowDialog() | out-null

