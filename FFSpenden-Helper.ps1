#$executionPolicy = Get-ExecutionPolicy
#Set-ExecutionPolicy -ExecutionPolicy Bypass

if (-NOT (Get-Module -ListAvailable -Name ImportExcel)){
    Install-Module ImportExcel -scope CurrentUser
}


Function Get-FileName($initialDirectory, $filter = "Excel (*.xlsx)| *.xlsx")
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = $filter
    $OpenFileDialog.ShowDialog() | Out-Null
    return $OpenFileDialog.FileName
}

Function Show-MessageboxInfo($message, $title, $buttons = [System.Windows.Forms.MessageBoxButtons]::OK)
{
    [System.Windows.MessageBox]::Show($message,$title,$buttons,[System.Windows.Forms.MessageBoxIcon]::Information)
}

Function Show-MessageboxError($message, $title, $buttons = [System.Windows.Forms.MessageBoxButtons]::OK)
{
    [System.Windows.MessageBox]::Show($message,$title,$buttons, [System.Windows.Forms.MessageBoxIcon]::Error)
}

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
$WPFbtn_chooseExcelFile.Add_Click({
    $WPFtb_loadedFile.Text = Get-FileName ".\"
})

$WPFbtn_GenerateSZRFile.Add_Click({
    if( -NOT (Test-Path $WPFtb_loadedFile.Text)) {
        Show-MessageboxError "Es wurde keine gültige Excel-Datei ausgewählt. Bitte nochmals überprüfen.", "Keine Datei gewählt."
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

$WPFbtn_chooseSZRFile.Add_Click({
    $WPFtb_loadedSZRFile.Text = Get-FileName ".\" "ZIP-Archiv (*.zip)| *.zip"
})

$WPFbtn_GenerateFOFile.Add_Click({
    if( -NOT (Test-Path $WPFtb_loadedFile.Text)) {
        Show-MessageboxError "Es wurde keine gültige Excel-Datei ausgewählt. Bitte nochmals überprüfen.", "Keine Datei gewählt."
        return
    }
    if( -NOT (Test-Path $WPFtb_loadedSZRFile.Text)) {
        Show-MessageboxError "Es wurde keine SZR-Antwort-Datei ausgewählt. Bitte nochmals überprüfen.", "Keine Datei gewählt."
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

