Param
(
    [parameter(Mandatory=$true)]
    [String]
    $CsvFileName
)



$SiteUrl = "https://m365x337444.sharepoint.com/sites/PorscheInformatik"



#-----------------------------------------------------------------------------------------
# functions 
#-----------------------------------------------------------------------------------------
function Write-Log([string]$logtext, [int]$level=0)
{
	$logdate = get-date -format "yyyy-MM-dd HH:mm:ss"
	if($level -eq 0)
	{
		$logtext = "[INFO] " + $logtext
		$text = "["+$logdate+"] - " + $logtext
		Write-Host $text
	}
	if($level -eq 1)
	{
		$logtext = "[WARNING] " + $logtext
		$text = "["+$logdate+"] - " + $logtext
		Write-Host $text -ForegroundColor Yellow
	}
	if($level -eq 2)
	{
		$logtext = "[ERROR] " + $logtext
		$text = "["+$logdate+"] - " + $logtext
		Write-Host $text -ForegroundColor Red
	}
	$text >> $logfile
}

#------------------------------------------
# creating log-file
#------------------------------------------
$path = ".\logs"
$date = get-date -format "yyyy-MM-dd-HH-mm"
$file = ("NewsLog_" + $date + ".log")
$logfile = $path + "\" + $file


Connect-PnPOnline -Url $SiteUrl 

Write-Log "$logTitle : Reading from file '$CsvFileName' ..."
$Items = Import-CSV $CsvFileName -Delimiter ';' 
Write-Log "$Items.Length news found"
foreach ($Item in $Items){ 
    $xmlFileName = $Item.Category + ".xml"
    $NewsFileName = $Item.Ident + ".aspx"
    $tmpXmlFileName = $Item.Ident + ".xml"
    $NewsTitle = $Item.Title
    $newsContent = $Item.Content    
    Write-Log "Creating News '$NewsTitle' from Template '$xmlFileName'..."
    
    # Create a XML document
    [xml]$xmlDoc = New-Object system.Xml.XmlDocument
  
    # Read the template file
    $xmlDoc = [xml](get-content (Resolve-Path  $xmlFileName)) 
    $xmlDoc.Provisioning.Templates.ProvisioningTemplate.ClientSidePages.ClientSidePage.PageName = $newsFileName
    $xmlDoc.Provisioning.Templates.ProvisioningTemplate.ClientSidePages.ClientSidePage.Title = $newsTitle
    $xmlDoc.Provisioning.Templates.ProvisioningTemplate.ClientSidePages.ClientSidePage.Sections.Section.Controls.CanvasControl.CanvasControlProperties.CanvasControlProperty.Value = $newsContent

    Write-Log "Writing File '$tmpXmlFileName' ..."
    $saveResult = $xmlDoc.Save("$pwd\$tmpXmlFileName")
 
    Write-Log "Uploading '$NewsTitle' ..."
    $result = Apply-PnPProvisioningTemplate $tmpXmlFileName
    Write-Log "Uploading '$NewsTitle' done"
}


