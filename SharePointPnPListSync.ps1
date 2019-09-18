Param
(
    [parameter(Mandatory=$true)]
    [String]
    $paramFileName
)

#-----------------------------------------------------------------------------------------
# functions 
#-----------------------------------------------------------------------------------------
function Get-Data
{
    [OutputType([Array])]
    Param(
        [Parameter(Mandatory)]
        [Object]$spec,
        [Parameter(Mandatory)]
        [Object]$fields,
        [Parameter(Mandatory)]
        [string]$logTitle

    )


    $type = $spec.Type.ToLower()
    Switch ($type)  {
        "SharePointOnline" {
            $items = Get-SharePoint-Online-List-Data -SiteUrl $spec.SiteUrl -ListName $spec.ListName -FieldNames $fields -LogTitle $logTitle
        }
        "csv" {
            $items = Get-Csv-Data -FilePath $spec.FilePath -LogTitle $logTitle
        }
        "json" {
            $items = Get-Content -Path $spec.FilePath -Raw | ConvertFrom-Json 
        }
        "sql" {
            $items = Get-Sql-Data -ConnectionString $spec.ConnectionString -Query $spec.Query
        }
        
    }
    
    $log.Info("$logTitle : $($items.length) Items found")
    $items
}


function Get-Sql-Data 
{
    Param(
        [Parameter(Mandatory)]
        [Object]$ConnectionString,
        [Parameter(Mandatory)]
        [string]$Query
    )
    $connection = New-Object System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = $ConnectionString
    $connection.Open()
    $command = $connection.CreateCommand()
    $command.CommandText = $Query
    $DataAdapter = new-object System.Data.SqlClient.SqlDataAdapter $command
    $dataset = new-object System.Data.Dataset
    $dummyCnt = $DataAdapter.Fill($dataset)
    $dt = $dataset.Tables[0]

    $items = New-Object System.Collections.ArrayList

    foreach ($row in $dt.Rows) {
        $rht = @{}
        for ($i = 0; $i -lt $dt.Columns.Count; $i++) {
            $colName = $dt.Columns[$i].ColumnName
            $rht.$colName = $row[$colName]
        }
        $items += [PSCustomObject]$rht
    }
    return $items
}


function Get-Csv-Data
{
    Param(
        [Parameter(Mandatory)]
        [Object]$FilePath,
        [Parameter(Mandatory)]
        [string]$LogTitle
    )

    #----------------------------------------
    # reading from csv file
    #----------------------------------------
    if ($FilePath) {
        $log.Info("$logTitle : Reading from file '$FilePath' ...")
        $Items = Import-CSV $FilePath -Delimiter ';' 
    } else {
        $log.Error("$logTitle : Missing parameter FilePath!")
    }
    return $Items
}
function Get-SharePoint-Online-List-Data
{
    Param(
        [Parameter(Mandatory)]
        [Object]$SiteUrl,
        [Parameter(Mandatory)]
        [string]$ListName,
        [Parameter(Mandatory)]
        [Object]$FieldNames,
        [Parameter(Mandatory)]
        [string]$LogTitle
    )

    #----------------------------------------
    # reading from SharePoint Online
    #----------------------------------------
    if ($SiteUrl) {
        $log.Info("$LogTitle : connecting to  $SiteUrl ...")
        try {
            Connect-PnPOnline -Url $SiteUrl 
            if ($ListName) {
                $log.Info("$logTitle : reading source items from list '$ListName' ...")
                $items = (Get-PnPListItem -List $ListName -Fields $FieldNames).FieldValues 
            } else {
                $log.Error("$logTitle : Missing ListName in configuration!")
            }   
        } catch {
            $e = $_.Exception
            $msg = $e.Message
            $log.Error("$logTitle : Error reading from SharePoint Online $msg")
        }
    } else {
        $log.Error("$logTitle : Missing SiteUrl in configuration!")
    }
    return $items
}
 

function Update-Data
{
    Param(
        [Parameter(Mandatory)]
        [Object]$Spec,
        [Parameter(Mandatory)]
        [Object]$SourceItem,
        [Parameter(Mandatory)]
        [Object]$TargetItem,
        [Parameter(Mandatory)]
        [Object]$Changes,
        [Parameter(Mandatory)]
        [string]$logTitle

    )
    Switch ($Spec.Type)  {
        "SharePointOnline" {
            try {
                $result = Set-PnPListItem -List $Spec.ListName -Identity $targetItem.ID -Values $Changes -ErrorAction Stop
                $log.Info("$logTitle Successfully updated SharePoint Online Item $($result.ID) in List $($Spec.ListName)")
            } catch {
                $e = $_.Exception
                $msg = $e.Message
                $log.Error("$logTitle Error updating SharePoint Online $msg")
            }
        }
        "csv" {
            $log.Fatal("$logTitle CSV Target not implemented !!!!")
        }
        "json" {
            $log.Fatal("$logTitle JSON Target not implemented !!!!")
        }
    }
}


function Insert-Data
{
    Param(
        [Parameter(Mandatory)]
        [Object]$Spec,
        [Parameter(Mandatory)]
        [Object]$SourceItem,
        [Parameter(Mandatory)]
        [Object]$Values,
        [Parameter(Mandatory)]
        [string]$logTitle

    )
    Switch ($Spec.Type)  {
        "SharePointOnline" {
            try {
                $result = Add-PnPListItem -List $Spec.ListName -Values $Values
                $log.Info("$logTitle Successfully added SharePoint Online Item $($result.ID) in List $($Spec.ListName)")
            } catch {
                $e = $_.Exception
                $msg = $e.Message
                $log.Error("$logTitle Error updating SharePoint Online $msg")
            }
        }
        "csv" {
            $log.Fatal("$logTitle CSV Target not implemented !!!!")
        }
        "json" {
            $log.Fatal("$logTitle JSON Target not implemented !!!!")
        }
    }
}

function Remove-Data
{
    Param(
        [Parameter(Mandatory)]
        [Object]$Spec,
        [Parameter(Mandatory)]
        [Object]$Item,
        [Parameter(Mandatory)]
        [string]$logTitle

    )
    Switch ($Spec.Type)  {
        "SharePointOnline" {
            try {
                $result = Remove-PnPListItem -List $Spec.ListName -Identity $Item.ID -Force
                $log.Info("$logTitle : Successfully removed SharePoint Online Item $($Item.ID) in List $($Spec.ListName)")
            } catch {
                $e = $_.Exception
                $msg = $e.Message
                $log.Error("$logTitle : Error removing item in SharePoint Online $msg")
            }
        }
        "csv" {
            $log.Fatal("$logTitle CSV Target not implemented !!!!")
        }
        "json" {
            $log.Fatal("$logTitle JSON Target not implemented !!!!")
        }
    }
}
function Get-Field-Value-As-String($value, $format, $attr, $values) {

    $result = $value
    switch($format) {
        "SPLink" {
            $url = $value.Url
            $descr = $value.Description
            #$result = $url
            $result = "$url, $descr"
        }
        "Url" {
            if ($attr.UrlDescription) {
                $descr = $attr.UrlDescription
                if ($descr.Substring(0,1) -eq "#") {
                    # replace the attribute with the value of a field
                    $varName = $descr.Substring(1)
                    $descr = $values.$varName
                }
                $result = "$value, $descr"           
            } else {
                $result = "$value, $value"
            }

        }
        "Date" {
            $val = $value -as [datetime]
            $result = $val.ToString('yyyy-MM-ddTHH:mm:ss')
        }
        "SPDate" {
            $val = $value -as [datetime]
            if($val) {
                $result = $val.ToString('yyyy-MM-ddTHH:mm:ss')
            } else {
                $result = $null
            }
        }
    }
    return $result
}

function Convert-UTCtoLocal
{
    param(
        [parameter(Mandatory=$true)]
        [String] $UTCTime
    )

    $strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName
    $TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)
    $LocalTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($UTCTime, $TZ)
    return $LocalTime
}
#--------------------------------------------------------------
# Initialize logger (log4net)
#--------------------------------------------------------------
[void][Reflection.Assembly]::LoadFile(([System.IO.Directory]::GetParent($MyInvocation.MyCommand.Path)).FullName+"\log4net.dll");
[log4net.LogManager]::ResetConfiguration();

#File Appender initialization
$FileApndr = new-object log4net.Appender.FileAppender(([log4net.Layout.ILayout](new-object log4net.Layout.PatternLayout('[%date{yyyy-MM-dd HH:mm:ss.fff} (%utcdate{yyyy-MM-dd HH:mm:ss.fff})] [%level] [%message]%n')),(([System.IO.Directory]::GetParent($MyInvocation.MyCommand.Path)).FullName+'\logs\SharePointSync.log'),$True));
$FileApndr.Threshold = [log4net.Core.Level]::All;
[log4net.Config.BasicConfigurator]::Configure($FileApndr);

#Colored Console Appender initialization
$ColConsApndr = new-object log4net.Appender.ColoredConsoleAppender(([log4net.Layout.ILayout](new-object log4net.Layout.PatternLayout('[%date{yyyy-MM-dd HH:mm:ss.fff}] %message%n'))));
$ColConsApndrDebugCollorScheme=new-object log4net.Appender.ColoredConsoleAppender+LevelColors; $ColConsApndrDebugCollorScheme.Level=[log4net.Core.Level]::Debug; $ColConsApndrDebugCollorScheme.ForeColor=[log4net.Appender.ColoredConsoleAppender+Colors]::Green;
$ColConsApndr.AddMapping($ColConsApndrDebugCollorScheme);
$ColConsApndrInfoCollorScheme=new-object log4net.Appender.ColoredConsoleAppender+LevelColors; $ColConsApndrInfoCollorScheme.level=[log4net.Core.Level]::Info; $ColConsApndrInfoCollorScheme.ForeColor=[log4net.Appender.ColoredConsoleAppender+Colors]::White;
$ColConsApndr.AddMapping($ColConsApndrInfoCollorScheme);
$ColConsApndrWarnCollorScheme=new-object log4net.Appender.ColoredConsoleAppender+LevelColors; $ColConsApndrWarnCollorScheme.level=[log4net.Core.Level]::Warn; $ColConsApndrWarnCollorScheme.ForeColor=[log4net.Appender.ColoredConsoleAppender+Colors]::Yellow;
$ColConsApndr.AddMapping($ColConsApndrWarnCollorScheme);
$ColConsApndrErrorCollorScheme=new-object log4net.Appender.ColoredConsoleAppender+LevelColors; $ColConsApndrErrorCollorScheme.level=[log4net.Core.Level]::Error; $ColConsApndrErrorCollorScheme.ForeColor=[log4net.Appender.ColoredConsoleAppender+Colors]::Red;
$ColConsApndr.AddMapping($ColConsApndrErrorCollorScheme);
$ColConsApndrFatalCollorScheme=new-object log4net.Appender.ColoredConsoleAppender+LevelColors; $ColConsApndrFatalCollorScheme.level=[log4net.Core.Level]::Fatal; $ColConsApndrFatalCollorScheme.ForeColor=([log4net.Appender.ColoredConsoleAppender+Colors]::HighIntensity -bxor [log4net.Appender.ColoredConsoleAppender+Colors]::Red);
$ColConsApndr.AddMapping($ColConsApndrFatalCollorScheme);
$ColConsApndr.ActivateOptions();
$ColConsApndr.Threshold = [log4net.Core.Level]::All;
[log4net.Config.BasicConfigurator]::Configure($ColConsApndr);

$Log=[log4net.LogManager]::GetLogger("root");


#--------------------------------------------------------------
# checking parameters
#--------------------------------------------------------------
$log.Info("Reading parameter file '$paramFileName' ....")
try {
    $param = Get-Content -ErrorAction Stop  -Raw -Path $paramFileName | ConvertFrom-Json 
} catch {
    $e = $_.Exception
    $msg = $e.Message
    $log.Fatal("Reading parameter file error: $msg")
    exit
}

if ($param.length -eq 0) {
    $log.Fatal("Reading parameter file error: File is empty!")
    exit    
}

#--------------------------------------------------------------
# starting jobs
#--------------------------------------------------------------
$log.Info("Starting jobs")
$jobIndex = 0
foreach($job in $param) {
    $jobIndex++
    $SourceItems = $null
    $TargetItems = $null
    
    #Get job title for logging
    $jobTitle = $job.Title;
    if (-Not $job.Title) {
        $jobTitle = "SyncJob $JobIndex"
        $log.Warn("Job title is missing - using default '$jobTitle'" )
    }
    $log.Info("$jobTitle : started")

    if ($job.Disabled) {
        $log.Warn("$jobTitle : disabled")
        continue
    }
    $jobSource = $job.Source;
    if (-Not $jobSource) {
        $log.Error("$jobTitle : No Source specified in job definition" )
        continue
    }
    $jobTarget = $job.Target;
    if (-Not $jobTarget) {
        $log.Error("$jobTitle : No Target specified in job definition" )
        continue
    }
    $jobMapping = $job.Mapping
    if (-Not $jobMapping) {
        $log.Error("$jobTitle : No Mapping specified in job definition" )
        continue
    }

    #Find all field names for source and Target
    $SourceFieldNames = [System.Collections.ArrayList]@();
    $TargetFieldNames =[System.Collections.ArrayList]@();

    for ($i = 0; $i -lt $jobMapping.Count; $i++) {
        if ($jobMapping[$i].Source) {
            $SourceFieldNames.Add($jobMapping[$i].Source) > $null
        }
        $TargetFieldNames.Add($jobMapping[$i].Target) > $null
    }

    $log.Info("$jobTitle : found $($SourceFieldNames.Count) source fields.")
    $log.Info("$jobTitle : found $($TargetFieldNames.Count) target fields.")

    #--------------------------------------------------------------
    # reading source items
    #--------------------------------------------------------------
    $log.Info("$jobTitle : Source : reading items ...")
    $SourceItems = $null
    $SourceItems = Get-Data -spec $jobSource -fields $SourceFieldNames -logTitle "$jobTitle : Source"
 

    if ($SourceItems.Length -gt 0) {
        #--------------------------------------------------------------
        # reading current target items
        #--------------------------------------------------------------
        $log.Info("$jobTitle : Target : reading items ...")
        $TargetItems = Get-Data -spec $jobTarget -fields $TargetFieldNames  -logTitle "$jobTitle : Target"   
 
        #--------------------------------------------------------------
        # checking source key fields
        #--------------------------------------------------------------
        $SourceKeyFieldName = $job.Source.KeyFieldName
        if (-Not $SourceKeyFieldName) {
            $log.Error("$jobTitle : Source : missing parameter KeyFieldName")
            continue
        }

        $SourceErrors = $SourceItems | Where-Object { -Not $_.$SourceKeyFieldName } 
        if ($SourceErrors.Length -gt 0) {
            $log.Error("$jobTitle : Source : $($SourceErrors.Length) item with missing value in '$SourceKeyFieldName'")
            continue          
        }

        #--------------------------------------------------------------
        # checking target key fields
        #--------------------------------------------------------------
        $TargetKeyFieldName = $job.Target.KeyFieldName
        if (-Not $TargetKeyFieldName) {
            $log.Error("$jobTitle : Target : missing parameter KeyFieldName")
            continue
        }
        $TargetErrors = $TargetItems | Where-Object { -Not $_.$TargetKeyFieldName } 
        if ($TargetErrors.Length -gt 0) {
            $log.Error("$jobTitle : Target : $($TargetErrors.Length) item with missing value in '$TargetKeyFieldName'")
            continue          
        }


        foreach ($Record in $SourceItems){
            #--------------------------------------------------------------
            # preparing source fields and values 
            #--------------------------------------------------------------
            $SourceKeyValue = $Record.$SourceKeyFieldName
            $SourceValues = @{}
            for ($i = 0; $i -lt $job.Mapping.Count; $i++) {
                if ($job.Mapping[$i].Source) {
                    $SourceFieldName = $job.Mapping[$i].Source
                    $SourceFieldValue = Get-Field-Value-As-String -value $Record.$SourceFieldName -format $job.Mapping[$i].SourceType -attr $job.Mapping[$i].SourceAttr -values $SourceValues
                    if (-Not $SourceValues.$SourceFieldName) {
                        $SourceValues.Add($SourceFieldName, $SourceFieldValue)
                    }
                }
            }

            #--------------------------------------------------------------
            # Checking if target item exists 
            #--------------------------------------------------------------
            $log.Info("$jobTitle : key '$SourceKeyValue' checking target  ....")
            $existingItem = $TargetItems | Where-Object {$_.$TargetKeyFieldName -eq $Record.$SourceKeyFieldName}

            if ($existingItem) {
                #--------------------------------------------------------------
                # target item exists, check for changed values 
                #--------------------------------------------------------------
                $log.Info("$jobTitle : Key '$SourceKeyValue' found in target, cecking changes ...")
                $Changes = @{}

                for ($i = 0; $i -lt $job.Mapping.Count; $i++) {
                    $SourceFieldName = $job.Mapping[$i].Source
                    $TargetFieldName = $job.Mapping[$i].Target
                    if ($SourceFieldName) {
                        # field does exist in source
                        $SourceFieldValue = $SourceValues.$SourceFieldName
                    } else {
                        # field does not come from source (constant value)
                        $SourceFieldValue = $job.Mapping[$i].Value
                    }

                    $TargetFieldValue = Get-Field-Value-As-String -value $existingItem.$TargetFieldName -format $job.Mapping[$i].TargetType -attr $job.Mapping[$i].TargetAttr

                    #--------------------------------------------------------------
                    # compare source and target values 
                    #--------------------------------------------------------------
                    if ($SourceFieldValue -ne $TargetFieldValue) {
                        $log.Info("$jobTitle : Key '$SourceKeyValue' differ $SourceFieldName : '$SourceFieldValue' with $TargetFieldName : '$TargetFieldValue'")
                        # value in source is different than value in target!
                        $Changes.Add($TargetFieldName, $SourceFieldValue)
                    } else {
                        $log.Info("$jobTitle : Key '$SourceKeyValue' equal  $SourceFieldName : '$SourceFieldValue' with $TargetFieldName : '$TargetFieldValue'")
                    }
                }
                $log.Info("$jobTitle : Key '$SourceKeyValue' $($Changes.Count) Changes")
                if ($Changes.Count -gt 0) {
                    #--------------------------------------------------------------
                    # there are changes! update changes in target 
                    #--------------------------------------------------------------
                    Update-Data -Spec $job.Target -SourceItem $Record -TargetItem $existingItem -Changes $Changes -logTitle "$jobTitle : Key '$SourceKeyValue'"
                }
        
            } else {
                #--------------------------------------------------------------
                # target item does not exists, prepare values 
                #--------------------------------------------------------------
                $Values = @{}

                for ($i = 0; $i -lt $job.Mapping.Count; $i++) {
                    $SourceFieldName = $job.Mapping[$i].Source
                    if ($SourceFieldName) {
                        # field exist in source 
                        $Value = $SourceValues.$SourceFieldName
                    } else {
                        # there is no source, this is a constant value!
                        $Value = $job.Mapping[$i].Value
                    }
                    $TargetFieldName = $job.Mapping[$i].Target 
                    $Values.Add($TargetFieldName, $Value )
                }
                #--------------------------------------------------------------
                # insert new item in target
                #--------------------------------------------------------------
                $log.Info("$jobTitle : key '$SourceKeyValue' No existing item in Target")
                Insert-Data -Spec $job.Target -SourceItem $Record -Values $Values -logTitle "$jobTitle : Key '$SourceKeyValue'"
            }
        }
    }

    if ($job.Target.DeletItemsNotInSource) {
        #--------------------------------------------------------------
        # if there are items in target but missing in source
        # delete these items in target 
        #--------------------------------------------------------------
        foreach($targetItem in $TargetItems) {
            $existingSource = $SourceItems | Where-Object {$_.$SourceKeyFieldName -eq $targetItem.$TargetKeyFieldName}
            if (-Not $existingSource) {
                #--------------------------------------------------------------
                # this key value does not exist in Source!
                #--------------------------------------------------------------
                Remove-Data -Spec $job.Target -Item $targetItem -logTitle "$jobTitle"        
            }
        }

    }

    $log.Info("$jobTitle : done")


}


