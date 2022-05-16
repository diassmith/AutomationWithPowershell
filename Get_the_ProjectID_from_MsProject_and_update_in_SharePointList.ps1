$sLogFileName = "list"
$sLogFileExtension = ".txt"
$sLogFilePath = "C:\ProgramData\LOG_ProjectID\"


#Creating Log system
if(Test-Path "$($sLogFilePath)$($sLogName)$($sLogFileExtension)" -PathType Leaf){


    $oLogFile = Get-Item "$($sLogFilePath)$($sLogName)$($sLogFileExtension)"

    if ($oLogFile.Length -ge 14680064) {

        $aLog = New-Object System.Collections.ArrayList
        $aLog = Get-Childitem $sLogFilePath

        $iActualLogIndex = 0
        $iGreaterLogIndex = 0

        $i = 0

        while($i -lt $aLog.Count){
            if ($aLog[$i].Name.IndexOf($sLogName) -eq 0) {
                $iActualLogIndex = $aLog[$i].Name.Replace($sLogName, "").Replace($sLogFileExtension, "")
                if ($iActualLogIndex -match '^\d+$') {
                    $iActualLogIndex = $iActualLogIndex.ToInt32($null)        
                    if ($iActualLogIndex -gt $iGreaterLogIndex) {
                        $iGreaterLogIndex = $iActualLogIndex
                    }
                }
            }
            $i++
        }

        $iGreaterLogIndex++

        if ($iGreaterLogIndex -gt 0 -and $iGreaterLogIndex -lt 10) {
            Rename-Item -Path "$($sLogFilePath)$($sLogName)$($sLogFileExtension)" -NewName "$($sLogFilePath)$($sLogName)00$($iGreaterLogIndex)$($sLogFileExtension)"
        } elseif ($iGreaterLogIndex -ge 10 -and $iGreaterLogIndex -lt 100) {
            Rename-Item -Path "$($sLogFilePath)$($sLogName)$($sLogFileExtension)" -NewName "$($sLogFilePath)$($sLogName)0$($iGreaterLogIndex)$($sLogFileExtension)"
        } elseif ($iGreaterLogIndex -le 999) {
            Rename-Item -Path "$($sLogFilePath)$($sLogName)$($sLogFileExtension)" -NewName "$($sLogFilePath)$($sLogName)$($iGreaterLogIndex)$($sLogFileExtension)"
        } else {
            Rename-Item -Path "$($sLogFilePath)$($sLogName)$($sLogFileExtension)" -NewName "$($sLogFilePath)$($sLogName)01$($sLogFileExtension)"
        }
        New-Item -Path "$($sLogFilePath)$($sLogFileName)$($sLogFileExtension)" | Out-Null
    }

}else{

    Start-Transcript -Append -Path "$($sLogFilePath)$($sLogFileName)$($sLogFileExtension)"

}

$time = Get-Date -format yyyyMMddHHmmssffff
Write-Host $time ": Starting looking to find PorjectId"

#Importing CSOM module 
Import-Module 'C:\Users\[USER]\.nuget\packages\microsoft.sharepointonline.csom\16.1.7918.1200\lib\net45\Microsoft.ProjectServer.Client.dll'
Import-Module 'C:\Users\[USER]\.nuget\packages\microsoft.sharepointonline.csom\16.1.7918.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll'
Import-Module 'C:\Users\[USER]\.nuget\packages\microsoft.sharepointonline.csom\16.1.7918.1200\lib\net45\Microsoft.SharePoint.Client.dll'

#Autentication on SP and PPM 
$InstanceURL = "https://ambientalcorp.sharepoint.com/sites/pwa"
$UserName = "[USER]"
$password = "[PASSAWORD]("
$securePass = ConvertTo-SecureString $password -AsPlainText -Force

#Array of items to will be deletes
$ItemsToDelete = New-Object System.Collections.ArrayList

#List
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $securePass)
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($InstanceURL)
$ctx.credentials = $creds

#Authentication in Project
$projContext = New-Object Microsoft.ProjectServer.Client.ProjectContext($InstanceURL)
[Microsoft.SharePoint.Client.SharePointOnlineCredentials]$spocreds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $securePass); 
$projContext.Credentials = $spocreds

#Reading all project
$time = Get-Date -format yyyyMMddHHmmssffff
Write-Host $time ": Reading all Projects"
$projContext.Load($projContext.Projects)
$projContext.ExecuteQuery()
$time = Get-Date -format yyyyMMddHHmmssffff
Write-Host $time ": Done, read all projects"

#Seting Variables of Control
$numberOfProject= $projContext.Projects.Count
$control_Project = $false
$count_Project = 0

#Attributing projects to a Variable
$Projects = $projContext.Projects

#Starting DataTable
$dt = New-Object System.Data.DataTable

#Add columns
$dt.Columns.Add("AccountId", "string")| Out-Null
$dt.Columns.Add("ProjectId", "string")| Out-Null
$dt.Columns.Add("Title", "string")| Out-Null


#Running all projects
Foreach ($Project in $Projects) {

        
        #Print project name and Id 
        $time = Get-Date -format yyyyMMddHHmmssffff
        Write-Host $time ": -----------------------------------------------------"
        Write-Host $time ": NProject Name:" $Project.Name.ToString() -ForegroundColor Green
        Write-Host $time ": Project ID:  " $Project.Id.ToString()   -ForegroundColor Green
        Write-Host $time ": -----------------------------------------------------"
        
        $time = Get-Date -format yyyyMMddHHmmssffff
        Write-Host $time ": Reading all tasks from projects" $Project.Name.ToString()
        
        #Begin object ProjectServer, Load projects,Project Id and load all project tasks
        $projContext = New-Object Microsoft.ProjectServer.Client.ProjectContext($InstanceURL)
        [Microsoft.SharePoint.Client.SharePointOnlineCredentials]$spocreds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $securePass); 
        $projContext.Credentials = $spocreds
        $projContext.Load($projContext.Projects)
        $projContext.ExecuteQuery()
        $draftProject = $projContext.Projects.GetByGuid($Project.id)
        $projContext.Load($draftProject.Tasks)
        $projContext.ExecuteQuery()

        $time = Get-Date -format yyyyMMddHHmmssffff
        Write-Host $time ": Done, read all task from the projects" $Project.Name.ToString()
        
        #Attributing all tasks to a Variable
        $task = $draftProject.Tasks  

        for($i = 0; $i -lt $task.Count; $i++){
                
            #if Account of project is equals Sharepoint List, search the tasks and Inserting project Id on reference task on list  
            if($task.Item($i)["Custom_f85bd515b274e81180dc00155d40d102"] -ne $null){

                $row = $dt.NewRow()
                $row["AccountId"] = $task.Item($i)["Custom_f85bd515b274e81180dc00155d40d102"]
                $row["ProjectId"] = $Project.Id.ToString()
                $row["Title"] = $Project.Name.ToString()
                $dt.Rows.Add($row)
                $control = $true

            }
        }

        $dw  = New-Object System.Data.DataView($dt)
        $dw.Sort="AccountId ASC"

        
        #$dw | Format-Table -AutoSize  # Or $dw | Out-GridView

        $dt.Close
        $dw.close


    }

    $projContext = $null

#Cleaning variable draftproject
if ($draftProject -ne $null) {

    $draftProjectl = $null

}

$time = Get-Date -format yyyyMMddHHmmssffff
Write-Host $time ": Finish the Data table and Data view" -ForegroundColor Green

$dw | Format-Table -AutoSize  # Or $dw | Out-GridView

$time = Get-Date
Write-Host $time "Getting the final view"

#Get First Item ID
$list1 = $ctx.Web.Lists.GetByTitle("[Sharepoint List]")
[Microsoft.SharePoint.Client.CamlQuery]$camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
$camlQuery.ViewXml = "<View><Query><OrderBy><FieldRef Name='ID' Ascending='true'/></OrderBy></Query><RowLimit>1</RowLimit></View>"
$list1Items = $list1.GetItems($camlQuery)
$fields = $list1.Fields
$ctx.Load($list1Items)
$ctx.Load($list1)
$ctx.Load($fields)
$ctx.ExecuteQuery()
$ID = $list1Items[0].Id

#Assign a variable to control the end of list
$bEndOfList = $False

#While not is the end of list
While(! $bEndOfList) {

    #Get List Items Into a List Filter Collection
    [Microsoft.SharePoint.Client.CamlQuery]$camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
    $camlQuery.ViewXml = "<View><Query><Where><Geq><FieldRef Name='ID'/><Value Type='Number'>$ID</Value></Geq></Where></Query><RowLimit>5000</RowLimit></View>"
    $list1Items = $list1.GetItems($camlQuery)
    $fields = $list1.Fields
    $ctx.Load($list1Items)
    $ctx.Load($Fields)
    $ctx.ExecuteQuery()

    for($i = 0; $i -lt $list1Items.Count; $i++){
            
            $time = Get-Date
            Write-Host $time "identifying  ProjectId"

            $dwFilteredRows = $dw.where({$_.AccountId -eq $list1Items.Item($i)["AccountId"]})

            Write-Host "View " $dw.Count
            Write-Host "Where "$dwFilteredRows.Count

            #checking if the AccountId exist in DataView
            if($dwFilteredRows.Count -gt 0){
                
                Write-Host " AccountId " $list1Items.Item($i)["AccountId"] -ForegroundColor Green
                Write-Host " View " $dwFilteredRows.Item(0)["AccountId"] -ForegroundColor Green

                $list1Items.Item($i)["ProjectId"] = $dwFilteredRows.Item(0)["ProjectId"]
                
                $list1Items[$i].Update()
                $ctx.ExecuteQuery()


                Write-Host "Qtd Lines " $dt.Rows.Count -ForegroundColor Yellow
                Write-Host "AccountId with filter " $dwFilteredRows.Item(0)["AccountId"] -ForegroundColor Yellow
                Write-Host  "AccountId SharePoint "$list1Items.Item($i)["AccountId"] -ForegroundColor Yellow

                $linhas = $dt.Select("AccountId = '" + $dwFilteredRows.Item(0)["AccountId"] + "'")
                
                for($linha = 0; $linha -lt $linhas.Count){

                    $linhas[$linha].Delete()

                    $linha++

                }

                Write-Host $dt.Rows.Count -ForegroundColor Green
            }else{

                Write-Host " AccountId " $list1Items.Item($i)["AccountId"] -ForegroundColor Green

                Write-Host " UA não existe no PWA " -ForegroundColor Green
                
                $list1Items[$i].DeleteObject()

                $ctx.ExecuteQuery()


            }

    }

    $dwFilteredRows.Clear()

    $time = Get-Date


    #Check if count of List Filter Collection is lower than 5000
    if($list1Items.Count -lt 5000) {
        $bEndOfList = $True
     } else {            
        $ID = $list1Items.Item($i -1)["ID"] + 1
     }
}

$time = Get-Date -format yyyyMMddHHmmssffff
Write-Host $time ": Done the update of ProjectId" -ForegroundColor Green



Stop-Transcript
Add-PSSnapin Microsoft.SharePoint.PowerShell -EA silentlycontinue