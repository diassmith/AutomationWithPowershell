$time = Get-Date

Write-Host $time -ForegroundColor Green
Import-Module 'C:\Users\[USER]\.nuget\packages\microsoft.sharepointonline.csom\16.1.7918.1200\lib\net45\Microsoft.ProjectServer.Client.dll'
Import-Module 'C:\Users\[USER]\.nuget\packages\microsoft.sharepointonline.csom\16.1.7918.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll'
Import-Module 'C:\Users\[USER]\.nuget\packages\microsoft.sharepointonline.csom\16.1.7918.1200\lib\net45\Microsoft.SharePoint.Client.dll'


#Autentication on Share Point Online and Microsoft Project Online 
$InstanceURL = "https://[TENANT]/sites/pwa"
$UserName = "[EmailUSER]"
$password = "[PASSWORD]"
$securePass = ConvertTo-SecureString $password -AsPlainText -Force

#Array of items to will be deletes 
$ItemsToDelete = New-Object System.Collections.ArrayList

#Credentials
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $securePass)

#Lists
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($InstanceURL)
$ctx.credentials = $creds
$ctx2 = New-Object Microsoft.SharePoint.Client.ClientContext($InstanceURL)
$ctx2.credentials = $creds


#Get List and the first Item ID
$list1 = $ctx.Web.Lists.GetByTitle("[ShapointList1Name]")
[Microsoft.SharePoint.Client.CamlQuery]$camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
$camlQuery.ViewXml = "<View><Query><OrderBy><FieldRef Name='ID' Ascending='true'/></OrderBy></Query><RowLimit>1</RowLimit></View>"
$list1Items = $list1.GetItems($camlQuery)
$ctx.Load($list1Items)
$ctx.Load($list1)
$ctx.ExecuteQuery()
$ID = $list1Items[0].Id

#Get List
$list2 = $ctx2.Web.Lists.GetByTitle("[ShapointList2Name]")
$ctx2.Load($list2)
$ctx2.ExecuteQuery()
    
#Assign a variable to control the end of list
$bEndOfList = $False
    
#While is not the end of list
While(! $bEndOfList) {   

    #Get List Items Into a List Filter Collection
    [Microsoft.SharePoint.Client.CamlQuery]$camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
    $camlQuery.ViewXml = "<View><Query><Where><Geq><FieldRef Name='ID'/><Value Type='Number'>$ID</Value></Geq></Where></Query><RowLimit>5000</RowLimit></View>"
    $list1Items = $list1.GetItems($camlQuery)
    $ctx.Load($list1Items)
    $ctx.ExecuteQuery()

    for($i = 0 ; $i -lt $list1Items.Count;$i++){
        Write-Host $i

        $ListItemCreationInformation = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
        $addNewItem = $list2.AddItem($ListItemCreationInformation)

        $addNewItem["Title"] = $list1Items.Item($i)["Title"]
        $addNewItem["AccountId"] = $list1Items.Item($i)["AccountId"]
        $addNewItem["budgetCost"] = $list1Items.Item($i)["TotalBudgetCost"]
        $addNewItem["ForeCastCost"] = $list1Items.Item($i)["TotalForecastCost"]
        $addNewItem["ActualCost"] = $list1Items.Item($i)["TotalActualCost"]
        $addNewItem.Update()
        $ctx2.ExecuteQuery() 

        $ListItemCreationInformation = $null
        $addNewItem = $null

    }

    #Check if count of List Filter Collection is lower than 5000
    if($list1Items.Count -lt 5000) {
        $bEndOfList = $True
    } else {            
        $ID = $list1Items.Item($i -1)["ID"] + 1
    }
}   
$time = Get-Date
Write-Host $time -ForegroundColor Green