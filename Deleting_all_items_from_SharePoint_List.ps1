Import-Module 'C:\Users\[USER]\.nuget\packages\microsoft.sharepointonline.csom\16.1.7918.1200\lib\net45\Microsoft.ProjectServer.Client.dll'
Import-Module 'C:\Users\[USER]\.nuget\packages\microsoft.sharepointonline.csom\16.1.7918.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll'
Import-Module 'C:\Users\[USER]\.nuget\packages\microsoft.sharepointonline.csom\16.1.7918.1200\lib\net45\Microsoft.SharePoint.Client.dll'

#Autentication on SP and PPM 
$InstanceURL = "https://[Tenant].sharepoint.com/sites/pwa"
$UserName = "[USER]"
$password = "[PASSWORD]"
$securePass = ConvertTo-SecureString $password -AsPlainText -Force

#Array of items to will be deletes
$ItemsToDelete = New-Object System.Collections.ArrayList

#List
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $securePass)
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($InstanceURL)
$ctx.credentials = $creds



$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $securePass)
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($InstanceURL)
$ctx.credentials = $creds

#Get First Item ID
$list1 = $ctx.Web.Lists.GetByTitle("[SHAREPOINT LIST]")
[Microsoft.SharePoint.Client.CamlQuery]$camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
$camlQuery.ViewXml = "<View><Query><OrderBy><FieldRef Name='ID' Ascending='true'/></OrderBy></Query><RowLimit>1</RowLimit></View>"
$list1Items = $list1.GetItems($camlQuery)
$fields = $list1.Fields
$ctx.Load($list1Items)
$ctx.Load($list1)
$ctx.Load($fields)
$ctx.ExecuteQuery()
$ID = $list1Items[0].Id
$list1Items.Count

#Assign a variable to control the end of list
$bEndOfList = $False  


While(! $bEndOfList) {

    #Get List Items Into a List Filter Collection
    [Microsoft.SharePoint.Client.CamlQuery]$camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
    $camlQuery.ViewXml = "<View><Query><Where><Geq><FieldRef Name='ID'/><Value Type='Number'>$ID</Value></Geq></Where></Query><RowLimit>5000</RowLimit></View>"
    $list1Items = $list1.GetItems($camlQuery)
    $fields = $list1.Fields
    $ctx.Load($list1Items)
    $ctx.Load($Fields)
    $ctx.ExecuteQuery()

    for($i = $list1Items.Count; $i -ge 0; $i--){ 
       
       if($list1Items.Item($i)["ProjectId"] -eq $null){ 

            $list1Items.Item($i)["AccountId"]

            $list1Items[$i].DeleteObject()
            $ctx.ExecuteQuery()

       }else{

        if($list1Items.Count -eq 0){
            
            $bEndOfList = $true

        }


       }

       
    }


}

