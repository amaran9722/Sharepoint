<# references : 
https://sharepointstew.wordpress.com/2015/03/09/csom-powershell-to-create-a-sharepoint-publishing-page-with-custom-or-oob-page-layout/
https://blogs.msdn.microsoft.com/prasannabalajim/2013/05/20/how-to-add-a-publishing-page-in-sharepoint-2013-using-client-object-model/
http://www.mavention.com/blog/provisioning-publishing-pages-powershell
http://www.c-sharpcorner.com/blogs/powershell-script-to-create-publishing-page-using-custom-pagelayout-in-sharepoint-2013
#>


#1) Load the input values for powershell - this can be moved to a configuration file

$url = "https://intranet.kolabrate.com.au"
$username = "barryj"
$password = "demo-1234"
$domain = "kolabrate"
#$PageLayoutName = "Corporate News Page Layout.aspx"
$PageLayoutName = "Corporate News Page Layout.aspx"
$csvpath = "csv/CorporateNewsNew.csv"


#2) Add sharepoint client dlls

Add-Type -Path "$pwd\SharePointDll\Microsoft.SharePoint.Client.dll"
Add-Type -Path "$pwd\SharePointDll\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "$pwd\SharePointDll\Microsoft.SharePoint.Client.Publishing.dll"


#3) Load special characters script

. "$pwd\RemoveSpecialCharacters.ps1"

$clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($url) 

<#4) Mixed mode authentication

$clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($url);
$credentials = New-Object System.Net.NetworkCredential($username, $password, $domain);
$clientContext.Credentials = $credentials;
$clientContext.AuthenticationMode = [Microsoft.SharePoint.Client.ClientAuthenticationMode]::Default

Function HandleMixedModeWebApplication()
{
  param([Parameter(Mandatory=$true)][object]$clientContext)  
  Add-Type -TypeDefinition @"
  using System;
  using Microsoft.SharePoint.Client;
   
  namespace Toth.SPOHelpers
  {
      public static class ClientContextHelper
      {
          public static void AddRequestHandler(ClientContext context)
          {
              context.ExecutingWebRequest += new EventHandler<WebRequestEventArgs>(RequestHandler);
          }
   
          private static void RequestHandler(object sender, WebRequestEventArgs e)
          {
              //Add the header that tells SharePoint to use Windows authentication.
              e.WebRequestExecutor.RequestHeaders.Remove("X-FORMS_BASED_AUTH_ACCEPTED");
              e.WebRequestExecutor.RequestHeaders.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
          }
      }
  }
"@ -ReferencedAssemblies "$pwd\SharePointDll\Microsoft.SharePoint.Client.dll", "$pwd\SharePointDll\Microsoft.SharePoint.Client.Runtime.dll";
  [Toth.SPOHelpers.ClientContextHelper]::AddRequestHandler($clientContext);
}
HandleMixedModeWebApplication $clientContext; #>


 #5) Get 'corporate news' page layout


    $rootWeb = $clientContext.Site.RootWeb
    $clientContext.Load($rootWeb)
    $clientContext.ExecuteQuery()
    
    $mpList = $rootWeb.Lists.GetByTitle('Master Page Gallery')
    $camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
    $camlQuery.ViewXml = '<View Scope="RecursiveAll"><Query><Where><Eq><FieldRef Name="FileLeafRef" /><Value Type="Text">Corporate News Page Layout.aspx</Value></Eq></Where></Query></View>'
    $items = $mpList.GetItems($camlQuery)
    $clientContext.Load($items)
    $clientContext.ExecuteQuery()

    $tpLayoutItem = $items[0]
    $clientContext.Load($tpLayoutItem)

    $web = $clientContext.Web
    $pubWeb = [Microsoft.SharePoint.Client.Publishing.PublishingWeb]::GetPublishingWeb($clientContext, $web)
    $clientContext.Load($pubWeb)

#6) Import the content csv
    
    $CorpNews = Import-Csv $csvpath


 #7)  Set the counter for looping images

   $counter = 1;

 #8) Create the page layout
   foreach($News in $CorpNews)
   {

        
        $pubPageInfo = New-Object Microsoft.SharePoint.Client.Publishing.PublishingPageInformation
        $pubPageInfo.Name = $News.PageName + ".aspx"
        $pubPageInfo.PageLayoutListItem = $tpLayoutItem 
        $pubPage = $pubWeb.AddPublishingpage($pubPageInfo)

        
        $clientContext.Load($pubPage)
        $clientContext.ExecuteQuery()


        #Page added. Now retrieve list item, check it out, update title and page content, and check back in.
        $listItem = $pubPage.get_listItem()
        $clientContext.Load($listItem)
        $clientContext.ExecuteQuery()

       
        # Set the attributes
        $listItem.Set_Item("Title", $News.Headline)
        $PageContent = Remove-Chars $News.PageContent
        $listItem.Set_Item("PublishingPageContent", $PageContent)
        $listItem.Set_Item("PublishingImageCaption", $News.ImageCaption)
        $date = Get-date
        $listItem.Set_Item("ArticleStartDate", $date) 
        
        
        $Image = "img (" + $counter  + ").jpg"
        $PageImage = "<img src='/PublishingImages/" + $Image  + "' width='370px' height ='370px' />"
        $RollupImage = "<img src='/PublishingImages/" + $Image  + "' width='180px' height ='180px' />" 

    
        
        
        
        $listItem.Set_Item("PublishingPageImage", $PageImage)
        $listItem.Set_Item("PublishingRollupImage", $RollupImage)
        $FlashNews=  If ($News.FlashNews -eq "Yes" ) {"$true"} Else {"$false"}
        $FeaturedNews = If ($News.FeaturedNews -eq "Yes" ) {"$true"} Else {"$false"}
        $listItem.Set_Item("Flash_x0020_News", $FlashNews)
        $listItem.Set_Item("Featured", $FeaturedNews)
        $authorname = $web.EnsureUser($News.Author)
        $listItem.Set_Item("PublishingContact", $authorname)
        $listItem.Set_Item("Author", $authorname)
        $listItem.Set_Item("Editor", $authorname)
        

        #Update list item
        $listItem.Update()
        $counter++
        
       
        # Check if content required for approval is checked or unchecked
        $listItem.File.CheckIn("", [Microsoft.SharePoint.Client.CheckinType]::MajorCheckIn)
        $listItem.File.Publish("")
       # $listItem.File.Approve("")
        $clientContext.Load($listItem)
        $clientContext.ExecuteQuery()


   }



 

 










