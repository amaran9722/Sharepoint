<# References - 
http://blogs.technet.com/b/fromthefield/archive/2014/02/18/office365-script-to-create-a-list-add-fields-and-change-the-default-view-using-csom.aspx
https://peterheibrink.wordpress.com/2014/05/19/create-lookup-field-using-powershell-and-csom/  - create a lookup column
https://www.itunity.com/article/loading-specific-values-lambda-expressions-sharepoint-csom-api-windows-powershell-1249 - Lamba expression in powershell
https://gist.githubusercontent.com/vman/51ae46fbb439fbf610c8/raw/d9af0763a0745aeb273517bc691ab5fd4955be43/TaxonomyTermSearch.js - Searching a term inside termset
#>


#1) Load the input values for powershell - this can be moved to a configuration file


$blogurl = "https://intranet.kolabrate.com.au/localnews"
$username = "barryj"
$password = "demo-1234"
$domain = "kolabrate"
$PageLayoutName = "Corporate News Page Layout.aspx"
$csvpath = "csv/LocalNews.csv"
$mmsGroupName = "Kolabrate"

#2) Add sharepoint client dlls

Add-Type -Path "$pwd\SharePointDll\Microsoft.SharePoint.Client.dll"
Add-Type -Path "$pwd\SharePointDll\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "$pwd\SharePointDll\Microsoft.SharePoint.Client.Taxonomy.dll"



#3) Load necessary scripts

. "$pwd\RemoveSpecialCharacters.ps1"
. "$pwd\LoadSPProperties.ps1"


 #4) Get Blog Post client context and load castto for Lookup

    $clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($blogurl)
    $rootWeb = $clientContext.web
    $clientContext.Load($rootWeb)
    $clientContext.ExecuteQuery()
   

    # load the field collection  - Use 'GetByInternalNameOrTitle' to get the handle on field
    $BlogList = $rootWeb.Lists.GetByTitle('Posts')
    $clientContext.load($BlogList.Fields)
    $clientContext.ExecuteQuery()

    # load the castto for lookup, MMS
    $castToMethodGeneric = [Microsoft.SharePoint.Client.ClientContext].GetMethod("CastTo")
    $castToMethodLookup = $castToMethodGeneric.MakeGenericMethod([Microsoft.SharePoint.Client.FieldLookup])
    $castToMethodTaxonomy = $castToMethodGeneric.MakeGenericMethod([Microsoft.SharePoint.Client.Taxonomy.TaxonomyField])
    
    #Lookup List
    $lookupfield = $BlogList.Fields.GetByInternalNameOrTitle("PostCategory")
    $lookupfield = $castToMethodLookup.Invoke($clientContext, $lookupfield)
    $clientContext.Load($lookupfield)
    $clientContext.executeQuery()  # lookup field 'PostCategory' is loaded completely
    
    $LookupList = $rootWeb.Lists.GetById([Guid]$lookupfield.LookupList)
    $camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
    $camlQuery.ViewXml = '<View Scope="RecursiveAll"><Query><Where><Eq><FieldRef Name="Title" /><Value Type="Text">Local News</Value></Eq></Where></Query></View>'    
    $coll =  $LookupList.GetItems($camlQuery)
    $clientContext.load($coll)
    $clientContext.executeQuery()


    #Taxonomy - Location Field

        <#1) Get clientcontext of parentweb & taxonomysession
          2) get the taxonomy session
          3) Get the termGroups
          4) Get the termset #>

    $rooturl = "https://" + ([System.Uri]$blogurl).Host 
    $taxonomy = Get-SPTaxonomySession -Site $rooturl
    # Retrieve 'Locations' Termgroup
    $LocationTermGroup = $taxonomy.TermStores.Groups | Where-Object {$_.Name -eq $mmsGroupName} 
    $LocationTermSet  = $LocationTermGroup.TermSets | Where-Object {$_.Name -eq "Locations"} 
    # Retrieve 'Division' Termgroup
    $DivisionTermGroup = $taxonomy.TermStores.Groups | Where-Object {$_.Name -eq $mmsGroupName} 
    $DivisionTermSet  = $LocationTermGroup.TermSets | Where-Object {$_.Name -eq "Division"} 
    
    
    
    #### pending fields - Note(Rich Text Field ), Currency, Percent,MultiChoice,Person,MultiPerson,Hyperlink,MultipleValue MMS,Calculated fields, external content types


#5) Import the localnews content 
    $LocalNews = Import-Csv $csvpath


 #6) Create blog  post article

   foreach($News in $LocalNews)
   {

        # Title,Body,PostCategory,Division,Featured,Include_x0020_in_x0020_Newslette,Locations

        #Update Blog Title & Body
        $ListItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
        $Item = $BlogList.AddItem($ListItemInfo)
        $Item["Title"] = $News.Title
        $Body = Remove-Chars $News.Body
        $Item["Body"] = $Body
        
        
        #Update Choice ( yes/No) column
        $FeaturedNews = If ($News.Featured -eq "Yes" ) {"$true"} Else {"$false"}
        $Item["Featured"] = $FeaturedNews
        $IncludeInNewsLetter = If ($News.Include_x0020_in_x0020_Newslette -eq "Yes" ) {"$true"} Else {"$false"}
        $Item["Include_x0020_in_x0020_Newslette"] = $IncludeInNewsLetter

      
        #Assign Lookup Column
        $CategoryItemLookupField = New-Object Microsoft.SharePoint.Client.FieldLookupValue
        $CategoryItemLookupField.set_lookupId($coll[0].id);
        $Item.set_item('PostCategory', $CategoryItemLookupField);
   
        
        #Update Date Column
        $date = Get-date
        $Item["PublishedDate"] = $Date

        #Update MMS Column - Locations
        $loctermSetId = $LocationTermSet.Terms | Where-Object {$_.Name -eq $News.Locations} | Select-Object -ExpandProperty Id | Select-Object -ExpandProperty Guid
        $item["Locations"] = $loctermSetId

        #Update MMS Column - Divsions
        $DivtermSetId = $DivisionTermSet.Terms | Where-Object {$_.Name -eq $News.Division} | Select-Object -ExpandProperty Id | Select-Object -ExpandProperty Guid
        $item["Division"] = $DivtermSetId


        <#Update the author
        $authorname = $rootweb.EnsureUser($News.Author)
        $item.Set_Item("Author", $authorname) #>
        

        #update the list
        $Item.Update()
        $clientContext.Load($Item)
        $clientContext.ExecuteQuery()
        


       

   } 






 

 










