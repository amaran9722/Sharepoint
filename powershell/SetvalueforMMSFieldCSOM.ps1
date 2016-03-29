<# References - 
    https://gist.githubusercontent.com/vman/51ae46fbb439fbf610c8/raw/d9af0763a0745aeb273517bc691ab5fd4955be43/TaxonomyTermSearch.js - Searching a term inside termset
#>


#1) Load sharepoint client dlls

    Add-Type -Path "$pwd\SharePointDll\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "$pwd\SharePointDll\Microsoft.SharePoint.Client.Runtime.dll"
    Add-Type -Path "$pwd\SharePointDll\Microsoft.SharePoint.Client.Taxonomy.dll"



Function SetvalueForTaxonomyField{

Param(

  [string]$blogurl, # URL of site/Web
  [string]$list, # name of the list that contains MMS column
  [string]$username,
  [string]$password,
  [string]$domain,
  [string]$mmsGroupName, # term group name
  [string]$mmstermsetName, # term set name
  [string]$mmstermname # Actual term name

)


#2) Get client context & termsets

    $clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($blogurl)
    $rootWeb = $clientContext.web
    $clientContext.Load($rootWeb)
    $clientContext.ExecuteQuery()
   

    # load the field collection
    $BlogList = $rootWeb.Lists.GetByTitle($list)
    $clientContext.load($BlogList.Fields)
    $clientContext.ExecuteQuery()

    # load the castto for MMS
    $castToMethodGeneric = [Microsoft.SharePoint.Client.ClientContext].GetMethod("CastTo")
    $castToMethodTaxonomy = $castToMethodGeneric.MakeGenericMethod([Microsoft.SharePoint.Client.Taxonomy.TaxonomyField])
    
    # Get the taxonomy Session - parent site collection
    $rooturl = "https://" + ([System.Uri]$blogurl).Host 
    $taxonomy = Get-SPTaxonomySession -Site $rooturl
    
    # Retrieve  Termgroup & Termset
    $TermGroup = $taxonomy.TermStores.Groups | Where-Object {$_.Name -eq $mmsGroupName} 
    $TermSet  =  $TermGroup.TermSets | Where-Object {$_.Name -eq $mmstermsetName} 
    
    #Update MMS Column  
    
    #option 1 : Find the term ID & update the list column - uncomment the below line
    #$loctermSetId = $LocationTermSet.Terms | Where-Object {$_.Name -eq $mmstermsetName} | Select-Object -ExpandProperty Id | Select-Object -ExpandProperty Guid

    #option 2 : Searching all the terms and updating the list column - As mentioned in the reference article - by Vardhaman Deshpande
    $lmi =  New-Object Microsoft.SharePoint.Client.Taxonomy.LabelMatchInformation.newObject($clientContext);
    $strmatch = New-Object Microsoft.SharePoint.Client.Taxonomy.StringMatchOption($clientContext)

        
    $lmi.set_termLabel("a"); #search terms.
    $lmi.set_defaultLabelOnly($true);
    $lmi.set_stringMatchOption($strmatch.startsWith);
    $lmi.set_resultCollectionSize(10); # total number of terms to bring back
    $lmi.set_trimUnavailable($true);

    $terms = $TermSet.getTerms($lmi);

    
    $item = $BlogList.Items[0] # This is hardcoded at the moment, but csv can be utilised to loop through all items
    $item["Locations"] = $terms 


    #update the list
    $Item.Update()
    $clientContext.Load($Item)
    $clientContext.ExecuteQuery()
        


}



#execute the main function 
SetvalueForTaxonomyField "https://intranet.kolabrate.com.au/localnews" "Posts" "barryj" "demo-1234" "kolabrate" "kolabrate" "Locations" "Sydney"
    










