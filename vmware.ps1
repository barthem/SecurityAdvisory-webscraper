# import selenium module
import-module selenium 

#functie om advisory naar teams te posten, moet in de loop geplaatst worden omdat die parralel draait. 
$webhookUrl = "https://ogd.webhook.office.com/webhookb2/e268ccf5-d2fd-4cb2-abbe-90f7fe7720ea@afca0a52-882c-4fa8-b71d-f6db2e36058b/IncomingWebhook/7e3718e9645a452aad07d2f6127ea978/72a09449-2f5e-491b-8cb4-2c2ebf993f92"
function Send-TeamsNotification {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Issue,
    
        [Parameter(Mandatory = $true)]
        [string]$WebhookUrl
    )
    
    # Build the JSON body using the Issue object
    $JSONBody = [PSCustomObject][Ordered]@{
        "@type"      = "MessageCard"
        "@context"   = "http://schema.org/extensions"
        "summary"    = "New VMWare Notification"
        "themeColor" = '0078D7'
        "title"      = "$($Issue.Title)"
        "text"       = @"
<b>Document ID:</b> $($Issue.DocumentID)<br>
<b>Document URL:</b> <a href='$($Issue.DocumentURL)'>$($Issue.DocumentURL)</a><br>
<b>Support Products:</b> $($Issue.SupportProducts)<br>
<b>Severity:</b> $($Issue.Severity)<br>
<b>Published Date:</b> $($Issue.PublishedDate)<br>
<b>Last Updated:</b> $($Issue.LastUpdated)
"@
    }
    $TeamMessageBody = ConvertTo-Json -InputObject $JSONBody -Depth 10 -Compress

    # Set up parameters for Invoke-RestMethod
    $parameters = @{
        URI         = $WebhookUrl
        Method      = 'POST'
        Body        = $TeamMessageBody
        ContentType = 'application/json'
    }

    # Send the notification
    Invoke-RestMethod @parameters
}
# }
#urls gevonden op https://blogs.vmware.com/security/2024/05/where-did-my-vmware-security-advisories-go.html
$url = "https://support.broadcom.com/web/ecx/security-advisory?segment=VC"

#start firefox webengine, maak hem headless (geen gui) omndat die draait als pipeline
$Driver = Start-SeFirefox -headless

#navigeer naar VMWare advisory page
$Driver.Navigate().GoToUrl($url)

# $Rows = $table.FindElementByTagName("tr")
# $rows.GetAttribute("outerHTML")


#try reading the pbody max 5 times. 
$i = 0
do {
    write-host "reading tbody"
    $tablecontent = $Driver.FindElementByTagName("tbody")
    # $tablecontent.GetAttribute("outerHTML")
    $i += 1

    #kill loop if element fails
    if ($i -gt 5) {
        throw "coudlnt read tbody"
        exit 
    }

} until (      
    <# Condition that stops the loop if it returns true #>
    $null -ne $tablecontent.displayed       
)

#vraag table html code op 
$tablecontent = $Driver.FindElementByTagName("tbody")

# $tablecontent.GetAttribute("outerHTML")

#lees de individuele table telements uit. 
$Rows = $tableContent.FindElementsByTagName("tr")

#parse alle individuele table elements, indien die van vandaag waren post hem in teams. 
$issues = $Rows | ForEach-Object -ThrottleLimit 10 -Parallel {
   

    #parse row naar individuele cellen met info
    $Cells = $_.FindElementsByTagName("td")

    # $Cells = $Rows[$iii].FindElementsByTagName("td")

    # lees de URL uit van de eerste element die lijdt naar advisory 
    $DocumentLink = $Cells[0].FindElementByTagName("a")

    # parse de cellen van de advisory table row. 
    $issue = [PSCustomObject]@{
        DocumentID      = $Cells[0].Text
        DocumentURL     = $DocumentLink.GetAttribute("href")
        Title           = $Cells[1].Text
        SupportProducts = $Cells[2].Text
        Severity        = $Cells[3].Text
        PublishedDate   = $Cells[4].Text
        LastUpdated     = $Cells[5].Text
    }
    write-host "done parsing $($issue.DocumentID)"
    $issue
}

write-host "closing webdriver."
$Driver.quit()


# vraag de datum vandaag op. 
$dateoftoday = (get-date).ToString("dd MMMM yyyy")
write-host "date of today: $dateoftoday "


#check if one of the security advisories was of today
$issues | ForEach-Object{
    $currentSecurityAdvisory = $_

    write-host "processing $($currentSecurityAdvisory.DocumentID).`nrelease date: $($currentSecurityAdvisory.PublishedDate)`n"
    if ($dateoftoday -eq $currentSecurityAdvisory.PublishedDate ) {
        write-host "found new advisory. posting it to teams. "
        Send-TeamsNotification -Issue $currentSecurityAdvisory -WebhookUrl $webhookUrl
    }
}