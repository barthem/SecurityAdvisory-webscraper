# import selenium module
import-module selenium 

# $currentDate = (Get-Date).ToString("dd MMMM yyyy")

#urls gevonden op https://blogs.vmware.com/security/2024/05/where-did-my-vmware-security-advisories-go.html
$url = "https://support.broadcom.com/web/ecx/security-advisory?segment=VC"

#start firefox webengine, maak hem headless (geen gui) omndat die draait als pipeline
$Driver = Start-SeFirefox -headless

#navigeer naar VMWare advisory page
$Driver.Navigate().GoToUrl($url)

# Find the table by ID or class (adjust this according to your specific case)
# $Table = $Driver.FindElementById("tableId")  # or FindElementByClassName if using a class
# $Table = $Driver.FindElementsByClassName("table")
# Get the raw HTML of the table
# $tableHTML = $Table.GetAttribute("outerHTML")
# $Rows = $table.FindElementByTagName("tr")

# $Rows = $table.FindElementByTagName("tr")
# $rows.GetAttribute("outerHTML")
# FindElementsByTagName


#vraag table html code op 
$tablecontent = $Driver.FindElementByTagName("tbody")
# $tablecontent.GetAttribute("outerHTML")

#lees de individuele table telements uit. 
$Rows = $tableContent.FindElementsByTagName("tr")

#parse alle individuele table elements, indien die van vandaag waren post hem in teams. 
$Rows | ForEach-Object -Parallel {

    #functie om advisory naar teams te posten, moet in de loop geplaatst worden omdat die parralel draait. 
    $webhookUrl = "https://ogd.webhook.office.com/webhookb2/05a60ba5-2c9a-46ad-9190-1cd78429d0b9@afca0a52-882c-4fa8-b71d-f6db2e36058b/IncomingWebhook/9278ec4c40dc472cb80ec3c6f2d12123/72a09449-2f5e-491b-8cb4-2c2ebf993f92"
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

    # Convert the object to JSON
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
    
    # vraag de datum vandaag op. 
    $dateoftoday = (get-date -day 18 -month 6 -year 2024).ToString("dd MMMM yyyy")

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

    #write-host "comparing: $dateoftoday with  $($issue.PublishedDate)"

    if($dateoftoday -eq $issue.PublishedDate ){
        write-host "found new advisory. posting it to teams. "
        Send-TeamsNotification -Issue $issue -WebhookUrl $webhookUrl
    }

    # Send-TeamsNotification -Issue $issue -WebhookUrl $webhookUrl
}

write-host "closing webdriver."
$Driver.quit()



# Send-TeamsNotification -Issue $issue -WebhookUrl $webhookUrl