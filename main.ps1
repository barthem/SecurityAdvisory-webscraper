# Navigate to the URL
import-module selenium 

# $currentDate = (Get-Date).ToString("dd MMMM yyyy")


$url = "https://support.broadcom.com/web/ecx/security-advisory?segment=VC"
$Driver = Start-SeFirefox -headless
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








$DataArray = @()

# foreach ($Row in $Rows) {

$tablecontent = $Driver.FindElementByTagName("tbody")
# $tablecontent.GetAttribute("outerHTML")

$Rows = $tableContent.FindElementsByTagName("tr")



$Rows | ForEach-Object -Parallel {

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
    
    $dateoftoday = (get-date -day 8 -month 5 -year 2024).ToString("dd MMMM yyyy")

    $Cells = $_.FindElementsByTagName("td")

    # $Cells = $Rows[$iii].FindElementsByTagName("td")

    # Get the anchor tag from the first cell which contains the URL of the advisory
    $DocumentLink = $Cells[0].FindElementByTagName("a")

    # Create a custom object for each row, now including the URL from the DocumentID cell
    $issue = [PSCustomObject]@{
        DocumentID      = $Cells[0].Text
        DocumentURL     = $DocumentLink.GetAttribute("href")
        Title           = $Cells[1].Text
        SupportProducts = $Cells[2].Text
        Severity        = $Cells[3].Text
        PublishedDate   = $Cells[4].Text
        LastUpdated     = $Cells[5].Text
    }
    write-host "sending a message about $($issue.Title)"

    write-host "comparing: $dateoftoday with  $($issue.PublishedDate)"
    if($dateoftoday -eq $issue.PublishedDate ){
        write-host "############$($issue.Title)########################"
        Send-TeamsNotification -Issue $issue -WebhookUrl $webhookUrl
    }

    # Send-TeamsNotification -Issue $issue -WebhookUrl $webhookUrl
}


$Driver.quit()



# Send-TeamsNotification -Issue $issue -WebhookUrl $webhookUrl