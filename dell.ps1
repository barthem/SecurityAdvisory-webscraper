Import-Module Selenium

# Setup the URL and the driver
$url = "https://www.dell.com/support/security/en-us"
# $webhookUrl = "https://ogd.webhook.office.com/webhookb2/e268ccf5-d2fd-4cb2-abbe-90f7fe7720ea@afca0a52-882c-4fa8-b71d-f6db2e36058b/IncomingWebhook/f4a9cfcb58d44b9e90d5cf900720edd1/72a09449-2f5e-491b-8cb4-2c2ebf993f92"

#only post items if they contain one of these keywords. 
$whitelist = @(
    "DRAC",
    "poweredge",
    "BIOS",
    "POWERSTORE",
    "live optics",
    "dell update manager",
    "UEFI",
    "TPM",
    "intel",
    "powervault",
    "unity"
)

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
        "summary"    = "New Veeam security advisory"
        "themeColor" = '0078D7'
        "title"      = "$($Issue.Title)"
        "text"       = @"
<b>name:</b> $($Issue.Name) $($Issue.Title)<br>
<b>advisory URL:</b> <a href='$($Issue.ArticleUrl)'>Advisory Link</a><br>
<b>type:</b> $($Issue.type)<br>
<b>CVE_ID:</b> $($Issue.CVE_ID)<br>
<b>Date Published:</b> $($Issue.Published)<br>
<b>Date updated:</b> $($Issue.Published)<br>
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


# Setup the URL and the driver
$Driver = Start-SeFirefox -Fullscreen 
$Driver.Navigate().GoToUrl($url)

#try reading the pbody max 5 times. 
$i = 0
do {
    write-host "attempting to read html tbody"
    write-host "attempt number: $i"
    try {
        $tableBody = $driver.FindElementByClassName("dds__tbody")
    }
    catch {
        write-host "failed to read tbody, trying again. attempt number $i"
    }

    # $tablecontent.GetAttribute("outerHTML")
    $i += 1

    #kill program if element fails
    if ($i -gt 5) {
        throw "coudlnt read tbody"
        exit 
    }
} until (      
    <# Condition that stops the loop if it returns true #>
    $null -ne $tableBody.displayed       
)
write-host "retrieval of html body was succesful!"

write-host "click on button to sort advisories on publisheddate - ascending"
#column with all the headers in it
$tableheadercolumn = $Driver.FindElementByClassName("dds__thead")
$individualTableHeaders = $tableheadercolumn.FindElementsByClassName("dds__th")

<#
0 - impact 
1 - titel
2 - type 
3 - CVE id
4 - published
5 - updated
#>
#click on published tab
$individualTableHeaders[4].click()
start-sleep -Seconds 5


# get action menus, one should have been triggered by the click
$tableheadercolumn = $Driver.FindElementSByClassName("dds__action-menu")


# <#
# 0 - impact 
# 1 - titel
# 2 - type 
# 3 - CVE id
# 4 - published
# 5 - updated
# #>
$buttons = $tableheadercolumn[4].FindElementsByTagName("button")

# <#buttons!
# 0 - unsorted
# 1 - ascending
# 2 - descending 
# #>
$buttons[2].click()


# $Driver.FindElementByXPath("/html/body/div[7]/div/ul/li[2]/button").click()


# # # Find the table body element
# $tableBody = $driver.FindElementByClassName("dds__tbody")

# # Get all the rows in the table body
# $rows = $tableBody.FindElementsByClassName("dds__tr")

# #loop thorugh all table elements, and parse them. 
# $advisories = $rows | ForEach-Object -ThrottleLimit 10 -Parallel {
#     $columns = $_.FindElementsByClassName("dds__td")

#     # Assuming columns are structured as:
#     # 0 - Impact, 1 - Title with Link, 2 - Type, 3 - CVE ID, 4 - Published Date, 5 - Updated Date

#     # Get the <a> tag from the Title column to extract the url and title
#     $titleLink = $columns[1].FindElementByTagName("a")

#     $entry = [PSCustomObject]@{
#         Impact     = $columns[0].Text
#         Title      = $titleLink.GetAttribute("innerText")        
#         ArticleUrl = $titleLink.GetAttribute("href")
#         Type       = $columns[2].Text
#         CVE_ID     = $columns[3].Text -replace '<.*?>', '' # Clean any HTML tags if present
#         Published  = $columns[4].Text
#         Updated    = $columns[5].Text
#     }
#     $entry
#     write-host "done parsing $($entry.Title)"
# }


# #create a regexpattern from all the products in the whitelist 
# $regexPattern = ($whitelist | ForEach-Object { [regex]::Escape($_) }) -join '|'

# #remove entries that dont contain the whitelisted keywords
# $advisories = $advisories |Where-Object {$_.Title -match $regexPattern}

# # get date of today, in the format of the security advisory. 
# $dateoftoday = (get-date).ToString("MMM dd yyyy").ToUpper()

# #check if a new advisory was posted today, if yes post to teams 
# $advisories | foreach-object {
#     write-host "processing  $($_.Title) `npublished: "

    
#     #compare published to date of today
#     if ($_.Published -eq $dateoftoday ) {
#         Write-Host " article $($_.Title) is from today!. posting it to teams"
#         # Send-TeamsNotification -Issue $_ -WebhookUrl $webhookUrl
#     }
# }

# # Close the Selenium driver
# $Driver.Quit()
