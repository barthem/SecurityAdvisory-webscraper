# Load the Selenium module
Import-Module Selenium

# Setup the URL and the driver
$url = "https://www.veeam.com/knowledge-base.html?type=security"
$webhookUrl = "https://ogd.webhook.office.com/webhookb2/e268ccf5-d2fd-4cb2-abbe-90f7fe7720ea@afca0a52-882c-4fa8-b71d-f6db2e36058b/IncomingWebhook/2c2f8ffef3ce4fb8acd36244a4d92ba4/72a09449-2f5e-491b-8cb4-2c2ebf993f92"

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
<b>Document URL:</b> <a href='$($Issue.Link)'>$($Issue.Link)</a><br>
<b>Support Products:</b> $($Issue.Product)<br>
<b>DatePublished:</b> $($Issue.DatePublished)<br>
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


#start firefox headless, becouse the pipeline agen haz no gui
$Driver = Start-SeFirefox -headless
$Driver.Navigate().GoToUrl($url)

# Find the articles by class name
$classname = "knowledge-base-listing__article"
$articles = $Driver.FindElementsByClassName($classname)



# Loop through each article and extract the required information
$articleData = $articles | foreach-object -Parallel {
    write-host "processing article..."
    $article = $_

    $name = $article.FindElementByClassName("knowledge-base-listing__article-id").Text
    $titleElement = $article.FindElementByClassName("knowledge-base-listing__article-title")
    $title = $titleElement.Text
    $link = $titleElement.GetAttribute("href")
    $description = $article.FindElementByClassName("knowledge-base-listing__article-description").Text

    # Extract date and product using regex or string manipulation
    $datePublished = if ($description -match 'Date published: (\d{4}-\d{2}-\d{2})') {
        $matches[1]
    }
    else {
        'Unknown'
    }

    $product = if ($description -match '\| Product: (.+)$') {
        $matches[1]
    }
    else {
        'Unknown'
    }

    # Create a custom object for each article
    [PSCustomObject]@{
        Name          = $name
        Title         = $title
        Link          = $link
        DatePublished = $datePublished
        Product       = $product
    }
}

# vraag de datum vandaag op. 
$dateoftoday = (get-date).ToString("yyyy-MM-dd")

#check of er vandaag iets nieuws gepost is, zo ja post het dan. 
$articleData | foreach-object {
    write-host "processing $($_.Name) $($_.Title) `npublished: $($_.DatePublished)`n"
    
    if ($_.DatePublished -eq $dateoftoday ) {

        Write-Host " article $($_.name) is from today!. posting it to teams"
        Send-TeamsNotification -Issue $_ -WebhookUrl $webhookUrl
    }
}

# Cleanup
$Driver.Quit()
# $articleData 