# Define Exchange mailbox credentials
$email = "your-email@example.com"
$password = "your-password"
$securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $email, $securePassword
$DownloadFolder = "C:\Temp
# Define Exchange server URL
$server = "https://outlook.office365.com/EWS/Exchange.asmx"

# Define SOAP request
$soapRequest = @"
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:xsd="http://www.w3.org/2001/XMLSchema"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    <soap:Header>
        <t:RequestServerVersion Version="Exchange2010_SP2" />
    </soap:Header>
    <soap:Body>
        <m:FindItem Traversal="Shallow">
            <m:ItemShape>
                <t:BaseShape>IdOnly</t:BaseShape>
                <t:AdditionalProperties>
                    <t:FieldURI FieldURI="item:HasAttachments" />
                </t:AdditionalProperties>
            </m:ItemShape>
            <m:IndexedPageItemView MaxEntriesReturned="50" Offset="0" BasePoint="Beginning" />
            <m:ParentFolderIds>
                <t:DistinguishedFolderId Id="inbox" />
            </m:ParentFolderIds>
        </m:FindItem>
    </soap:Body>
</soap:Envelope>
"@

# Create a web request
$webRequest = [System.Net.HttpWebRequest]::Create($server)
$webRequest.Method = "POST"
$webRequest.ContentType = "text/xml; charset=utf-8"
$webRequest.Headers.Add("Authorization", "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($email):$($password)")))
$webRequest.Accept = "text/xml"

# Load the SOAP request into the request stream
$soapBytes = [System.Text.Encoding]::UTF8.GetBytes($soapRequest)
$webRequest.ContentLength = $soapBytes.Length
$requestStream = $webRequest.GetRequestStream()
$requestStream.Write($soapBytes, 0, $soapBytes.Length)
$requestStream.Close()

# Get the response
$response = $webRequest.GetResponse()
$responseStream = $response.GetResponseStream()
$reader = New-Object System.IO.StreamReader($responseStream)
$responseXml = $reader.ReadToEnd()
$reader.Close()
$responseStream.Close()

# Load the response XML
[xml]$responseDoc = $responseXml

# Process the response to find attachments
foreach ($item in $responseDoc.SelectNodes("//t:Message", $namespaceManager)) {
    if ($item.HasAttachments -eq "true") {
        # Get item ID and change key
        $itemId = $item.ItemId.Id
        $itemChangeKey = $item.ItemId.ChangeKey

        # Define SOAP request to get the attachment
        $getAttachmentRequest = @"
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:xsd="http://www.w3.org/2001/XMLSchema"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    <soap:Header>
        <t:RequestServerVersion Version="Exchange2010_SP2" />
    </soap:Header>
    <soap:Body>
        <m:GetItem>
            <m:ItemShape>
                <t:BaseShape>AllProperties</t:BaseShape>
            </m:ItemShape>
            <m:ItemIds>
                <t:ItemId Id="$itemId" ChangeKey="$itemChangeKey"/>
            </m:ItemIds>
        </m:GetItem>
    </soap:Body>
</soap:Envelope>
"@

        # Create a web request to get the item
        $webRequest = [System.Net.HttpWebRequest]::Create($server)
        $webRequest.Method = "POST"
        $webRequest.ContentType = "text/xml; charset=utf-8"
        $webRequest.Headers.Add("Authorization", "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($email):$($password)")))
        $webRequest.Accept = "text/xml"

        # Load the SOAP request into the request stream
        $soapBytes = [System.Text.Encoding]::UTF8.GetBytes($getAttachmentRequest)
        $webRequest.ContentLength = $soapBytes.Length
        $requestStream = $webRequest.GetRequestStream()
        $requestStream.Write($soapBytes, 0, $soapBytes.Length)
        $requestStream.Close()

        # Get the response
        $response = $webRequest.GetResponse()
        $responseStream = $response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($responseStream)
        $responseXml = $reader.ReadToEnd()
        $reader.Close()
        $responseStream.Close()

        # Load the response XML
        [xml]$responseDoc = $responseXml

        # Process the response to download attachments
        foreach ($attachment in $responseDoc.SelectNodes("//t:FileAttachment", $namespaceManager)) {
            $attachmentId = $attachment.AttachmentId.Id
            $attachmentName = $attachment.Name

            # Define SOAP request to get the attachment content
            $getAttachmentContentRequest = @"
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:xsd="http://www.w3.org/2001/XMLSchema"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    <soap:Header>
        <t:RequestServerVersion Version="Exchange2010_SP2" />
    </soap:Header>
    <soap:Body>
        <m:GetAttachment>
            <m:AttachmentIds>
                <t:AttachmentId Id="$attachmentId"/>
            </m:AttachmentIds>
        </m:GetAttachment>
    </soap:Body>
</soap:Envelope>
"@

            # Create a web request to get the attachment content
            $webRequest = [System.Net.HttpWebRequest]::Create($server)
            $webRequest.Method = "POST"
            $webRequest.ContentType = "text/xml; charset=utf-8"
            $webRequest.Headers.Add("Authorization", "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($email):$($password)")))
            $webRequest.Accept = "text/xml"

            # Load the SOAP request into the request stream
            $soapBytes = [System.Text.Encoding]::UTF8.GetBytes($getAttachmentContentRequest)
            $webRequest.ContentLength = $soapBytes.Length
            $requestStream = $webRequest.GetRequestStream()
            $requestStream.Write($soapBytes, 0, $soapBytes.Length)
            $requestStream.Close()

            # Get the response
            $response = $webRequest.GetResponse()
            $responseStream = $response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($responseStream)
            $responseXml = $reader.ReadToEnd()
            $reader.Close()
            $responseStream.Close()

            # Load the response XML
            [xml]$responseDoc = $responseXml

            # Decode and save the attachment content
            $attachmentContent = $responseDoc.SelectSingleNode("//t:Base64Binary", $namespaceManager).InnerText
            $attachmentBytes = [Convert]::FromBase64String($attachmentContent)
            $fileName = "C:\$DownloadFolder\\$attachmentName"
            [System.IO.File]::WriteAllBytes($fileName, $attachmentBytes)
            Write-Output "Downloaded $attachmentName"
        }
    }
}
