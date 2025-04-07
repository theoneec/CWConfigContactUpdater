# ==== CONFIGURATION (Override from command line if available) ====
param(
    [Parameter(Mandatory = $true)][string]$CWCompanyId,
    [Parameter(Mandatory = $true)][string]$PublicKey,
    [Parameter(Mandatory = $true)][string]$PrivateKey,
    [Parameter(Mandatory = $true)][string]$ClientId,
    [Parameter(Mandatory = $true)][string]$CompanyId,
    [Parameter(Mandatory = $true)][string]$Site,
    [string]$ApiVersion = "2025.1",  # Default to just version if user doesn't override
    [switch]$BypassBaseUrlCheck
)

# ==== VALIDATION ====
if (-not $CWCompanyId -or $CWCompanyId.Trim() -eq "")     { throw "❌ Missing CWCompanyId (use -CWCompanyId)" }
if (-not $PublicKey -or $PublicKey.Trim() -eq "")         { throw "❌ Missing PublicKey (use -PublicKey)" }
if (-not $PrivateKey -or $PrivateKey.Trim() -eq "")       { throw "❌ Missing PrivateKey (use -PrivateKey)" }
if (-not $ClientId -or $ClientId.Trim() -eq "")           { throw "❌ Missing ClientId (use -ClientId)" }
if (-not $CompanyId -or $CompanyId.Trim() -eq "")         { throw "❌ Missing CompanyId (use -CompanyId)" }
if (-not $Site -or $Site.Trim() -eq "")                   { throw "❌ Missing Site URL (use -Site)" }

# Validate known base URLs or allow custom local override
$validBaseHosts = @(
    "https://api-na.myconnectwise.net",
    "https://api-au.myconnectwise.net",
    "https://api-eu.myconnectwise.net"
)
try {
    $uriCheck = [uri]$Site
    $base = "$($uriCheck.Scheme)://$($uriCheck.Host.ToLower())"
    if (-not $BypassBaseUrlCheck -and ($validBaseHosts -notcontains $base)) {
        Write-Warning "⚠️ Site host '$base' not in recognized ConnectWise regions. Assuming local or custom environment."
    }
} catch {
    throw "❌ Invalid Site URL format: $_"
}

# Expand short ApiVersion if needed
if ($ApiVersion -notmatch "^application/vnd\\.connectwise\\.com\+json; version=") {
    $ApiVersion = "application/vnd.connectwise.com+json; version=" + $ApiVersion
}
# ==== AUTH HEADERS ====
$EncodedAuth = [Convert]::ToBase64String(
    [Text.Encoding]::ASCII.GetBytes("${CWCompanyId}+${PublicKey}:${PrivateKey}")
)
$Headers = @{
    "Authorization" = "Basic $EncodedAuth"
    "clientId"      = $ClientId
    "Accept"        = $ApiVersion
    "Content-Type"  = "application/json"
}

# ==== SETUP ====
$ConfigsUrl = "/v4_6_release/apis/3.0/company/configurations"
$siteconfigurl = $Site + $ConfigsUrl
$Condition = [uri]::EscapeDataString("company/identifier=""$CompanyId""")
Write-Host "$siteconfigurl"
$AllConfigs = @()
$Page = 1
$PageSize = 100


# ==== STEP 1: Retrieve full JSON for each config and export all fields to CSV ====
Write-Host "\n📦 Step 1: Retrieving full configuration JSONs and preparing snapshot..."
$DetailedConfigs = @()

try {
    do {
        $FinalUrl = $siteconfigurl + "?page=$Page&pageSize=$PageSize&conditions=$Condition"
        Write-Host "$FinalUrl"
        Write-Host "📦 Retrieving config page $Page..."
        $PageData = Invoke-RestMethod -Uri $FinalUrl -Headers $Headers -Method Get

        if (-not $PageData) { break }

        foreach ($config in $PageData) {
            $ConfigId = $config.id
            $ConfigDetailUrl = "$Site/v4_6_release/apis/3.0/company/configurations/$ConfigId"
            try {
                $DetailedConfig = Invoke-RestMethod -Uri $ConfigDetailUrl -Headers $Headers -Method Get
                $DetailedConfigs += $DetailedConfig

                # Save JSON for audit
                $JsonPath = "step2config$ConfigId.json"
                $DetailedConfig | ConvertTo-Json -Depth 100 | Out-File -Encoding UTF8 -FilePath $JsonPath
            } catch {
                Write-Warning "⚠️ Failed to retrieve config ID ${ConfigId}: $_"
            }
        }

        $Page++
    } while ($PageData.Count -eq $PageSize)
} catch {
    Write-Error ("❌ Error on config page ${Page}: " + $_.Exception.Message)
}

# Export all detailed configs to CSV
$Step1Csv = "step1_config_snapshot.csv"
$DetailedConfigs | Select-Object * | Export-Csv -Path $Step1Csv -NoTypeInformation -Encoding UTF8
Write-Host "📄 Exported Step 1 config snapshot with full JSON data to: $Step1Csv"

# ==== STEP 2: Extract contact fields and append to step1 CSV ====
Write-Host "\n🔁 Step 2: Extracting contact fields from full config JSONs..."
$Step2Csv = "step2_config_snapshot.csv"

if (Test-Path $Step1Csv) {
    $Step1Data = Import-Csv $Step1Csv
    $Step2Output = foreach ($row in $Step1Data) {
        $id = $row.id
        $jsonPath = "step2config$id.json"
        if (Test-Path $jsonPath) {
            try {
                $jsonData = Get-Content -Path $jsonPath -Raw | ConvertFrom-Json
                $contactName = $jsonData.contact.name
                $contactId = $jsonData.contact.id
                $contactHref = $jsonData.contact._info.contact_href

                $row | Add-Member -NotePropertyName "contact.name" -NotePropertyValue $contactName -Force
                $row | Add-Member -NotePropertyName "contact.id" -NotePropertyValue $contactId -Force
                $row | Add-Member -NotePropertyName "contact._info" -NotePropertyValue $contactHref -Force
            } catch {
                Write-Warning "⚠️ Failed to parse $jsonPath"
            }
        } else {
            Write-Warning "⚠️ JSON file not found for config ID $id"
        }
        $row
    }

    $Step2Output | Export-Csv -Path $Step2Csv -NoTypeInformation -Encoding UTF8
    Write-Host "📄 Exported updated config snapshot to: $Step2Csv"
} else {
    Write-Error "❌ Step 1 CSV not found. Cannot continue step 2."
}



# ==== STEP 3: Extract key config and contact columns to simplified CSV ====
Write-Host "\n🧹 Step 3: Simplifying configuration data for review..."
$Step3Csv = "step3_clean_configs.csv"

if (Test-Path $Step2Csv) {
    $Step2Data = Import-Csv $Step2Csv
    $Step3Data = $Step2Data | Select-Object id, name, lastLoginName, activeFlag, 'contact.name', 'contact.id', 'contact._info'
    $Step3Data | Export-Csv -Path $Step3Csv -NoTypeInformation -Encoding UTF8
    Write-Host "📄 Exported simplified config data to: $Step3Csv"
} else {
    Write-Error "❌ Step 2 CSV not found. Cannot continue step 3."
}

# ==== STEP 4: Refresh all company contacts and export to CSV ====
Write-Host "\n🔁 Step 4: Retrieving all contacts for company ID $CompanyId..."
$AllContacts = @()
$Page = 1
$PageSize = 100
$BaseUrl = "$Site/v4_6_release/apis/3.0/company/contacts"
$Condition = [uri]::EscapeDataString("company/identifier='$CompanyId'")

try {
    do {
        $Url = $BaseUrl+"?conditions=$Condition&pageSize=$PageSize&page=$Page"
        Write-Host "📄 Fetching contacts page $Page..."
        $Response = Invoke-RestMethod -Uri $Url -Headers $Headers -Method Get

        if ($Response) {
            $AllContacts += $Response
        }

        $Page++
    } while ($Response.Count -eq $PageSize)
} catch {
    Write-Error "❌ Failed to retrieve contacts: $($_.Exception.Message)"
    $AllContacts = @()
}

# ==== STEP 4A: Extract relevant fields and save contacts ====
$ProcessedContacts = foreach ($c in $AllContacts) {
    $defaultEmail = $null
    if ($c.communicationItems) {
        $defaultEmail = $c.communicationItems |
            Where-Object { $_.communicationType -eq "Email" -and $_.defaultFlag } |
            Select-Object -ExpandProperty value -First 1
    }

    [pscustomobject]@{
        id         = $c.id
        firstName  = $c.firstName
        lastName   = $c.lastName
        title      = $c.title
        email      = $defaultEmail
        phone      = $c.defaultPhoneNbr
    }
}

$Step4aCsv = "step4a_company_contacts.csv"
$ProcessedContacts | Export-Csv -Path $Step4aCsv -NoTypeInformation -Encoding UTF8
Write-Host "\n📇 Exported full company contacts to: $Step4aCsv"

# ==== STEP 4B: Guess contact from lastLoginName and compare ====
Write-Host "\n🧠 Step 4B: Analyzing contact match guesses..."
$Step4bCsv = "step4b_contact_guess.csv"

$KnownContacts = $ProcessedContacts | ForEach-Object {
    ($_.firstName + " " + $_.lastName).Trim().ToLower()
}

$Step3Data = Import-Csv $Step3Csv

$ContactMatchResults = foreach ($row in $Step3Data) {
    $lastLogin = $row.lastLoginName
    $guessedFirst = $null
    $guessedLast = $null
    $guessedFull = $null
    $match = $false
    $exists = $false

    if ($lastLogin -and $lastLogin -match "\\") {
        $username = $lastLogin -split "\\" | Select-Object -Last 1

        $split = [regex]::Split($username, "(?=[A-Z])") | Where-Object { $_ -ne "" }

        if ($split.Count -gt 1) {
            $guessedFirst = $split[0]
            $guessedLast  = ($split[1..($split.Count - 1)] -join " ")
        } elseif ($split.Count -eq 1) {
            $guessedFirst = $split[0]
        }

        $guessedFull = "$guessedFirst $guessedLast".Trim()
        $guessedFullLower = $guessedFull.ToLower()

        if ($KnownContacts -contains $guessedFullLower) {
            $exists = $true
        }

        $actualName = $row.'contact.name'
        if ($actualName -and $guessedFull -and ($actualName.ToLower() -eq $guessedFullLower)) {
            $match = $true
        }
    }

    [pscustomobject]@{
        id             = $row.id
        name           = $row.name
        lastLoginName  = $row.lastLoginName
        activeFlag     = $row.activeFlag
        'contact.name' = $row.'contact.name'
        guessedFirst   = $guessedFirst
        guessedLast    = $guessedLast
        guessedFull    = $guessedFull
        matched        = $match
        exists         = $exists
    }
}

$ContactMatchResults | Export-Csv -Path $Step4bCsv -NoTypeInformation -Encoding UTF8
Write-Host "📄 Exported contact match comparison to: $Step4bCsv"

# ==== STEP 5: Push updates where matched = FALSE, exists = TRUE, and activeFlag is TRUE ====
Write-Host "\n🚀 Step 5: Preparing configuration updates for unmatched-but-existing contacts..."
$Step4bCsv = "step4b_contact_guess.csv"

if (-not (Test-Path $Step4bCsv) -or -not (Test-Path $Step4aCsv)) {
    Write-Error "❌ Required Step 4 CSVs not found. Ensure step 4 ran successfully."
    exit 1
}

$GuessData = Import-Csv $Step4bCsv
$ContactData = Import-Csv $Step4aCsv

foreach ($row in $GuessData | Where-Object { $_.matched -eq 'FALSE' -and $_.exists -eq 'TRUE' -and $_.activeFlag -eq 'TRUE' }) {
    $ConfigId = $row.id
    $GuessedName = $row.guessedFull.Trim()

    $ContactMatch = $ContactData | Where-Object {
        ("$($_.firstName) $($_.lastName)").Trim().ToLower() -eq $GuessedName.ToLower()
    } | Select-Object -First 1

    if (-not $ContactMatch) {
        Write-Warning "⚠️ No contact found in Step 4A for guessed name '$GuessedName'"
        continue
    }

    $ContactId   = [int]$ContactMatch.id
    $ContactName = "$($ContactMatch.firstName) $($ContactMatch.lastName)"
    $ContactHref = "$Site/v4_6_release/apis/3.0/company/contacts/$ContactId"

    $ConfigUrl = "$Site/v4_6_release/apis/3.0/company/configurations/$ConfigId"
    try {
        $ConfigRaw = Invoke-RestMethod -Method Get -Uri $ConfigUrl -Headers $Headers

        if ($ConfigRaw.activeFlag -eq $false) {
            Write-Host "⚠️ Skipping config ID $ConfigId (inactive)"
            continue
        }

        $JsonPath = "config-${ConfigId}.json"
        $ConfigRaw | ConvertTo-Json -Depth 100 | Out-File -Encoding UTF8 -FilePath $JsonPath

        $ConfigRaw.contact = [PSCustomObject]@{
            id    = $ContactId
            name  = $ContactName
            _info = @{ contact_href = $ContactHref }
        }

        $ModifiedJson = $ConfigRaw | ConvertTo-Json -Depth 100
        $FinalJsonFile = "config-${ConfigId}-to-push.json"
        $ModifiedJson | Out-File -Encoding UTF8 -FilePath $FinalJsonFile

        Write-Host "📤 Updating configuration ID $ConfigId with contact '$ContactName'..."
        $Response = Invoke-RestMethod -Uri $ConfigUrl -Method Put -Headers $Headers -Body (Get-Content -Raw -Path $FinalJsonFile)
        Write-Host "✅ Configuration ID $ConfigId successfully updated."
    } catch {
        Write-Error ("❌ Failed to update config ID ${ConfigId}: " + $_.Exception.Message)
    }
}

# ==== STEP 6: CLEANUP TEMP FILES ====
Write-Host "\n🧹 Step 6: Cleaning up temporary files..."
$CleanupPatterns = @("step1*.csv", "step2*.csv", "step3*.csv", "step4*.csv","step2*.json", "config-*.json", "config-*-to-push.json")

foreach ($pattern in $CleanupPatterns) {
    $FilesToDelete = Get-ChildItem -Path . -Filter $pattern -File -ErrorAction SilentlyContinue
    foreach ($file in $FilesToDelete) {
        try {
            Remove-Item -Path $file.FullName -Force
            Write-Host "🗑️ Deleted: $($file.Name)"
        } catch {
            Write-Warning "⚠️ Could not delete $($file.Name): $_"
        }
    }
}

Write-Host "✅ Cleanup complete."