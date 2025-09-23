Exchange Online Mailbox and Folder Detailed Report Script
Version: 2.4
Description: Generates comprehensive mailbox and folder statistics with retention information
and intelligent storage analysis in a tabbed interface
Accepts UPN, Email Address, or Mail Nickname/SAM Account Name as input
param(
[Parameter(Mandatory=$true)]
[string]$Identity,

text

[Parameter(Mandatory=$false)]
[string]$OutputPath = ".\MailboxReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html",

[Parameter(Mandatory=$false)]
[switch]$ExportToCSV
)

Function to connect to Exchange Online
function Connect-ExchangeOnlineIfNeeded {
try {
Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null
Write-Host "Already connected to Exchange Online" -ForegroundColor Green
}
catch {
Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
try {
Connect-ExchangeOnline -ShowBanner:$false
Write-Host "Successfully connected to Exchange Online" -ForegroundColor Green
}
catch {
Write-Host "Failed to connect to Exchange Online. Please ensure you have the Exchange Online module installed." -ForegroundColor Red
Write-Host "Install-Module -Name ExchangeOnlineManagement" -ForegroundColor Yellow
exit 1
}
}
}

Function to resolve user identity
function Resolve-UserIdentity {
param([string]$Identity)

text

Write-Host "Resolving user identity: $Identity" -ForegroundColor Cyan

# Try to get mailbox with the provided identity
$mailbox = $null

try {
    # First try direct lookup
    $mailbox = Get-EXOMailbox -Identity $Identity -ErrorAction SilentlyContinue
    
    if (-not $mailbox) {
        # Try as Alias/MailNickname
        $mailbox = Get-EXOMailbox -Filter "Alias -eq '$Identity'" -ErrorAction SilentlyContinue
    }
    
    if (-not $mailbox) {
        # Try as SAMAccountName
        $mailbox = Get-EXOMailbox -Filter "SamAccountName -eq '$Identity'" -ErrorAction SilentlyContinue
    }
    
    if (-not $mailbox) {
        # Try as EmailAddress
        $mailbox = Get-EXOMailbox -Filter "EmailAddresses -eq 'smtp:$Identity'" -ErrorAction SilentlyContinue
    }
    
    if (-not $mailbox) {
        # Try as PrimarySmtpAddress
        $mailbox = Get-EXOMailbox -Filter "PrimarySmtpAddress -eq '$Identity'" -ErrorAction SilentlyContinue
    }
    
    if ($mailbox) {
        Write-Host "User found: $($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))" -ForegroundColor Green
        return $mailbox.UserPrincipalName
    }
    else {
        throw "Unable to find mailbox with identity: $Identity"
    }
}
catch {
    Write-Host "Error resolving user identity: $_" -ForegroundColor Red
    exit 1
}
}

Function to format bytes to readable format
function Format-Bytes {
param($Size)

text

if ($Size -eq $null -or $Size -eq 0) {
    return "0 B"
}

$bytes = 0

# Handle if Size is already a string with bytes value
if ($Size -is [string]) {
    if ($Size -match '([\d.]+)\s*([KMGT]B).*KATEX_INLINE_OPEN([\d,]+)\s*bytesKATEX_INLINE_CLOSE') {
        $bytes = [long]($matches[3] -replace ',', '')
    }
    elseif ($Size -match '([\d,]+)\s*bytes') {
        $bytes = [long]($matches[1] -replace ',', '')
    }
    else {
        return $Size
    }
}
elseif ($Size.GetType().Name -eq 'Unlimited`1') {
    return "Unlimited"
}
elseif ($Size.GetType().Name -eq 'ByteQuantifiedSize') {
    $bytes = $Size.ToBytes()
}
else {
    $bytes = [long]$Size
}

if ($bytes -ge 1TB) {
    return "{0:N2} TB" -f ($bytes / 1TB)
}
elseif ($bytes -ge 1GB) {
    return "{0:N2} GB" -f ($bytes / 1GB)
}
elseif ($bytes -ge 1MB) {
    return "{0:N2} MB" -f ($bytes / 1MB)
}
elseif ($bytes -ge 1KB) {
    return "{0:N2} KB" -f ($bytes / 1KB)
}
else {
    return "{0} B" -f $bytes
}
}

Function to get bytes from size object or string
function Get-BytesFromSize {
param($Size)

text

if ($Size -eq $null) {
    return 0
}

# If it's already a number, return it
if ($Size -is [int] -or $Size -is [long] -or $Size -is [double]) {
    return [long]$Size
}

# If it's a string, parse it
if ($Size -is [string]) {
    if ($Size -match '([\d.]+)\s*([KMGT]B).*KATEX_INLINE_OPEN([\d,]+)\s*bytesKATEX_INLINE_CLOSE') {
        return [long]($matches[3] -replace ',', '')
    }
    elseif ($Size -match '([\d,]+)\s*bytes') {
        return [long]($matches[1] -replace ',', '')
    }
    elseif ($Size -match '^([\d,]+)$') {
        return [long]($Size -replace ',', '')
    }
    else {
        return 0
    }
}

# If it has a ToBytes method, use it
if ($Size.PSObject.Methods.Name -contains 'ToBytes') {
    return $Size.ToBytes()
}

# If it has a Value property that has ToBytes
if ($Size.Value -and $Size.Value.PSObject.Methods.Name -contains 'ToBytes') {
    return $Size.Value.ToBytes()
}

return 0
}

Function to parse storage limit string
function Parse-StorageLimit {
param($Limit)

text

if ($Limit -eq $null) {
    return "Not Set"
}

if ($Limit.ToString() -eq "Unlimited") {
    return "Unlimited"
}

return Format-Bytes -Size $Limit
}

Function to analyze mailbox for potential issues
function Get-MailboxHealthAnalysis {
param(
$MailboxInfo,
$FolderDetails,
$MailboxConfig
)

text

$analysis = @()
$warnings = @()
$recommendations = @()

# Check if mailbox is approaching quota limits
$primarySizeBytes = Get-BytesFromSize -Size $MailboxInfo.'Primary Mailbox Size'
$prohibitSendBytes = Get-BytesFromSize -Size $MailboxInfo.'Prohibit Send Quota'
$warningBytes = Get-BytesFromSize -Size $MailboxInfo.'Issue Warning Quota'

if ($prohibitSendBytes -gt 0 -and $primarySizeBytes -ge ($prohibitSendBytes * 0.9)) {
    $analysis += @{
        Type = "Warning"
        Message = "Mailbox is approaching prohibit send quota (90% full)"
        Severity = "High"
    }
    $warnings += "Mailbox is at " + ([math]::Round(($primarySizeBytes/$prohibitSendBytes)*100, 2)) + "% of prohibit send quota"
    $recommendations += "Consider archiving old items or requesting quota increase"
}
elseif ($warningBytes -gt 0 -and $primarySizeBytes -ge ($warningBytes * 0.8)) {
    $analysis += @{
        Type = "Warning"
        Message = "Mailbox is approaching warning quota (80% full)"
        Severity = "Medium"
    }
    $warnings += "Mailbox is at " + ([math]::Round(($primarySizeBytes/$warningBytes)*100, 2)) + "% of warning quota"
    $recommendations += "Consider cleaning up unnecessary items or enabling archive"
}

# Check for large folders that might need attention
$largeFolders = $FolderDetails | Where-Object { $_.'Folder Size (Bytes)' -gt 1GB } | Sort-Object 'Folder Size (Bytes)' -Descending
if ($largeFolders.Count -gt 0) {
    $analysis += @{
        Type = "Information"
        Message = "Found $($largeFolders.Count) folders larger than 1GB"
        Severity = "Medium"
    }
    $largeFolders | ForEach-Object {
        $warnings += "Large folder: $($_.'Folder Name') - $($_.'Folder Size')"
        $recommendations += "Review contents of '$($_.'Folder Name')' folder for archiving potential"
    }
}

# Check for folders with many items
$manyItemsFolders = $FolderDetails | Where-Object { $_.'Item Count' -replace ',','' -as [int] -gt 10000 } | Sort-Object 'Item Count' -Descending
if ($manyItemsFolders.Count -gt 0) {
    $analysis += @{
        Type = "Information"
        Message = "Found $($manyItemsFolders.Count) folders with more than 10,000 items"
        Severity = "Low"
    }
    $manyItemsFolders | ForEach-Object {
        $warnings += "Folder with many items: $($_.'Folder Name') - $($_.'Item Count') items"
        $recommendations += "Consider splitting '$($_.'Folder Name')' folder or implementing retention policies"
    }
}

# Check retention policy status
if ($MailboxInfo.'Retention Policy' -eq "None") {
    $analysis += @{
        Type = "Warning"
        Message = "No retention policy applied to mailbox"
        Severity = "Medium"
    }
    $warnings += "No retention policy is configured"
    $recommendations += "Consider applying appropriate retention policy to manage mailbox growth"
}

# Check archive status
if ($MailboxInfo.'Archive Status' -ne "Active" -and $primarySizeBytes -gt 5GB) {
    $analysis += @{
        Type = "Recommendation"
        Message = "Mailbox larger than 5GB but archive not enabled"
        Severity = "Medium"
    }
    $warnings += "Mailbox exceeds 5GB without archive"
    $recommendations += "Consider enabling online archive for better storage management"
}

# Check litigation hold status for large mailboxes
if ($MailboxInfo.'Litigation Hold' -eq "Enabled" -and $primarySizeBytes -gt 10GB) {
    $analysis += @{
        Type = "Information"
        Message = "Litigation hold enabled on large mailbox (>10GB)"
        Severity = "Low"
    }
    $warnings += "Litigation hold may prevent automatic cleanup of items"
    $recommendations += "Monitor mailbox growth and consider alternative preservation methods if needed"
}

# Check retention hold status
if ($MailboxInfo.'Retention Hold' -eq "Enabled") {
    $analysis += @{
        Type = "Information"
        Message = "Retention hold is enabled"
        Severity = "Low"
    }
    $warnings += "Retention policies are temporarily suspended"
    $recommendations += "Verify if retention hold is still needed and schedule end date if appropriate"
}

# Check for folders without retention tags
$foldersWithoutRetention = $FolderDetails | Where-Object { $_.'Retention Tag' -eq "None" -and $_.'Folder Type' -notin @('Root', 'RecoverableItems', 'Calendar', 'Contacts') }
if ($foldersWithoutRetention.Count -gt 5) {
    $analysis += @{
        Type = "Information"
        Message = "$($foldersWithoutRetention.Count) folders without retention tags"
        Severity = "Low"
    }
    $warnings += "Multiple folders lack retention policies"
    $recommendations += "Consider applying default retention policy to ensure proper lifecycle management"
}

# Check for very old items
$currentDate = Get-Date
$oldItemsFolders = $FolderDetails | Where-Object { 
    $_.'Oldest Item' -ne "N/A" -and 
    [datetime]::ParseExact($_.'Oldest Item', 'yyyy-MM-dd HH:mm:ss', $null) -lt $currentDate.AddYears(-3) 
}

if ($oldItemsFolders.Count -gt 0) {
    $analysis += @{
        Type = "Information"
        Message = "Found $($oldItemsFolders.Count) folders with items older than 3 years"
        Severity = "Low"
    }
    $oldItemsFolders | ForEach-Object {
        $oldestDate = [datetime]::ParseExact($_.'Oldest Item', 'yyyy-MM-dd HH:mm:ss', $null)
        $yearsOld = [math]::Round(($currentDate - $oldestDate).TotalDays / 365, 1)
        $warnings += "Old items in '$($_.'Folder Name')' - oldest item is $yearsOld years old"
        $recommendations += "Consider archiving items older than 3 years in '$($_.'Folder Name')' folder"
    }
}

# Calculate storage distribution
$totalSize = ($FolderDetails | Measure-Object -Property 'Folder Size (Bytes)' -Sum).Sum
$largeFoldersPercentage = ($largeFolders | Measure-Object -Property 'Folder Size (Bytes)' -Sum).Sum / $totalSize * 100

if ($largeFoldersPercentage -gt 50) {
    $analysis += @{
        Type = "Information"
        Message = "Top folders consume $([math]::Round($largeFoldersPercentage, 1))% of total mailbox space"
        Severity = "Low"
    }
}

return @{
    Analysis = $analysis
    Warnings = $warnings
    Recommendations = $recommendations
}
}

Function to get retention policy information
function Get-RetentionPolicyInfo {
param($MailboxInfo)

text

# Get retention policy details
$retentionPolicyInfo = @()

if ($MailboxInfo.'Retention Policy' -ne "None") {
    try {
        $policy = Get-RetentionPolicy -Identity $MailboxInfo.'Retention Policy' -ErrorAction SilentlyContinue
        if ($policy) {
            $retentionPolicyInfo += "Policy Name: $($policy.Name)"
            $retentionPolicyInfo += "Description: $($policy.Description)"
            $retentionPolicyInfo += "Is Default: $($policy.IsDefault)"
            
            # Get retention policy tags
            $tags = Get-RetentionPolicyTag -ErrorAction SilentlyContinue | Where-Object { $_.RetentionPolicy -eq $policy.Name }
            if ($tags) {
                $retentionPolicyInfo += "`nAssociated Tags:"
                foreach ($tag in $tags) {
                    $retentionPolicyInfo += "  - $($tag.Name) ($($tag.RetentionAction), $($tag.AgeLimitForRetention) days)"
                }
            }
        }
    }
    catch {
        $retentionPolicyInfo += "Could not retrieve detailed retention policy information"
    }
} else {
    $retentionPolicyInfo += "No retention policy applied to this mailbox"
}

return $retentionPolicyInfo -join "`n"
}

Main script execution
try {
# Connect to Exchange Online if needed
Connect-ExchangeOnlineIfNeeded

text

# Resolve user identity
$userPrincipalName = Resolve-UserIdentity -Identity $Identity

Write-Host "`nGathering mailbox information for: $userPrincipalName" -ForegroundColor Cyan

# Get mailbox information with specific properties
$mailbox = Get-EXOMailbox -Identity $userPrincipalName -PropertySets All

# Get mailbox statistics
Write-Host "Getting mailbox statistics..." -ForegroundColor Yellow
$mailboxStats = Get-EXOMailboxStatistics -Identity $userPrincipalName -IncludeSoftDeletedRecipients

# Get mailbox configuration for additional limits
Write-Host "Getting mailbox configuration..." -ForegroundColor Yellow
$mailboxConfig = Get-Mailbox -Identity $userPrincipalName

# Get archive information if archive is enabled
$archiveStats = $null
if ($mailbox.ArchiveStatus -eq "Active") {
    Write-Host "Getting archive statistics..." -ForegroundColor Yellow
    try {
        $archiveStats = Get-EXOMailboxStatistics -Identity $userPrincipalName -Archive -ErrorAction Stop
    }
    catch {
        Write-Host "Warning: Could not retrieve archive statistics. Error: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Get retention hold information
$retentionHold = $mailboxConfig | Select-Object RetentionHoldEnabled, StartDateForRetentionHold, EndDateForRetentionHold

# Get folder statistics
Write-Host "Gathering folder statistics..." -ForegroundColor Yellow
$folders = Get-EXOMailboxFolderStatistics -Identity $userPrincipalName -IncludeOldestAndNewestItems

# Get archive folder statistics if archive is enabled
$archiveFolders = @()
if ($mailbox.ArchiveStatus -eq "Active") {
    try {
        Write-Host "Gathering archive folder statistics..." -ForegroundColor Yellow
        $archiveFolders = Get-EXOMailboxFolderStatistics -Identity $userPrincipalName -Archive -IncludeOldestAndNewestItems -ErrorAction Stop
        Write-Host "Found $($archiveFolders.Count) archive folders" -ForegroundColor Green
    }
    catch {
        Write-Host "Warning: Could not retrieve archive folder statistics. Error: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Get retention tags
Write-Host "Getting retention policy tags..." -ForegroundColor Yellow
$retentionTags = @{}
$retentionTagsByName = @{}
try {
    Get-RetentionPolicyTag -ErrorAction SilentlyContinue | ForEach-Object {
        $retentionTags[$_.Guid.ToString()] = $_
        $retentionTagsByName[$_.Name] = $_
        # Also add without spaces for matching
        $nameNoSpaces = $_.Name -replace '\s', ''
        $retentionTagsByName[$nameNoSpaces] = $_
    }
    Write-Host "Found $($retentionTags.Count) retention policy tags" -ForegroundColor Green
}
catch {
    Write-Host "Warning: Could not retrieve retention policy tags. Some retention information may be missing." -ForegroundColor Yellow
}

# Calculate sizes
$primarySize = if ($mailboxStats.TotalItemSize) { 
    if ($mailboxStats.TotalItemSize.Value) {
        Format-Bytes -Size $mailboxStats.TotalItemSize.Value
    } else {
        Format-Bytes -Size $mailboxStats.TotalItemSize
    }
} else { "0 B" }

$primaryDumpsterSize = if ($mailboxStats.TotalDeletedItemSize) { 
    if ($mailboxStats.TotalDeletedItemSize.Value) {
        Format-Bytes -Size $mailboxStats.TotalDeletedItemSize.Value
    } else {
        Format-Bytes -Size $mailboxStats.TotalDeletedItemSize
    }
} else { "0 B" }

$archiveSize = if ($archiveStats -and $archiveStats.TotalItemSize) { 
    if ($archiveStats.TotalItemSize.Value) {
        Format-Bytes -Size $archiveStats.TotalItemSize.Value
    } else {
        Format-Bytes -Size $archiveStats.TotalItemSize
    }
} else { "N/A" }

$archiveDumpsterSize = if ($archiveStats -and $archiveStats.TotalDeletedItemSize) { 
    if ($archiveStats.TotalDeletedItemSize.Value) {
        Format-Bytes -Size $archiveStats.TotalDeletedItemSize.Value
    } else {
        Format-Bytes -Size $archiveStats.TotalDeletedItemSize
    }
} else { "N/A" }

# Prepare mailbox information object
$mailboxInfo = [PSCustomObject]@{
    'Name (Alias)' = "$($mailbox.Name) ($($mailbox.Alias))"
    'Display Name' = $mailbox.DisplayName
    'Primary SMTP Address' = $mailbox.PrimarySmtpAddress
    'User Principal Name' = $mailbox.UserPrincipalName
    'SAM Account Name' = $mailbox.SamAccountName
    'Retention Policy' = if ($mailbox.RetentionPolicy) { $mailbox.RetentionPolicy } else { "None" }
    'Archive Status' = $mailbox.ArchiveStatus
    'Archive Name' = if ($mailbox.ArchiveName) { $mailbox.ArchiveName -join ", " } else { "N/A" }
    'Litigation Hold' = if ($mailbox.LitigationHoldEnabled) { "Enabled" } else { "Disabled" }
    'Litigation Hold Date' = if ($mailbox.LitigationHoldDate) { $mailbox.LitigationHoldDate.ToString("yyyy-MM-dd") } else { "N/A" }
    'Litigation Hold Owner' = if ($mailbox.LitigationHoldOwner) { $mailbox.LitigationHoldOwner } else { "N/A" }
    'Retention Hold' = if ($retentionHold.RetentionHoldEnabled) { "Enabled" } else { "Disabled" }
    'Retention Hold Start Date' = if ($retentionHold.StartDateForRetentionHold) { $retentionHold.StartDateForRetentionHold.ToString("yyyy-MM-dd") } else { "N/A" }
    'Retention Hold End Date' = if ($retentionHold.EndDateForRetentionHold) { $retentionHold.EndDateForRetentionHold.ToString("yyyy-MM-dd") } else { "N/A" }
    'Single Item Recovery' = if ($mailbox.SingleItemRecoveryEnabled) { "Enabled" } else { "Disabled" }
    'Primary Mailbox Size' = $primarySize
    'Primary Item Count' = "{0:N0}" -f $mailboxStats.ItemCount
    'Primary Dumpster Size' = $primaryDumpsterSize
    'Primary Dumpster Item Count' = "{0:N0}" -f $mailboxStats.DeletedItemCount
    'Archive Mailbox Size' = $archiveSize
    'Archive Item Count' = if ($archiveStats) { "{0:N0}" -f $archiveStats.ItemCount } else { "N/A" }
    'Archive Dumpster Size' = $archiveDumpsterSize
    'Archive Dumpster Item Count' = if ($archiveStats) { "{0:N0}" -f $archiveStats.DeletedItemCount } else { "N/A" }
    'Issue Warning Quota' = Parse-StorageLimit -Limit $mailbox.IssueWarningQuota
    'Prohibit Send Quota' = Parse-StorageLimit -Limit $mailbox.ProhibitSendQuota
    'Prohibit Send/Receive Quota' = Parse-StorageLimit -Limit $mailbox.ProhibitSendReceiveQuota
    'Archive Warning Quota' = Parse-StorageLimit -Limit $mailbox.ArchiveWarningQuota
    'Archive Quota' = Parse-StorageLimit -Limit $mailbox.ArchiveQuota
    'Max Send Size' = Parse-StorageLimit -Limit $mailboxConfig.MaxSendSize
    'Max Receive Size' = Parse-StorageLimit -Limit $mailboxConfig.MaxReceiveSize
    'Recoverable Items Quota' = Parse-StorageLimit -Limit $mailbox.RecoverableItemsQuota
    'Recoverable Items Warning Quota' = Parse-StorageLimit -Limit $mailbox.RecoverableItemsWarningQuota
}

# Process folder information
Write-Host "Processing folder information..." -ForegroundColor Yellow
$folderDetails = @()

foreach ($folder in $folders) {
    $retentionTag = "None"
    $retentionAction = "N/A"
    $ageLimitForRetention = "N/A"
    
    if ($folder.DeletePolicy) {
        $policyString = $folder.DeletePolicy.ToString()
        
        # Try to match by GUID first
        if ($retentionTags.ContainsKey($policyString)) {
            $tag = $retentionTags[$policyString]
            $retentionTag = $tag.Name
            $retentionAction = $tag.RetentionAction
            $ageLimitForRetention = if ($tag.AgeLimitForRetention) { $tag.AgeLimitForRetention.ToString() } else { "N/A" }
        }
        # Try to match by name if it contains parentheses (format: GUID (Name))
        elseif ($policyString -match 'KATEX_INLINE_OPEN(.*?)KATEX_INLINE_CLOSE') {
            $tagName = $matches[1]
            if ($retentionTagsByName.ContainsKey($tagName)) {
                $tag = $retentionTagsByName[$tagName]
                $retentionTag = $tag.Name
                $retentionAction = $tag.RetentionAction
                $ageLimitForRetention = if ($tag.AgeLimitForRetention) { $tag.AgeLimitForRetention.ToString() } else { "N/A" }
            }
            else {
                # If we can't find it in our tags, at least show the name
                $retentionTag = $tagName
            }
        }
        # If it's just a name without GUID
        elseif ($retentionTagsByName.ContainsKey($policyString)) {
            $tag = $retentionTagsByName[$policyString]
            $retentionTag = $tag.Name
            $retentionAction = $tag.RetentionAction
            $ageLimitForRetention = if ($tag.AgeLimitForRetention) { $tag.AgeLimitForRetention.ToString() } else { "N/A" }
        }
        else {
            # Show whatever we have
            $retentionTag = $policyString
        }
    }
    
    $folderSize = if ($folder.FolderSize) { 
        Format-Bytes -Size $folder.FolderSize 
    } else { "0 B" }
    
    $folderSizeBytes = Get-BytesFromSize -Size $folder.FolderSize
    
    $folderDetails += [PSCustomObject]@{
        'Location' = 'Primary Mailbox'
        'Folder Name' = $folder.Name
        'Folder Path' = $folder.FolderPath
        'Folder Type' = $folder.FolderType
        'Folder Size' = $folderSize
        'Folder Size (Bytes)' = $folderSizeBytes
        'Item Count' = "{0:N0}" -f $folder.ItemsInFolder
        'Subfolder Count' = "{0:N0}" -f $folder.SubfolderCount
        'Deleted Items' = "{0:N0}" -f $folder.DeletedItemsInFolder
        'Retention Tag' = $retentionTag
        'Retention Action' = $retentionAction
        'Age Limit' = $ageLimitForRetention
        'Oldest Item' = if ($folder.OldestItemReceivedDate) { 
            $folder.OldestItemReceivedDate.ToString("yyyy-MM-dd HH:mm:ss") 
        } else { "N/A" }
        'Newest Item' = if ($folder.NewestItemReceivedDate) { 
            $folder.NewestItemReceivedDate.ToString("yyyy-MM-dd HH:mm:ss") 
        } else { "N/A" }
    }
}

# Process archive folder information if available
foreach ($folder in $archiveFolders) {
    $retentionTag = "None"
    $retentionAction = "N/A"
    $ageLimitForRetention = "N/A"
    
    if ($folder.DeletePolicy) {
        $policyString = $folder.DeletePolicy.ToString()
        
        # Try to match by GUID first
        if ($retentionTags.ContainsKey($policyString)) {
            $tag = $retentionTags[$policyString]
            $retentionTag = $tag.Name
            $retentionAction = $tag.RetentionAction
            $ageLimitForRetention = if ($tag.AgeLimitForRetention) { $tag.AgeLimitForRetention.ToString() } else { "N/A" }
        }
        # Try to match by name if it contains parentheses (format: GUID (Name))
        elseif ($policyString -match 'KATEX_INLINE_OPEN(.*?)KATEX_INLINE_CLOSE') {
            $tagName = $matches[1]
            if ($retentionTagsByName.ContainsKey($tagName)) {
                $tag = $retentionTagsByName[$tagName]
                $retentionTag = $tag.Name
                $retentionAction = $tag.RetentionAction
                $ageLimitForRetention = if ($tag.AgeLimitForRetention) { $tag.AgeLimitForRetention.ToString() } else { "N/A" }
            }
            else {
                # If we can't find it in our tags, at least show the name
                $retentionTag = $tagName
            }
        }
        # If it's just a name without GUID
        elseif ($retentionTagsByName.ContainsKey($policyString)) {
            $tag = $retentionTagsByName[$policyString]
            $retentionTag = $tag.Name
            $retentionAction = $tag.RetentionAction
            $ageLimitForRetention = if ($tag.AgeLimitForRetention) { $tag.AgeLimitForRetention.ToString() } else { "N/A" }
        }
        else {
            # Show whatever we have
            $retentionTag = $policyString
        }
    }
    
    $folderSize = if ($folder.FolderSize) { 
        Format-Bytes -Size $folder.FolderSize 
    } else { "0 B" }
    
    $folderSizeBytes = Get-BytesFromSize -Size $folder.FolderSize
    
    $folderDetails += [PSCustomObject]@{
        'Location' = 'Archive Mailbox'
        'Folder Name' = $folder.Name
        'Folder Path' = $folder.FolderPath
        'Folder Type' = $folder.FolderType
        'Folder Size' = $folderSize
        'Folder Size (Bytes)' = $folderSizeBytes
        'Item Count' = "{0:N0}" -f $folder.ItemsInFolder
        'Subfolder Count' = "{0:N0}" -f $folder.SubfolderCount
        'Deleted Items' = "{0:N0}" -f $folder.DeletedItemsInFolder
        'Retention Tag' = $retentionTag
        'Retention Action' = $retentionAction
        'Age Limit' = $ageLimitForRetention
        'Oldest Item' = if ($folder.OldestItemReceivedDate) { 
            $folder.OldestItemReceivedDate.ToString("yyyy-MM-dd HH:mm:ss") 
        } else { "N/A" }
        'Newest Item' = if ($folder.NewestItemReceivedDate) { 
            $folder.NewestItemReceivedDate.ToString("yyyy-MM-dd HH:mm:ss") 
        } else { "N/A" }
    }
}

# Perform mailbox health analysis
Write-Host "Performing mailbox health analysis..." -ForegroundColor Yellow
$healthAnalysis = Get-MailboxHealthAnalysis -MailboxInfo $mailboxInfo -FolderDetails $folderDetails -MailboxConfig $mailboxConfig

# Get retention policy information
$retentionPolicyInfo = Get-RetentionPolicyInfo -MailboxInfo $mailboxInfo

# Generate HTML Report
Write-Host "Generating HTML report..." -ForegroundColor Yellow

# Sort folders by size (descending)
$sortedFolders = $folderDetails | Sort-Object -Property 'Folder Size (Bytes)' -Descending

# HTML Content with tabs
$html = @"
<!DOCTYPE html><html lang="en"> <head> <meta charset="UTF-8"> <meta name="viewport" content="width=device-width, initial-scale=1.0"> <title>Mailbox Report - $($mailbox.DisplayName)</title> <style> * { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
text

    body {
        background-color: #f5f7fa;
        color: #2c3e50;
        line-height: 1.6;
        padding: 20px;
    }
    
    .container {
        max-width: 1800px;
        margin: 0 auto;
        background: white;
        border-radius: 10px;
        box-shadow: 0 5px 15px rgba(0, 0, 0, 0.08);
        overflow: hidden;
    }
    
    header {
        background: linear-gradient(135deg, #3498db 0%, #2c3e50 100%);
        color: white;
        padding: 25px 30px;
    }
    
    h1 {
        font-size: 28px;
        font-weight: 600;
        margin-bottom: 5px;
    }
    
    .user-info {
        display: flex;
        gap: 20px;
        margin-top: 10px;
        flex-wrap: wrap;
    }
    
    .user-info div {
        background: rgba(255, 255, 255, 0.15);
        padding: 8px 15px;
        border-radius: 20px;
        font-size: 14px;
    }
    
    /* Tabs styling */
    .tabs {
        display: flex;
        background-color: #f8f9fa;
        border-bottom: 1px solid #e9ecef;
    }
    
    .tab {
        padding: 15px 25px;
        cursor: pointer;
        font-weight: 600;
        color: #7f8c8d;
        border-bottom: 3px solid transparent;
        transition: all 0.3s ease;
    }
    
    .tab:hover {
        color: #3498db;
        background-color: rgba(52, 152, 219, 0.05);
    }
    
    .tab.active {
        color: #3498db;
        border-bottom-color: #3498db;
        background-color: white;
    }
    
    .tab-content {
        display: none;
        padding: 25px;
    }
    
    .tab-content.active {
        display: block;
    }
    
    /* Summary cards */
    .summary-cards {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
        gap: 20px;
        margin-bottom: 25px;
    }
    
    .card {
        background: white;
        border-radius: 8px;
        padding: 20px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        text-align: center;
        border-left: 4px solid #3498db;
    }
    
    .card.warning {
        border-left-color: #e74c3c;
    }
    
    .card.info {
        border-left-color: #3498db;
    }
    
    .card.recommendation {
        border-left-color: #27ae60;
    }
    
    .card-value {
        font-size: 24px;
        font-weight: 700;
        color: #2c3e50;
        margin-bottom: 5px;
    }
    
    .card-label {
        font-size: 14px;
        color: #7f8c8d;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    /* Health analysis */
    .health-analysis {
        margin-bottom: 25px;
    }
    
    .analysis-section {
        margin-bottom: 20px;
    }
    
    .analysis-title {
        font-size: 18px;
        font-weight: 600;
        margin-bottom: 15px;
        color: #2c3e50;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    
    .analysis-items {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
        gap: 15px;
    }
    
    .analysis-item {
        background: white;
        border-radius: 8px;
        padding: 15px;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
        border-left: 4px solid #3498db;
    }
    
    .analysis-item.warning {
        border-left-color: #e74c3c;
    }
    
    .analysis-item.info {
        border-left-color: #3498db;
    }
    
    .analysis-item.recommendation {
        border-left-color: #27ae60;
    }
    
    .analysis-item-title {
        font-weight: 600;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    
    .analysis-item-content {
        font-size: 14px;
        color: #546e7a;
    }
    
    /* Mailbox info table */
    .info-table {
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 25px;
    }
    
    .info-table th {
        background-color: #f1f8ff;
        padding: 12px 15px;
        text-align: left;
        font-weight: 600;
        color: #2c3e50;
        border-bottom: 2px solid #ddd;
    }
    
    .info-table td {
        padding: 12px 15px;
        border-bottom: 1px solid #e9ecef;
    }
    
    .info-table tr:hover {
        background-color: #f8f9fa;
    }
    
    /* Folder table */
    .folder-table-container {
        overflow-x: auto;
    }
    
    .folder-table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 15px;
        font-size: 14px;
        min-width: 1000px;
    }
    
    .folder-table th {
        background-color: #f1f8ff;
        padding: 12px 15px;
        text-align: left;
        font-weight: 600;
        color: #2c3e50;
        border-bottom: 2px solid #ddd;
    }
    
    .folder-table td {
        padding: 12px 15px;
        border-bottom: 1px solid #e9ecef;
    }
    
    .folder-table tr:hover {
        background-color: #f8f9fa;
    }
    
    .folder-primary {
        background-color: #f0f9ff;
    }
    
    .folder-archive {
        background-color: #f0fff4;
    }
    
    .size-col {
        text-align: right;
        font-family: monospace;
    }
    
    .count-col {
        text-align: right;
        font-family: monospace;
    }
    
    .folder-type {
        font-size: 12px;
        padding: 3px 8px;
        border-radius: 12px;
        background-color: #e3f2fd;
        color: #1565c0;
    }
    
    .retention-none {
        color: #e74c3c;
    }
    
    .retention-set {
        color: #27ae60;
    }
    
    /* Controls */
    .controls {
        display: flex;
        justify-content: space-between;
        padding: 15px 0;
        flex-wrap: wrap;
        gap: 15px;
    }
    
    .filter-group {
        display: flex;
        gap: 15px;
        flex-wrap: wrap;
    }
    
    .filter-item {
        display: flex;
        align-items: center;
        gap: 8px;
    }
    
    .filter-item label {
        font-size: 14px;
        font-weight: 600;
        color: #2c3e50;
    }
    
    select, input {
        padding: 8px 12px;
        border: 1px solid #ddd;
        border-radius: 4px;
        font-size: 14px;
        min-width: 150px;
    }
    
    .view-options {
        display: flex;
        gap: 10px;
    }
    
    .btn {
        padding: 8px 16px;
        background-color: white;
        border: 1px solid #ddd;
        border-radius: 4px;
        cursor: pointer;
        font-size: 14px;
        display: flex;
        align-items: center;
        gap: 5px;
        transition: all 0.2s ease;
    }
    
    .btn:hover {
        background-color: #f1f1f1;
    }
    
    .btn.active {
        background-color: #3498db;
        color: white;
        border-color: #3498db;
    }
    
    /* Footer */
    footer {
        text-align: center;
        padding: 20px;
        color: #7f8c8d;
        font-size: 14px;
        border-top: 1px solid #e9ecef;
    }
    
    @media (max-width: 768px) {
        .summary-cards {
            grid-template-columns: 1fr;
        }
        
        .tabs {
            flex-direction: column;
        }
        
        .tab {
            border-bottom: 1px solid #e9ecef;
            border-left: 3px solid transparent;
        }
        
        .tab.active {
            border-left-color: #3498db;
            border-bottom-color: #e9ecef;
        }
        
        .controls {
            flex-direction: column;
        }
        
        .filter-group {
            flex-direction: column;
            align-items: flex-start;
        }
        
        .view-options {
            width: 100%;
            justify-content: center;
        }
    }
</style>
</head> <body> <div class="container"> <header> <h1>Mailbox Analysis Report</h1> <div class="user-info"> <div>$($mailbox.DisplayName)</div> <div>$($mailbox.PrimarySmtpAddress)</div> <div>Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm")</div> </div> </header>
text

    <div class="tabs">
        <div class="tab active" data-tab="overview">Overview</div>
        <div class="tab" data-tab="mailbox-info">Mailbox Information</div>
        <div class="tab" data-tab="health-analysis">Health Analysis</div>
        <div class="tab" data-tab="warnings">Warnings</div>
        <div class="tab" data-tab="recommendations">Recommendations</div>
        <div class="tab" data-tab="folder-stats">Folder Statistics</div>
    </div>
    
    <!-- Overview Tab -->
    <div class="tab-content active" id="overview">
        <h2>Mailbox Overview</h2>
        
        <div class="summary-cards">
            <div class="card">
                <div class="card-value">$($mailboxInfo.'Primary Mailbox Size')</div>
                <div class="card-label">Primary Mailbox Size</div>
            </div>
            <div class="card">
                <div class="card-value">$($mailboxInfo.'Archive Mailbox Size')</div>
                <div class="card-label">Archive Mailbox Size</div>
            </div>
            <div class="card $(if($mailboxInfo.'Retention Hold' -eq 'Enabled'){'warning'})">
                <div class="card-value">$($mailboxInfo.'Retention Hold')</div>
                <div class="card-label">Retention Hold Status</div>
            </div>
            <div class="card $(if($mailboxInfo.'Litigation Hold' -eq 'Enabled'){'warning'})">
                <div class="card-value">$($mailboxInfo.'Litigation Hold')</div>
                <div class="card-label">Litigation Hold Status</div>
            </div>
        </div>
        
        <div class="health-analysis">
            <div class="analysis-section">
                <h3 class="analysis-title">Quick Health Summary</h3>
                <div class="analysis-items">
"@

text

# Add health summary items
$summaryItems = $healthAnalysis.Analysis | Select-Object -First 4
if ($summaryItems.Count -gt 0) {
    foreach ($item in $summaryItems) {
        $severityClass = $item.Severity.ToLower()
        $html += @"
                    <div class="analysis-item $severityClass">
                        <div class="analysis-item-title">
                            <span>$($item.Type)</span>
                        </div>
                        <div class="analysis-item-content">$($item.Message)</div>
                    </div>
"@
}
} else {
$html += @"
<div class="analysis-item info">
<div class="analysis-item-content">No significant issues detected</div>
</div>
"@
}

text

$html += @"
                </div>
            </div>
        </div>
        
        <div class="analysis-section">
            <h3 class="analysis-title">Top 5 Largest Folders</h3>
            <table class="info-table">
                <thead>
                    <tr>
                        <th>Folder Name</th>
                        <th>Location</th>
                        <th>Size</th>
                        <th>Items</th>
                        <th>Retention Tag</th>
                    </tr>
                </thead>
                <tbody>
"@

text

# Add top 5 folders
$topFolders = $sortedFolders | Select-Object -First 5
foreach ($folder in $topFolders) {
    $html += @"
                    <tr>
                        <td>$($folder.'Folder Name')</td>
                        <td>$($folder.Location)</td>
                        <td>$($folder.'Folder Size')</td>
                        <td>$($folder.'Item Count')</td>
                        <td>$($folder.'Retention Tag')</td>
                    </tr>
"@
}

text

$html += @"
                </tbody>
            </table>
        </div>
    </div>
    
    <!-- Mailbox Information Tab -->
    <div class="tab-content" id="mailbox-info">
        <h2>Mailbox Details</h2>
        
        <table class="info-table">
            <tr>
                <th>Property</th>
                <th>Value</th>
                <th>Property</th>
                <th>Value</th>
            </tr>
            <tr>
                <td>Display Name</td>
                <td>$($mailboxInfo.'Display Name')</td>
                <td>Primary SMTP Address</td>
                <td>$($mailboxInfo.'Primary SMTP Address')</td>
            </tr>
            <tr>
                <td>User Principal Name</td>
                <td>$($mailboxInfo.'User Principal Name')</td>
                <td>SAM Account Name</td>
                <td>$($mailboxInfo.'SAM Account Name')</td>
            </tr>
            <tr>
                <td>Retention Policy</td>
                <td>$($mailboxInfo.'Retention Policy')</td>
                <td>Archive Status</td>
                <td>$($mailboxInfo.'Archive Status')</td>
            </tr>
            <tr>
                <td>Litigation Hold</td>
                <td>$($mailboxInfo.'Litigation Hold')</td>
                <td>Litigation Hold Date</td>
                <td>$($mailboxInfo.'Litigation Hold Date')</td>
            </tr>
            <tr>
                <td>Retention Hold</td>
                <td>$($mailboxInfo.'Retention Hold')</td>
                <td>Retention Hold Period</td>
                <td>$($mailboxInfo.'Retention Hold Start Date') to $($mailboxInfo.'Retention Hold End Date')</td>
            </tr>
            <tr>
                <td>Single Item Recovery</td>
                <td>$($mailboxInfo.'Single Item Recovery')</td>
                <td>Primary Mailbox Size</td>
                <td>$($mailboxInfo.'Primary Mailbox Size')</td>
            </tr>
            <tr>
                <td>Primary Item Count</td>
                <td>$($mailboxInfo.'Primary Item Count')</td>
                <td>Archive Mailbox Size</td>
                <td>$($mailboxInfo.'Archive Mailbox Size')</td>
            </tr>
            <tr>
                <td>Archive Item Count</td>
                <td>$($mailboxInfo.'Archive Item Count')</td>
                <td>Issue Warning Quota</td>
                <td>$($mailboxInfo.'Issue Warning Quota')</td>
            </tr>
            <tr>
                <td>Prohibit Send Quota</td>
                <td>$($mailboxInfo.'Prohibit Send Quota')</td>
                <td>Prohibit Send/Receive Quota</td>
                <td>$($mailboxInfo.'Prohibit Send/Receive Quota')</td>
            </tr>
            <tr>
                <td>Archive Warning Quota</td>
                <td>$($mailboxInfo.'Archive Warning Quota')</td>
                <td>Archive Quota</td>
                <td>$($mailboxInfo.'Archive Quota')</td>
            </tr>
        </table>
        
        <div class="analysis-section">
            <h3 class="analysis-title">Retention Policy Details</h3>
            <div class="analysis-item info">
                <div class="analysis-item-content">
                    <pre>$retentionPolicyInfo</pre>
                </div>
            </div>
        </div>
    </div>
    
    <!-- Health Analysis Tab -->
    <div class="tab-content" id="health-analysis">
        <h2>Mailbox Health Analysis</h2>
        
        <div class="analysis-items">
"@

text

# Add health analysis items
foreach ($item in $healthAnalysis.Analysis) {
    $severityClass = $item.Severity.ToLower()
    $html += @"
            <div class="analysis-item $severityClass">
                <div class="analysis-item-title">
                    <span>$($item.Type)</span>
                </div>
                <div class="analysis-item-content">$($item.Message)</div>
            </div>
"@
}

text

$html += @"
        </div>
    </div>
    
    <!-- Warnings Tab -->
    <div class="tab-content" id="warnings">
        <h2>Mailbox Warnings</h2>
        
        <div class="analysis-items">
"@

text

if ($healthAnalysis.Warnings.Count -gt 0) {
    foreach ($warning in $healthAnalysis.Warnings) {
        $html += @"
            <div class="analysis-item warning">
                <div class="analysis-item-content">$warning</div>
            </div>
"@
}
} else {
$html += @"
<div class="analysis-item info">
<div class="analysis-item-content">No warnings detected</div>
</div>
"@
}

text

$html += @"
        </div>
    </div>
    
    <!-- Recommendations Tab -->
    <div class="tab-content" id="recommendations">
        <h2>Mailbox Recommendations</h2>
        
        <div class="analysis-items">
"@

text

if ($healthAnalysis.Recommendations.Count -gt 0) {
    foreach ($recommendation in $healthAnalysis.Recommendations) {
        $html += @"
            <div class="analysis-item recommendation">
                <div class="analysis-item-content">$recommendation</div>
            </div>
"@
}
} else {
$html += @"
<div class="analysis-item info">
<div class="analysis-item-content">No recommendations at this time</div>
</div>
"@
}

text

$html += @"
        </div>
    </div>
    
    <!-- Folder Statistics Tab -->
    <div class="tab-content" id="folder-stats">
        <h2>Folder Statistics</h2>
        
        <div class="controls">
            <div class="filter-group">
                <div class="filter-item">
                    <label for="location-filter">Location:</label>
                    <select id="location-filter">
                        <option value="all">All Locations</option>
                        <option value="Primary Mailbox">Primary Mailbox</option>
                        <option value="Archive Mailbox">Archive Mailbox</option>
                    </select>
                </div>
                <div class="filter-item">
                    <label for="size-filter">Min Size:</label>
                    <select id="size-filter">
                        <option value="0">All Sizes</option>
                        <option value="1048576">>1 MB</option>
                        <option value="10485760">>10 MB</option>
                        <option value="104857600">>100 MB</option>
                        <option value="1073741824">>1 GB</option>
                    </select>
                </div>
                <div class="filter-item">
                    <label for="search">Search:</label>
                    <input type="text" id="search" placeholder="Search folders...">
                </div>
            </div>
            <div class="view-options">
                <button class="btn active" id="view-detailed">Detailed View</button>
                <button class="btn" id="view-summary">Summary View</button>
            </div>
        </div>
        
        <div class="folder-table-container">
            <table class="folder-table" id="folder-table">
                <thead>
                    <tr>
                        <th>Location</th>
                        <th>Folder Name</th>
                        <th>Folder Type</th>
                        <th class="size-col">Size</th>
                        <th class="count-col">Items</th>
                        <th class="count-col">Subfolders</th>
                        <th class="count-col">Deleted Items</th>
                        <th>Retention Tag</th>
                        <th>Oldest Item</th>
                        <th>Newest Item</th>
                    </tr>
                </thead>
                <tbody>
"@

text

# Add folder rows
foreach ($folder in $sortedFolders) {
    $locationClass = if ($folder.Location -eq "Primary Mailbox") { "folder-primary" } else { "folder-archive" }
    $retentionClass = if ($folder.'Retention Tag' -eq "None") { "retention-none" } else { "retention-set" }
    
    $html += @"
                    <tr class="$locationClass">
                        <td>$($folder.Location)</td>
                        <td>$($folder.'Folder Name')</td>
                        <td><span class="folder-type">$($folder.'Folder Type')</span></td>
                        <td class="size-col">$($folder.'Folder Size')</td>
                        <td class="count-col">$($folder.'Item Count')</td>
                        <td class="count-col">$($folder.'Subfolder Count')</td>
                        <td class="count-col">$($folder.'Deleted Items')</td>
                        <td class="$retentionClass">$($folder.'Retention Tag')</td>
                        <td>$($folder.'Oldest Item')</td>
                        <td>$($folder.'Newest Item')</td>
                    </tr>
"@
}

text

$html += @"
                </tbody>
            </table>
        </div>
    </div>
    
    <footer>
        <p>Report generated on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") by eDLS Litigation Services Team</p>
	const footer = document.getElementById('page-footer');
    </footer>
</div>

<script>
    // Tab functionality
    document.addEventListener('DOMContentLoaded', function() {
        const tabs = document.querySelectorAll('.tab');
        const tabContents = document.querySelectorAll('.tab-content');
        
        tabs.forEach(tab => {
            tab.addEventListener('click', () => {
                const tabId = tab.getAttribute('data-tab');
                
                // Remove active class from all tabs and contents
                tabs.forEach(t => t.classList.remove('active'));
                tabContents.forEach(c => c.classList.remove('active'));
                
                // Add active class to clicked tab and corresponding content
                tab.classList.add('active');
                document.getElementById(tabId).classList.add('active');
            });
        });
        
        // Folder table filtering and view toggle
        const folderTable = document.getElementById('folder-table');
        const folderRows = folderTable.querySelectorAll('tbody tr');
        const locationFilter = document.getElementById('location-filter');
        const sizeFilter = document.getElementById('size-filter');
        const searchInput = document.getElementById('search');
        const viewDetailedBtn = document.getElementById('view-detailed');
        const viewSummaryBtn = document.getElementById('view-summary');
        
        // Function to filter rows
        function filterRows() {
            const locationValue = locationFilter.value;
            const sizeValue = parseInt(sizeFilter.value);
            const searchValue = searchInput.value.toLowerCase();
            
            folderRows.forEach(row => {
                const location = row.cells[0].textContent;
                const sizeText = row.cells[3].textContent;
                const folderName = row.cells[1].textContent.toLowerCase();
                
                // Convert size to bytes for comparison
                let sizeBytes = 0;
                if (sizeText.includes('TB')) {
                    sizeBytes = parseFloat(sizeText) * 1099511627776;
                } else if (sizeText.includes('GB')) {
                    sizeBytes = parseFloat(sizeText) * 1073741824;
                } else if (sizeText.includes('MB')) {
                    sizeBytes = parseFloat(sizeText) * 1048576;
                } else if (sizeText.includes('KB')) {
                    sizeBytes = parseFloat(sizeText) * 1024;
                } else if (sizeText.includes('B')) {
                    sizeBytes = parseFloat(sizeText);
                }
                
                const locationMatch = locationValue === 'all' || location === locationValue;
                const sizeMatch = sizeValue === 0 || sizeBytes >= sizeValue;
                const searchMatch = folderName.includes(searchValue);
                
                if (locationMatch && sizeMatch && searchMatch) {
                    row.style.display = '';
                } else {
                    row.style.display = 'none';
                }
            });
        }
        
        // Event listeners for filters
        locationFilter.addEventListener('change', filterRows);
        sizeFilter.addEventListener('change', filterRows);
        searchInput.addEventListener('input', filterRows);
        
        // View toggle functionality
        viewDetailedBtn.addEventListener('click', function() {
            viewDetailedBtn.classList.add('active');
            viewSummaryBtn.classList.remove('active');
            // Show all columns
            for (let i = 0; i < folderTable.rows[0].cells.length; i++) {
                folderTable.rows[0].cells[i].style.display = '';
            }
            for (let i = 1; i < folderTable.rows.length; i++) {
                for (let j = 0; j < folderTable.rows[i].cells.length; j++) {
                    folderTable.rows[i].cells[j].style.display = '';
                }
            }
        });
        
        viewSummaryBtn.addEventListener('click', function() {
            viewSummaryBtn.classList.add('active');
            viewDetailedBtn.classList.remove('active');
            // Hide some columns for summary view
            const columnsToHide = [5, 6, 7, 8, 9]; // Column indices to hide (0-based)
            for (let i = 0; i < folderTable.rows[0].cells.length; i++) {
                if (columnsToHide.includes(i)) {
                    folderTable.rows[0].cells[i].style.display = 'none';
                }
            }
            for (let i = 1; i < folderTable.rows.length; i++) {
                for (let j = 0; j < folderTable.rows[i].cells.length; j++) {
                    if (columnsToHide.includes(j)) {
                        folderTable.rows[i].cells[j].style.display = 'none';
                    }
                }
            }
        });
    });
</script>
</body> </html> "@
text

# Save HTML report
$html | Out-File -FilePath $OutputPath -Encoding UTF8
Write-Host "HTML report saved to: $OutputPath" -ForegroundColor Green

# Export to CSV if requested
if ($ExportToCSV) {
    $csvPath = $OutputPath -replace '\.html$', '.csv'
    $folderDetails | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "CSV data exported to: $csvPath" -ForegroundColor Green
}

# Display summary in console
Write-Host "`n=== MAILBOX SUMMARY ===" -ForegroundColor Cyan
Write-Host "User: $($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))" -ForegroundColor White
Write-Host "Primary Mailbox Size: $($mailboxInfo.'Primary Mailbox Size')" -ForegroundColor White
Write-Host "Archive Mailbox Size: $($mailboxInfo.'Archive Mailbox Size')" -ForegroundColor White
Write-Host "Retention Hold: $($mailboxInfo.'Retention Hold')" -ForegroundColor White
if ($mailboxInfo.'Retention Hold' -eq "Enabled") {
    Write-Host "Retention Hold Period: $($mailboxInfo.'Retention Hold Start Date') to $($mailboxInfo.'Retention Hold End Date')" -ForegroundColor White
}

Write-Host "`n=== HEALTH ANALYSIS ===" -ForegroundColor Cyan
if ($healthAnalysis.Analysis.Count -gt 0) {
    foreach ($item in $healthAnalysis.Analysis) {
        $color = switch ($item.Severity) {
            "High" { "Red" }
            "Medium" { "Yellow" }
            "Low" { "Green" }
            default { "White" }
        }
        Write-Host "[$($item.Type)] $($item.Message)" -ForegroundColor $color
    }
} else {
    Write-Host "No issues detected" -ForegroundColor Green
}

Write-Host "`n=== TOP 5 LARGEST FOLDERS ===" -ForegroundColor Cyan
$topFolders = $sortedFolders | Select-Object -First 5
foreach ($folder in $topFolders) {
    Write-Host "$($folder.Location): $($folder.'Folder Name') - $($folder.'Folder Size') ($($folder.'Item Count') items)" -ForegroundColor White
}

# Open the report in default browser
try {
    Write-Host "`nOpening report in default browser..." -ForegroundColor Yellow
    Start-Process $OutputPath
}
catch {
    Write-Host "Could not open report automatically. Please open the file manually: $OutputPath" -ForegroundColor Yellow
}
}
catch {
Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
Write-Host "Script terminated." -ForegroundColor Red
}
