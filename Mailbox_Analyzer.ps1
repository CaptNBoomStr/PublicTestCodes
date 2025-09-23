<#
.SYNOPSIS
    Novartis Exchange Online Mailbox Analysis Report - Enhanced Edition v3.0
.DESCRIPTION
    Comprehensive analysis of Exchange Online mailboxes including storage, quota, retention,
    permissions, rules, delegates, forwarding, security scoring, and detailed folder analysis.
.AUTHOR
    eDLS Team - Mohd Azhar Uddin
.VERSION
    3.0
.UPDATES
    - Added archive and auto-expanding archive status
    - Added total mailbox count (Primary, Archive, Auxiliary)
    - Enhanced Inbox Rules Analysis with detailed reporting
    - Added Folder Permissions display (Inbox, Calendar, Contacts, etc.)
    - Added comprehensive Security Overview with intelligent scoring
#>

#Requires -Modules ExchangeOnlineManagement

# Function to display banner
function Show-Banner {
    Write-Host "`n" -NoNewline
    Write-Host "==========================================================" -ForegroundColor Cyan
    Write-Host "  Novartis Exchange Online Mailbox Analysis Report v3.0  " -ForegroundColor Yellow
    Write-Host "==========================================================" -ForegroundColor Cyan
    Write-Host ""
}

# Function to get user input
function Get-UserInput {
    param(
        [string]$Prompt,
        [string]$DefaultValue = ""
    )
    
    if ($DefaultValue) {
        $input = Read-Host "$Prompt [$DefaultValue]"
        if ([string]::IsNullOrWhiteSpace($input)) {
            return $DefaultValue
        }
        return $input
    } else {
        do {
            $input = Read-Host $Prompt
        } while ([string]::IsNullOrWhiteSpace($input))
        return $input
    }
}

# Function to connect to Exchange Online
function Connect-ExchangeOnlineSession {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    try {
        # Check if already connected
        $test = Get-EXOMailbox -ResultSize 1 -ErrorAction SilentlyContinue
        if ($test) {
            Write-Host "Already connected to Exchange Online" -ForegroundColor Green
            return $true
        }
        
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Host "Successfully connected to Exchange Online" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to connect to Exchange Online: $_" -ForegroundColor Red
        return $false
    }
}

# Function to convert size strings to bytes
function Convert-SizeToBytes {
    param($Size)
    
    if ($Size -eq $null -or $Size -eq 0) {
        return 0
    }
    
    # Handle if Size is already a number
    if ($Size -is [int] -or $Size -is [long] -or $Size -is [double]) {
        return [long]$Size
    }
    
    # Handle if Size is a string
    if ($Size -is [string]) {
        if ($Size -match '([\d.]+)\s*([KMGT]B).*KATEX_INLINE_OPEN([\d,]+)\s*bytesKATEX_INLINE_CLOSE') {
            return [long]($matches[3] -replace ',', '')
        }
        elseif ($Size -match '([\d,]+)\s*bytes') {
            return [long]($matches[1] -replace ',', '')
        }
        elseif ($Size -match '([\d,\.]+)\s*([KMGT]B)') {
            $number = [double]($matches[1] -replace ',', '')
            $unit = $matches[2]
            
            switch ($unit) {
                'KB' { return $number * 1KB }
                'MB' { return $number * 1MB }
                'GB' { return $number * 1GB }
                'TB' { return $number * 1TB }
                default { return $number }
            }
        }
        else {
            return 0
        }
    }
    
    # Handle ByteQuantifiedSize objects
    if ($Size.GetType().Name -eq 'ByteQuantifiedSize' -or $Size.PSObject.Methods.Name -contains 'ToBytes') {
        return $Size.ToBytes()
    }
    
    # Handle objects with Value property
    if ($Size.Value -and $Size.Value.PSObject.Methods.Name -contains 'ToBytes') {
        return $Size.Value.ToBytes()
    }
    
    return 0
}

# Function to format bytes to readable string
function Format-ByteSize {
    param([long]$Bytes)
    
    if ($Bytes -eq 0) { return "0 B" }
    
    $sizes = @('B', 'KB', 'MB', 'GB', 'TB')
    $i = [Math]::Floor([Math]::Log($Bytes, 1024))
    $size = [Math]::Round($Bytes / [Math]::Pow(1024, $i), 2)
    
    return "$size $($sizes[$i])"
}

# Function to calculate percentage
function Get-Percentage {
    param(
        [long]$Used,
        [long]$Limit
    )
    
    if ($Limit -eq 0) { return 0 }
    return [Math]::Round(($Used / $Limit) * 100, 2)
}

# Function to determine mailbox health status
function Get-MailboxHealthStatus {
    param(
        [long]$UsedSize,
        [long]$WarningQuota,
        [long]$ProhibitSendQuota,
        [long]$ProhibitSendReceiveQuota
    )
    
    if ($ProhibitSendReceiveQuota -gt 0 -and $UsedSize -ge $ProhibitSendReceiveQuota) {
        return @{
            Status = "Prohibit Send/Receive"
            Color = "#e74c3c"
            Icon = "&#10060;"
        }
    }
    elseif ($ProhibitSendQuota -gt 0 -and $UsedSize -ge $ProhibitSendQuota) {
        return @{
            Status = "Prohibit Send"
            Color = "#ff9800"
            Icon = "&#9940;"
        }
    }
    elseif ($WarningQuota -gt 0 -and $UsedSize -ge $WarningQuota) {
        return @{
            Status = "Warning"
            Color = "#ffc107"
            Icon = "&#9888;"
        }
    }
    else {
        return @{
            Status = "Healthy"
            Color = "#4caf50"
            Icon = "&#9989;"
        }
    }
}

# Enhanced function to get complete retention information
function Get-CompleteRetentionInfo {
    param([string]$UserPrincipalName)
    
    Write-Host "  Retrieving complete retention information..." -ForegroundColor Gray
    
    $retentionInfo = @{
        PolicyName = "None"
        PolicyTags = @()
        AllTags = @{}
        TagsByName = @{}
    }
    
    try {
        # Get mailbox retention policy
        $mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
        
        if ($mailbox.RetentionPolicy) {
            $retentionInfo.PolicyName = $mailbox.RetentionPolicy
            Write-Host "    Retention Policy: $($mailbox.RetentionPolicy)" -ForegroundColor Gray
            
            # Get retention policy details
            try {
                $policy = Get-RetentionPolicy -Identity $mailbox.RetentionPolicy -ErrorAction Stop
                if ($policy -and $policy.RetentionPolicyTagLinks) {
                    # Get tags linked to this policy
                    foreach ($tagLink in $policy.RetentionPolicyTagLinks) {
                        try {
                            $tag = Get-RetentionPolicyTag -Identity $tagLink -ErrorAction Stop
                            if ($tag) {
                                $retentionInfo.PolicyTags += $tag
                                Write-Host "      Found policy tag: $($tag.Name)" -ForegroundColor Gray
                            }
                        }
                        catch {
                            # Individual tag might not be accessible
                        }
                    }
                }
            }
            catch {
                Write-Host "    Could not retrieve retention policy details" -ForegroundColor Yellow
            }
        }
        
        # Get all retention tags in the organization
        try {
            $allTags = Get-RetentionPolicyTag -ErrorAction Stop
            foreach ($tag in $allTags) {
                $retentionInfo.AllTags[$tag.Guid.ToString()] = $tag
                $retentionInfo.TagsByName[$tag.Name] = $tag
                # Also add without spaces for matching
                $nameNoSpaces = $tag.Name -replace '\s', ''
                $retentionInfo.TagsByName[$nameNoSpaces] = $tag
            }
            Write-Host "    Found $($allTags.Count) total retention tags in organization" -ForegroundColor Gray
        }
        catch {
            Write-Host "    Warning: Could not retrieve retention policy tags" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "    Error retrieving retention information: $_" -ForegroundColor Yellow
    }
    
    return $retentionInfo
}

# Enhanced function to get inbox rules analysis
function Get-InboxRulesAnalysis {
    param([string]$UserPrincipalName)
    
    Write-Host "  Analyzing inbox rules..." -ForegroundColor Gray
    $rulesAnalysis = @{
        TotalRules = 0
        EnabledRules = 0
        DisabledRules = 0
        ForwardingRules = 0
        DeletingRules = 0
        MovingRules = 0
        SuspiciousRules = @()
        AllRules = @()
    }
    
    try {
        $rules = Get-InboxRule -Mailbox $UserPrincipalName -ErrorAction Stop
        $rulesAnalysis.TotalRules = $rules.Count
        
        foreach ($rule in $rules) {
            $ruleInfo = @{
                Name = $rule.Name
                Enabled = $rule.Enabled
                Priority = $rule.Priority
                Description = if ($rule.Description) { $rule.Description } else { "No description" }
                Conditions = @()
                Actions = @()
                IsSuspicious = $false
                SuspiciousReasons = @()
                RiskLevel = "Low"
            }
            
            # Track enabled/disabled
            if ($rule.Enabled) { $rulesAnalysis.EnabledRules++ } else { $rulesAnalysis.DisabledRules++ }
            
            # Check for forwarding
            if ($rule.ForwardTo -or $rule.ForwardAsAttachmentTo -or $rule.RedirectTo) {
                $rulesAnalysis.ForwardingRules++
                $ruleInfo.Actions += "Forwarding"
                
                # Check if forwarding to external domain
                $forwardAddresses = @()
                if ($rule.ForwardTo) { $forwardAddresses += $rule.ForwardTo }
                if ($rule.ForwardAsAttachmentTo) { $forwardAddresses += $rule.ForwardAsAttachmentTo }
                if ($rule.RedirectTo) { $forwardAddresses += $rule.RedirectTo }
                
                foreach ($addr in $forwardAddresses) {
                    $ruleInfo.Actions += "Forward to: $addr"
                    if ($addr -notmatch "@novartis\.") {
                        $ruleInfo.IsSuspicious = $true
                        $ruleInfo.SuspiciousReasons += "External forwarding to: $addr"
                        $ruleInfo.RiskLevel = "High"
                    }
                }
            }
            
            # Check for deletion
            if ($rule.DeleteMessage) {
                $rulesAnalysis.DeletingRules++
                $ruleInfo.Actions += "Delete Message"
                if ($rule.Enabled) {
                    $ruleInfo.IsSuspicious = $true
                    $ruleInfo.SuspiciousReasons += "Active deletion rule"
                    $ruleInfo.RiskLevel = "Medium"
                }
            }
            
            # Check for permanent deletion
            if ($rule.PermanentDelete) {
                $ruleInfo.Actions += "Permanent Delete"
                $ruleInfo.IsSuspicious = $true
                $ruleInfo.SuspiciousReasons += "Permanent deletion (bypasses Deleted Items)"
                $ruleInfo.RiskLevel = "High"
            }
            
            # Check for moving messages
            if ($rule.MoveToFolder) {
                $rulesAnalysis.MovingRules++
                $ruleInfo.Actions += "Move to folder: $($rule.MoveToFolder)"
            }
            
            # Add conditions
            if ($rule.From) { $ruleInfo.Conditions += "From: $($rule.From -join ', ')" }
            if ($rule.SentTo) { $ruleInfo.Conditions += "Sent To: $($rule.SentTo -join ', ')" }
            if ($rule.SubjectContainsWords) { $ruleInfo.Conditions += "Subject Contains: $($rule.SubjectContainsWords -join ', ')" }
            if ($rule.BodyContainsWords) { $ruleInfo.Conditions += "Body Contains: $($rule.BodyContainsWords -join ', ')" }
            if ($rule.FromAddressContainsWords) { $ruleInfo.Conditions += "From Address Contains: $($rule.FromAddressContainsWords -join ', ')" }
            
            # Additional suspicious patterns
            if ($rule.MarkAsRead -and $rule.DeleteMessage) {
                $ruleInfo.IsSuspicious = $true
                $ruleInfo.SuspiciousReasons += "Marks as read and deletes (hiding activity)"
                $ruleInfo.RiskLevel = "High"
            }
            
            if ($ruleInfo.IsSuspicious) {
                $rulesAnalysis.SuspiciousRules += $ruleInfo
            }
            
            $rulesAnalysis.AllRules += $ruleInfo
        }
        
        Write-Host "    Found $($rulesAnalysis.TotalRules) inbox rules" -ForegroundColor Gray
        if ($rulesAnalysis.SuspiciousRules.Count -gt 0) {
            Write-Host "    WARNING: $($rulesAnalysis.SuspiciousRules.Count) suspicious rules detected!" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "    Error analyzing inbox rules: $_" -ForegroundColor Yellow
    }
    
    return $rulesAnalysis
}

# Enhanced function to get delegates and permissions with folder permissions
function Get-DelegatesAndPermissions {
    param([string]$UserPrincipalName)
    
    Write-Host "  Analyzing delegates and permissions..." -ForegroundColor Gray
    $permissionsAnalysis = @{
        FullAccessDelegates = @()
        SendAsDelegates = @()
        SendOnBehalfDelegates = @()
        FolderPermissions = @{}
        UnusualPermissions = @()
        TotalDelegateCount = 0
    }
    
    try {
        # Get mailbox for SendOnBehalf
        $mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
        
        # Get Full Access permissions
        Write-Host "    Checking Full Access permissions..." -ForegroundColor Gray
        try {
            $fullAccess = Get-MailboxPermission -Identity $UserPrincipalName -ErrorAction Stop | 
                Where-Object { $_.IsInherited -eq $false -and $_.User -notlike "NT AUTHORITY\*" -and $_.User -notlike "S-1-5-*" }
            
            foreach ($perm in $fullAccess) {
                if ($perm.AccessRights -contains "FullAccess") {
                    $delegate = @{
                        User = $perm.User.ToString()
                        AccessRights = $perm.AccessRights -join ", "
                        IsInherited = $perm.IsInherited
                        Deny = $perm.Deny
                    }
                    $permissionsAnalysis.FullAccessDelegates += $delegate
                }
            }
        }
        catch {
            Write-Host "      Could not retrieve Full Access permissions" -ForegroundColor Yellow
        }
        
        # Get Send As permissions
        Write-Host "    Checking Send As permissions..." -ForegroundColor Gray
        try {
            $sendAs = Get-RecipientPermission -Identity $UserPrincipalName -ErrorAction Stop | 
                Where-Object { $_.Trustee -ne "NT AUTHORITY\SELF" }
            
            foreach ($perm in $sendAs) {
                $delegate = @{
                    User = $perm.Trustee
                    AccessRights = $perm.AccessRights -join ", "
                }
                $permissionsAnalysis.SendAsDelegates += $delegate
            }
        }
        catch {
            Write-Host "      Could not retrieve Send As permissions" -ForegroundColor Yellow
        }
        
        # Get Send on Behalf permissions from mailbox properties
        Write-Host "    Checking Send on Behalf permissions..." -ForegroundColor Gray
        if ($mailbox.GrantSendOnBehalfTo -and $mailbox.GrantSendOnBehalfTo.Count -gt 0) {
            foreach ($user in $mailbox.GrantSendOnBehalfTo) {
                $delegate = @{
                    User = $user.ToString()
                }
                $permissionsAnalysis.SendOnBehalfDelegates += $delegate
            }
        }
        
        # Enhanced folder-level permissions for key folders
        Write-Host "    Checking folder permissions..." -ForegroundColor Gray
        $foldersToCheck = @("Calendar", "Contacts", "Inbox", "SentItems", "Tasks", "Notes")
        
        foreach ($folderName in $foldersToCheck) {
            try {
                $folderPath = "$($UserPrincipalName):\$folderName"
                $folderPerms = Get-MailboxFolderPermission -Identity $folderPath -ErrorAction SilentlyContinue
                
                $folderPermList = @()
                $defaultAccess = "None"
                $anonymousAccess = "None"
                $hasExternalAccess = $false
                
                foreach ($perm in $folderPerms) {
                    if ($perm.User -like "Default") {
                        $defaultAccess = $perm.AccessRights -join ", "
                        if ($defaultAccess -ne "None" -and $defaultAccess -ne "AvailabilityOnly") {
                            $hasExternalAccess = $true
                        }
                    }
                    elseif ($perm.User -like "Anonymous") {
                        $anonymousAccess = $perm.AccessRights -join ", "
                        if ($anonymousAccess -ne "None") {
                            $hasExternalAccess = $true
                        }
                    }
                    else {
                        $folderPermList += @{
                            User = $perm.User.ToString()
                            AccessRights = $perm.AccessRights -join ", "
                        }
                    }
                }
                
                $permissionsAnalysis.FolderPermissions[$folderName] = @{
                    Permissions = $folderPermList
                    DefaultAccess = $defaultAccess
                    AnonymousAccess = $anonymousAccess
                    HasExternalAccess = $hasExternalAccess
                    TotalPermissions = $folderPermList.Count
                }
                
                if ($folderPermList.Count -gt 0) {
                    Write-Host "      $folderName : $($folderPermList.Count) custom permissions" -ForegroundColor Gray
                }
            }
            catch {
                # Folder might not exist or not accessible
            }
        }
        
        # Check for unusual permissions
        foreach ($folderName in $permissionsAnalysis.FolderPermissions.Keys) {
            $folderPerm = $permissionsAnalysis.FolderPermissions[$folderName]
            if ($folderPerm.HasExternalAccess) {
                $permissionsAnalysis.UnusualPermissions += "External access on $folderName folder"
            }
            if ($folderPerm.TotalPermissions -gt 10) {
                $permissionsAnalysis.UnusualPermissions += "High number of permissions on $folderName folder"
            }
        }
        
        # Calculate total delegate count
        $permissionsAnalysis.TotalDelegateCount = 
            $permissionsAnalysis.FullAccessDelegates.Count + 
            $permissionsAnalysis.SendAsDelegates.Count + 
            $permissionsAnalysis.SendOnBehalfDelegates.Count
        
        Write-Host "    Found $($permissionsAnalysis.TotalDelegateCount) total delegates" -ForegroundColor Gray
    }
    catch {
        Write-Host "    Error analyzing permissions: $_" -ForegroundColor Yellow
    }
    
    return $permissionsAnalysis
}

# Function to get forwarding configuration
function Get-ForwardingConfiguration {
    param([string]$UserPrincipalName)
    
    Write-Host "  Analyzing forwarding configuration..." -ForegroundColor Gray
    $forwardingAnalysis = @{
        ForwardingEnabled = $false
        ForwardingAddress = "None"
        ForwardingSmtpAddress = "None"
        DeliverToMailboxAndForward = $false
        IsExternalForwarding = $false
    }
    
    try {
        $mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
        
        # Check internal forwarding
        if ($mailbox.ForwardingAddress) {
            $forwardingAnalysis.ForwardingEnabled = $true
            $forwardingAnalysis.ForwardingAddress = $mailbox.ForwardingAddress.ToString()
            $forwardingAnalysis.DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
        }
        
        # Check SMTP forwarding (external)
        if ($mailbox.ForwardingSmtpAddress) {
            $forwardingAnalysis.ForwardingEnabled = $true
            $forwardingAnalysis.ForwardingSmtpAddress = $mailbox.ForwardingSmtpAddress.ToString()
            $forwardingAnalysis.DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
            
            # Check if external domain
            if ($mailbox.ForwardingSmtpAddress -notmatch "@novartis\.") {
                $forwardingAnalysis.IsExternalForwarding = $true
            }
        }
        
        Write-Host "    Forwarding status: $(if ($forwardingAnalysis.ForwardingEnabled) { 'Enabled' } else { 'Disabled' })" -ForegroundColor Gray
    }
    catch {
        Write-Host "    Error analyzing forwarding: $_" -ForegroundColor Yellow
    }
    
    return $forwardingAnalysis
}

# Function to analyze mailbox audit settings
function Get-MailboxAuditAnalysis {
    param([string]$UserPrincipalName)
    
    Write-Host "  Analyzing mailbox audit settings..." -ForegroundColor Gray
    $auditAnalysis = @{
        AuditEnabled = $false
        AuditLogAgeLimit = "Not Set"
        AuditOwnerActions = @()
        AuditDelegateActions = @()
        AuditAdminActions = @()
    }
    
    try {
        $mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
        
        $auditAnalysis.AuditEnabled = $mailbox.AuditEnabled
        
        if ($mailbox.AuditLogAgeLimit) {
            $auditAnalysis.AuditLogAgeLimit = $mailbox.AuditLogAgeLimit.ToString()
        }
        
        if ($mailbox.AuditOwner) {
            $auditAnalysis.AuditOwnerActions = $mailbox.AuditOwner
        }
        
        if ($mailbox.AuditDelegate) {
            $auditAnalysis.AuditDelegateActions = $mailbox.AuditDelegate
        }
        
        if ($mailbox.AuditAdmin) {
            $auditAnalysis.AuditAdminActions = $mailbox.AuditAdmin
        }
        
        Write-Host "    Audit status: $(if ($auditAnalysis.AuditEnabled) { 'Enabled' } else { 'Disabled' })" -ForegroundColor Gray
    }
    catch {
        Write-Host "    Error analyzing audit settings: $_" -ForegroundColor Yellow
    }
    
    return $auditAnalysis
}

# Enhanced function to get comprehensive security analysis with scoring
function Get-SecurityAnalysis {
    param(
        [object]$MailboxData
    )
    
    Write-Host "  Performing security analysis..." -ForegroundColor Gray
    
    $securityAnalysis = @{
        OverallScore = 100
        CategoryScores = @{}
        SecurityIssues = @()
        Recommendations = @()
        RiskLevel = "Low"
    }
    
    # Initialize category scores
    $categories = @{
        "Access Control" = 100
        "Forwarding & Rules" = 100
        "Audit & Compliance" = 100
        "Data Protection" = 100
        "External Access" = 100
    }
    
    # Access Control Analysis
    if ($MailboxData.Permissions) {
        # Check for excessive delegates
        if ($MailboxData.Permissions.TotalDelegateCount -gt 5) {
            $categories["Access Control"] -= 20
            $securityAnalysis.SecurityIssues += @{
                Category = "Access Control"
                Issue = "High number of delegates ($($MailboxData.Permissions.TotalDelegateCount))"
                Severity = "Medium"
                Impact = -20
            }
            $securityAnalysis.Recommendations += "Review and minimize delegate access"
        }
        
        # Check for folder permissions
        foreach ($folder in $MailboxData.Permissions.FolderPermissions.Keys) {
            $folderPerm = $MailboxData.Permissions.FolderPermissions[$folder]
            if ($folderPerm.HasExternalAccess) {
                $categories["External Access"] -= 15
                $securityAnalysis.SecurityIssues += @{
                    Category = "External Access"
                    Issue = "External access enabled on $folder folder"
                    Severity = "High"
                    Impact = -15
                }
            }
        }
    }
    
    # Forwarding & Rules Analysis
    if ($MailboxData.Forwarding) {
        if ($MailboxData.Forwarding.IsExternalForwarding) {
            $categories["Forwarding & Rules"] -= 30
            $securityAnalysis.SecurityIssues += @{
                Category = "Forwarding & Rules"
                Issue = "External email forwarding is enabled"
                Severity = "High"
                Impact = -30
            }
            $securityAnalysis.Recommendations += "Disable external forwarding or implement strict controls"
        }
        
        if ($MailboxData.Forwarding.ForwardingEnabled -and -not $MailboxData.Forwarding.DeliverToMailboxAndForward) {
            $categories["Forwarding & Rules"] -= 10
            $securityAnalysis.SecurityIssues += @{
                Category = "Forwarding & Rules"
                Issue = "Forwarding without keeping copy in mailbox"
                Severity = "Medium"
                Impact = -10
            }
        }
    }
    
    if ($MailboxData.InboxRules) {
        if ($MailboxData.InboxRules.SuspiciousRules.Count -gt 0) {
            $impact = [Math]::Min(30, $MailboxData.InboxRules.SuspiciousRules.Count * 10)
            $categories["Forwarding & Rules"] -= $impact
            $securityAnalysis.SecurityIssues += @{
                Category = "Forwarding & Rules"
                Issue = "$($MailboxData.InboxRules.SuspiciousRules.Count) suspicious inbox rules detected"
                Severity = "High"
                Impact = -$impact
            }
            $securityAnalysis.Recommendations += "Review and remove suspicious inbox rules"
        }
        
        if ($MailboxData.InboxRules.DeletingRules -gt 0) {
            $categories["Data Protection"] -= 15
            $securityAnalysis.SecurityIssues += @{
                Category = "Data Protection"
                Issue = "$($MailboxData.InboxRules.DeletingRules) rules that delete messages"
                Severity = "Medium"
                Impact = -15
            }
        }
    }
    
    # Audit & Compliance Analysis
    if ($MailboxData.AuditSettings) {
        if (-not $MailboxData.AuditSettings.AuditEnabled) {
            $categories["Audit & Compliance"] -= 25
            $securityAnalysis.SecurityIssues += @{
                Category = "Audit & Compliance"
                Issue = "Mailbox auditing is disabled"
                Severity = "High"
                Impact = -25
            }
            $securityAnalysis.Recommendations += "Enable mailbox auditing for security monitoring"
        }
    }
    
    # Data Protection Analysis
    if ($MailboxData.HoldInfo) {
        if (-not $MailboxData.HoldInfo.SingleItemRecoveryEnabled) {
            $categories["Data Protection"] -= 10
            $securityAnalysis.SecurityIssues += @{
                Category = "Data Protection"
                Issue = "Single Item Recovery is disabled"
                Severity = "Low"
                Impact = -10
            }
            $securityAnalysis.Recommendations += "Enable Single Item Recovery for data protection"
        }
        
        if ($MailboxData.HoldInfo.RetentionPolicy -eq "None") {
            $categories["Data Protection"] -= 15
            $securityAnalysis.SecurityIssues += @{
                Category = "Data Protection"
                Issue = "No retention policy applied"
                Severity = "Medium"
                Impact = -15
            }
            $securityAnalysis.Recommendations += "Apply appropriate retention policy"
        }
    }
    
    # Calculate overall score
    $totalScore = 0
    foreach ($category in $categories.Keys) {
        $score = [Math]::Max(0, $categories[$category])
        $securityAnalysis.CategoryScores[$category] = $score
        $totalScore += $score
    }
    $securityAnalysis.OverallScore = [Math]::Round($totalScore / $categories.Count, 0)
    
    # Determine risk level
    if ($securityAnalysis.OverallScore -ge 85) {
        $securityAnalysis.RiskLevel = "Low"
    } elseif ($securityAnalysis.OverallScore -ge 70) {
        $securityAnalysis.RiskLevel = "Medium"
    } elseif ($securityAnalysis.OverallScore -ge 50) {
        $securityAnalysis.RiskLevel = "High"
    } else {
        $securityAnalysis.RiskLevel = "Critical"
    }
    
    Write-Host "    Security Score: $($securityAnalysis.OverallScore)/100 - Risk Level: $($securityAnalysis.RiskLevel)" -ForegroundColor $(
        if ($securityAnalysis.RiskLevel -eq "Low") { "Green" }
        elseif ($securityAnalysis.RiskLevel -eq "Medium") { "Yellow" }
        else { "Red" }
    )
    
    return $securityAnalysis
}

# Enhanced function to get detailed folder analysis with complete retention
function Get-FolderAnalysis {
    param(
        [string]$UserPrincipalName,
        [object]$Mailbox,
        [object]$ArchiveStats,
        [object]$RetentionInfo
    )
    
    Write-Host "  Performing detailed folder analysis..." -ForegroundColor Gray
    
    $folderDetails = @()
    $retentionTags = $RetentionInfo.AllTags
    $retentionTagsByName = $RetentionInfo.TagsByName
    
    try {
        # Get ALL folder statistics (not just a subset)
        $folders = Get-EXOMailboxFolderStatistics -Identity $UserPrincipalName -IncludeOldestAndNewestItems -ErrorAction Stop
        Write-Host "    Processing $($folders.Count) primary mailbox folders..." -ForegroundColor Gray
        
        foreach ($folder in $folders) {
            $folderSize = Convert-SizeToBytes $folder.FolderSize
            
            # Enhanced retention information retrieval
            $retentionTag = "None"
            $retentionAction = "N/A"
            $ageLimitForRetention = "N/A"
            $retentionEnabled = "No"
            
            # Check Delete Policy
            if ($folder.DeletePolicy) {
                $policyString = $folder.DeletePolicy.ToString()
                
                # Try to match by GUID first
                if ($retentionTags.ContainsKey($policyString)) {
                    $tag = $retentionTags[$policyString]
                    $retentionTag = $tag.Name
                    $retentionAction = if ($tag.RetentionAction) { $tag.RetentionAction.ToString() } else { "N/A" }
                    $ageLimitForRetention = if ($tag.AgeLimitForRetention) { "$($tag.AgeLimitForRetention) days" } else { "N/A" }
                    $retentionEnabled = if ($tag.RetentionEnabled) { "Yes" } else { "No" }
                }
                # Try to extract name from GUID (Name) format
                elseif ($policyString -match 'KATEX_INLINE_OPEN(.*?)KATEX_INLINE_CLOSE') {
                    $tagName = $matches[1]
                    if ($retentionTagsByName.ContainsKey($tagName)) {
                        $tag = $retentionTagsByName[$tagName]
                        $retentionTag = $tag.Name
                        $retentionAction = if ($tag.RetentionAction) { $tag.RetentionAction.ToString() } else { "N/A" }
                        $ageLimitForRetention = if ($tag.AgeLimitForRetention) { "$($tag.AgeLimitForRetention) days" } else { "N/A" }
                        $retentionEnabled = if ($tag.RetentionEnabled) { "Yes" } else { "No" }
                    }
                    else {
                        $retentionTag = $tagName
                    }
                }
                # Try direct name match
                elseif ($retentionTagsByName.ContainsKey($policyString)) {
                    $tag = $retentionTagsByName[$policyString]
                    $retentionTag = $tag.Name
                    $retentionAction = if ($tag.RetentionAction) { $tag.RetentionAction.ToString() } else { "N/A" }
                    $ageLimitForRetention = if ($tag.AgeLimitForRetention) { "$($tag.AgeLimitForRetention) days" } else { "N/A" }
                    $retentionEnabled = if ($tag.RetentionEnabled) { "Yes" } else { "No" }
                }
                else {
                    $retentionTag = $policyString
                }
            }
            
            # Check Archive Policy if no Delete Policy
            if ($folder.ArchivePolicy -and $retentionTag -eq "None") {
                $archivePolicyString = $folder.ArchivePolicy.ToString()
                if ($retentionTags.ContainsKey($archivePolicyString)) {
                    $tag = $retentionTags[$archivePolicyString]
                    $retentionTag = "$($tag.Name) (Archive)"
                    $retentionAction = "MoveToArchive"
                    $ageLimitForRetention = if ($tag.AgeLimitForRetention) { "$($tag.AgeLimitForRetention) days" } else { "N/A" }
                    $retentionEnabled = if ($tag.RetentionEnabled) { "Yes" } else { "No" }
                }
            }
            
            $folderDetail = @{
                Location = "Primary Mailbox"
                Name = $folder.Name
                FolderPath = $folder.FolderPath
                FolderType = $folder.FolderType
                FolderSize = Format-ByteSize $folderSize
                FolderSizeBytes = $folderSize
                ItemsInFolder = $folder.ItemsInFolder
                SubfolderCount = if ($folder.SubfolderCount) { $folder.SubfolderCount } else { 0 }
                DeletedItemsInFolder = $folder.DeletedItemsInFolder
                IsRecoverable = $folder.FolderPath -like "*\Recoverable Items*"
                RetentionTag = $retentionTag
                RetentionAction = $retentionAction
                RetentionEnabled = $retentionEnabled
                AgeLimit = $ageLimitForRetention
                OldestItem = if ($folder.OldestItemReceivedDate) { 
                    $folder.OldestItemReceivedDate.ToString("yyyy-MM-dd HH:mm:ss") 
                } else { "N/A" }
                NewestItem = if ($folder.NewestItemReceivedDate) { 
                    $folder.NewestItemReceivedDate.ToString("yyyy-MM-dd HH:mm:ss") 
                } else { "N/A" }
                ItemAge = if ($folder.OldestItemReceivedDate) {
                    [Math]::Round(((Get-Date) - $folder.OldestItemReceivedDate).TotalDays, 0)
                } else { 0 }
            }
            
            $folderDetails += $folderDetail
        }
        
        # Get ALL archive folder statistics if archive is enabled
        if ($Mailbox.ArchiveStatus -eq "Active") {
            try {
                $archiveFolders = Get-EXOMailboxFolderStatistics -Identity $UserPrincipalName -Archive -IncludeOldestAndNewestItems -ErrorAction Stop
                Write-Host "    Processing $($archiveFolders.Count) archive mailbox folders..." -ForegroundColor Gray
                
                foreach ($folder in $archiveFolders) {
                    $folderSize = Convert-SizeToBytes $folder.FolderSize
                    
                    $retentionTag = "None"
                    $retentionAction = "N/A"
                    $ageLimitForRetention = "N/A"
                    $retentionEnabled = "No"
                    
                    if ($folder.DeletePolicy) {
                        $policyString = $folder.DeletePolicy.ToString()
                        
                        if ($retentionTags.ContainsKey($policyString)) {
                            $tag = $retentionTags[$policyString]
                            $retentionTag = $tag.Name
                            $retentionAction = if ($tag.RetentionAction) { $tag.RetentionAction.ToString() } else { "N/A" }
                            $ageLimitForRetention = if ($tag.AgeLimitForRetention) { "$($tag.AgeLimitForRetention) days" } else { "N/A" }
                            $retentionEnabled = if ($tag.RetentionEnabled) { "Yes" } else { "No" }
                        }
                        elseif ($policyString -match 'KATEX_INLINE_OPEN(.*?)KATEX_INLINE_CLOSE') {
                            $tagName = $matches[1]
                            if ($retentionTagsByName.ContainsKey($tagName)) {
                                $tag = $retentionTagsByName[$tagName]
                                $retentionTag = $tag.Name
                                $retentionAction = if ($tag.RetentionAction) { $tag.RetentionAction.ToString() } else { "N/A" }
                                $ageLimitForRetention = if ($tag.AgeLimitForRetention) { "$($tag.AgeLimitForRetention) days" } else { "N/A" }
                                $retentionEnabled = if ($tag.RetentionEnabled) { "Yes" } else { "No" }
                            }
                            else {
                                $retentionTag = $tagName
                            }
                        }
                        elseif ($retentionTagsByName.ContainsKey($policyString)) {
                            $tag = $retentionTagsByName[$policyString]
                            $retentionTag = $tag.Name
                            $retentionAction = if ($tag.RetentionAction) { $tag.RetentionAction.ToString() } else { "N/A" }
                            $ageLimitForRetention = if ($tag.AgeLimitForRetention) { "$($tag.AgeLimitForRetention) days" } else { "N/A" }
                            $retentionEnabled = if ($tag.RetentionEnabled) { "Yes" } else { "No" }
                        }
                        else {
                            $retentionTag = $policyString
                        }
                    }
                    
                    $folderDetail = @{
                        Location = "Archive Mailbox"
                        Name = $folder.Name
                        FolderPath = $folder.FolderPath
                        FolderType = $folder.FolderType
                        FolderSize = Format-ByteSize $folderSize
                        FolderSizeBytes = $folderSize
                        ItemsInFolder = $folder.ItemsInFolder
                        SubfolderCount = if ($folder.SubfolderCount) { $folder.SubfolderCount } else { 0 }
                        DeletedItemsInFolder = $folder.DeletedItemsInFolder
                        IsRecoverable = $folder.FolderPath -like "*\Recoverable Items*"
                        RetentionTag = $retentionTag
                        RetentionAction = $retentionAction
                        RetentionEnabled = $retentionEnabled
                        AgeLimit = $ageLimitForRetention
                        OldestItem = if ($folder.OldestItemReceivedDate) { 
                            $folder.OldestItemReceivedDate.ToString("yyyy-MM-dd HH:mm:ss") 
                        } else { "N/A" }
                        NewestItem = if ($folder.NewestItemReceivedDate) { 
                            $folder.NewestItemReceivedDate.ToString("yyyy-MM-dd HH:mm:ss") 
                        } else { "N/A" }
                        ItemAge = if ($folder.OldestItemReceivedDate) {
                            [Math]::Round(((Get-Date) - $folder.OldestItemReceivedDate).TotalDays, 0)
                        } else { 0 }
                    }
                    
                    $folderDetails += $folderDetail
                }
            }
            catch {
                Write-Host "    Could not retrieve archive folder statistics" -ForegroundColor Yellow
            }
        }
        
        Write-Host "    Total folders processed: $($folderDetails.Count)" -ForegroundColor Gray
    }
    catch {
        Write-Host "    Error during folder analysis: $_" -ForegroundColor Yellow
    }
    
    return $folderDetails
}

# Enhanced main mailbox analysis function
function Get-MailboxAnalysis {
    param(
        [string]$UserPrincipalName
    )
    
    Write-Host "`nAnalyzing mailbox: $UserPrincipalName" -ForegroundColor Cyan
    
    $result = @{
        BasicInfo = $null
        StorageMetrics = @()
        QuotaDetails = @()
        HoldInfo = $null
        LargeFolders = @()
        RecoverableItems = @()
        FolderAnalysis = @()
        InboxRules = $null
        Permissions = $null
        Forwarding = $null
        AuditSettings = $null
        SecurityAnalysis = $null
        RetentionInfo = $null
        Error = $null
    }
    
    try {
        # Get complete retention information first
        $retentionInfo = Get-CompleteRetentionInfo -UserPrincipalName $UserPrincipalName
        $result.RetentionInfo = $retentionInfo
        
        # Get mailbox information
        Write-Host "  Getting mailbox information..." -ForegroundColor Gray
        $exoMailbox = Get-EXOMailbox -Identity $UserPrincipalName -PropertySets All -ErrorAction Stop
        $mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
        $mailboxStats = Get-EXOMailboxStatistics -Identity $UserPrincipalName -ErrorAction Stop
        
        # Get mailbox type
        $mailboxType = "Active"
        if ($mailbox.IsInactiveMailbox) { $mailboxType = "Inactive" }
        elseif ($mailbox.IsSoftDeletedByRemove) { $mailboxType = "Soft-Deleted" }
        
        # Count mailboxes
        $mailboxCount = @{
            Primary = 1
            Archive = if ($mailbox.ArchiveStatus -eq "Active") { 1 } else { 0 }
            Auxiliary = 0
            Total = 1
        }
        
        # Check for auxiliary mailboxes
        if ($mailbox.AuxMailboxParentObjectId) {
            $mailboxCount.Auxiliary++
        }
        $mailboxCount.Total = $mailboxCount.Primary + $mailboxCount.Archive + $mailboxCount.Auxiliary
        
        # Enhanced Basic Information
        $result.BasicInfo = @{
            DisplayName = $mailbox.DisplayName
            EmailAddress = $mailbox.PrimarySmtpAddress
            Alias = $mailbox.Alias
            MailboxType = $mailboxType
            RecipientTypeDetails = $mailbox.RecipientTypeDetails
            LastLogonTime = if ($mailboxStats.LastLogonTime) { $mailboxStats.LastLogonTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
            ItemCount = $mailboxStats.ItemCount
            DeletedItemCount = $mailboxStats.DeletedItemCount
            Database = $mailbox.Database
            WhenCreated = $mailbox.WhenCreated.ToString("yyyy-MM-dd HH:mm:ss")
            HiddenFromAddressListsEnabled = $mailbox.HiddenFromAddressListsEnabled
            EmailAddressPolicyEnabled = $mailbox.EmailAddressPolicyEnabled
            MaxSendSize = if ($mailbox.MaxSendSize) { $mailbox.MaxSendSize.ToString() } else { "Unlimited" }
            MaxReceiveSize = if ($mailbox.MaxReceiveSize) { $mailbox.MaxReceiveSize.ToString() } else { "Unlimited" }
            RetentionPolicy = if ($mailbox.RetentionPolicy) { $mailbox.RetentionPolicy } else { "None" }
            # New fields
            ArchiveEnabled = ($mailbox.ArchiveStatus -eq "Active")
            ArchiveStatus = $mailbox.ArchiveStatus
            ArchiveName = if ($mailbox.ArchiveName) { $mailbox.ArchiveName -join ", " } else { "N/A" }
            ArchiveDatabase = if ($mailbox.ArchiveDatabase) { $mailbox.ArchiveDatabase } else { "N/A" }
            AutoExpandingArchiveEnabled = $mailbox.AutoExpandingArchiveEnabled
            MailboxCount = $mailboxCount
        }
        
        # Storage Metrics for Primary Mailbox
        $primarySize = Convert-SizeToBytes $mailboxStats.TotalItemSize
        $deletedSize = Convert-SizeToBytes $mailboxStats.TotalDeletedItemSize
        
        $storageMetric = @{
            Type = "Primary Mailbox"
            ItemSize = Format-ByteSize $primarySize
            ItemSizeBytes = $primarySize
            DeletedItemSize = Format-ByteSize $deletedSize
            DeletedItemSizeBytes = $deletedSize
            TotalSize = Format-ByteSize ($primarySize + $deletedSize)
            TotalSizeBytes = ($primarySize + $deletedSize)
            ItemCount = $mailboxStats.ItemCount
            DeletedItemCount = $mailboxStats.DeletedItemCount
        }
        $result.StorageMetrics += $storageMetric
        
        # Check for Archive Mailbox
        $archiveStats = $null
        if ($mailbox.ArchiveStatus -ne "None") {
            Write-Host "  Getting archive mailbox information..." -ForegroundColor Gray
            try {
                $archiveStats = Get-EXOMailboxStatistics -Identity $UserPrincipalName -Archive -ErrorAction SilentlyContinue
                if ($archiveStats) {
                    $archiveSize = Convert-SizeToBytes $archiveStats.TotalItemSize
                    $archiveDeletedSize = Convert-SizeToBytes $archiveStats.TotalDeletedItemSize
                    
                    $archiveMetric = @{
                        Type = "Archive Mailbox"
                        ItemSize = Format-ByteSize $archiveSize
                        ItemSizeBytes = $archiveSize
                        DeletedItemSize = Format-ByteSize $archiveDeletedSize
                        DeletedItemSizeBytes = $archiveDeletedSize
                        TotalSize = Format-ByteSize ($archiveSize + $archiveDeletedSize)
                        TotalSizeBytes = ($archiveSize + $archiveDeletedSize)
                        ItemCount = $archiveStats.ItemCount
                        DeletedItemCount = $archiveStats.DeletedItemCount
                    }
                    $result.StorageMetrics += $archiveMetric
                }
            }
            catch {
                Write-Host "    Could not retrieve archive statistics" -ForegroundColor Yellow
            }
        }
        
        # Quota Details
        Write-Host "  Analyzing quota information..." -ForegroundColor Gray
        
        $warningQuota = Convert-SizeToBytes $mailbox.IssueWarningQuota
        $prohibitSendQuota = Convert-SizeToBytes $mailbox.ProhibitSendQuota
        $prohibitSendReceiveQuota = Convert-SizeToBytes $mailbox.ProhibitSendReceiveQuota
        
        $healthStatus = Get-MailboxHealthStatus -UsedSize $primarySize `
            -WarningQuota $warningQuota `
            -ProhibitSendQuota $prohibitSendQuota `
            -ProhibitSendReceiveQuota $prohibitSendReceiveQuota
        
        $quotaDetail = @{
            Type = "Primary Mailbox"
            CurrentSize = Format-ByteSize $primarySize
            WarningQuota = if ($warningQuota -gt 0) { Format-ByteSize $warningQuota } else { "Unlimited" }
            ProhibitSendQuota = if ($prohibitSendQuota -gt 0) { Format-ByteSize $prohibitSendQuota } else { "Unlimited" }
            ProhibitSendReceiveQuota = if ($prohibitSendReceiveQuota -gt 0) { Format-ByteSize $prohibitSendReceiveQuota } else { "Unlimited" }
            UsagePercentage = if ($prohibitSendReceiveQuota -gt 0) { Get-Percentage -Used $primarySize -Limit $prohibitSendReceiveQuota } else { 0 }
            Status = $healthStatus.Status
            StatusColor = $healthStatus.Color
            StatusIcon = $healthStatus.Icon
        }
        $result.QuotaDetails += $quotaDetail
        
        # Hold Information
        Write-Host "  Getting hold and retention information..." -ForegroundColor Gray
        
        $retentionHoldInfo = $mailbox | Select-Object RetentionHoldEnabled, StartDateForRetentionHold, EndDateForRetentionHold
        
        $result.HoldInfo = @{
            LitigationHoldEnabled = $mailbox.LitigationHoldEnabled
            LitigationHoldDate = if ($mailbox.LitigationHoldDate) { $mailbox.LitigationHoldDate.ToString("yyyy-MM-dd") } else { "N/A" }
            LitigationHoldOwner = if ($mailbox.LitigationHoldOwner) { $mailbox.LitigationHoldOwner } else { "N/A" }
            LitigationHoldDuration = if ($mailbox.LitigationHoldDuration) { "$($mailbox.LitigationHoldDuration) days" } else { "Unlimited" }
            InPlaceHolds = if ($mailbox.InPlaceHolds -and $mailbox.InPlaceHolds.Count -gt 0) { $mailbox.InPlaceHolds -join ", " } else { "None" }
            DelayHoldApplied = $mailbox.DelayHoldApplied
            ComplianceTagHoldApplied = $mailbox.ComplianceTagHoldApplied
            RetentionHoldEnabled = $retentionHoldInfo.RetentionHoldEnabled
            RetentionHoldStartDate = if ($retentionHoldInfo.StartDateForRetentionHold) { $retentionHoldInfo.StartDateForRetentionHold.ToString("yyyy-MM-dd") } else { "N/A" }
            RetentionHoldEndDate = if ($retentionHoldInfo.EndDateForRetentionHold) { $retentionHoldInfo.EndDateForRetentionHold.ToString("yyyy-MM-dd") } else { "N/A" }
            ElcProcessingDisabled = $mailbox.ElcProcessingDisabled
            RetentionPolicy = if ($mailbox.RetentionPolicy) { $mailbox.RetentionPolicy } else { "None" }
            RetentionComment = if ($mailbox.RetentionComment) { $mailbox.RetentionComment } else { "None" }
            RetentionUrl = if ($mailbox.RetentionUrl) { $mailbox.RetentionUrl } else { "None" }
            SingleItemRecoveryEnabled = $mailbox.SingleItemRecoveryEnabled
            RetainDeletedItemsFor = $mailbox.RetainDeletedItemsFor.ToString()
            RecoverableItemsQuota = if ($mailbox.RecoverableItemsQuota) { $mailbox.RecoverableItemsQuota.ToString() } else { "Default" }
            RecoverableItemsWarningQuota = if ($mailbox.RecoverableItemsWarningQuota) { $mailbox.RecoverableItemsWarningQuota.ToString() } else { "Default" }
        }
        
        # Get additional analyses
        $result.InboxRules = Get-InboxRulesAnalysis -UserPrincipalName $UserPrincipalName
        $result.Permissions = Get-DelegatesAndPermissions -UserPrincipalName $UserPrincipalName
        $result.Forwarding = Get-ForwardingConfiguration -UserPrincipalName $UserPrincipalName
        $result.AuditSettings = Get-MailboxAuditAnalysis -UserPrincipalName $UserPrincipalName
        
        # Get comprehensive security analysis
        $result.SecurityAnalysis = Get-SecurityAnalysis -MailboxData $result
        
        # Get detailed folder analysis with retention info
        $folderAnalysis = Get-FolderAnalysis -UserPrincipalName $UserPrincipalName -Mailbox $mailbox -ArchiveStats $archiveStats -RetentionInfo $retentionInfo
        $result.FolderAnalysis = $folderAnalysis
        
        # Process for large folders and recoverable items
        $largePrimaryFolders = $folderAnalysis | 
            Where-Object { $_.Location -eq "Primary Mailbox" -and -not $_.IsRecoverable -and $_.FolderSizeBytes -ge 1GB } |
            Sort-Object FolderSizeBytes -Descending |
            Select-Object -First 5
        
        $result.LargeFolders = $largePrimaryFolders
        
        $recoverableFolders = $folderAnalysis | 
            Where-Object { $_.IsRecoverable } |
            Sort-Object FolderSizeBytes -Descending
        
        $result.RecoverableItems = $recoverableFolders
        
    }
    catch {
        $result.Error = $_.Exception.Message
        Write-Host "  Error analyzing mailbox: $_" -ForegroundColor Red
    }
    
    return $result
}

# Enhanced HTML report generation function
function Generate-HTMLReport {
    param(
        [hashtable]$MailboxData,
        [string]$UserAlias
    )
    
    $reportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $reportTitle = "Novartis Exchange Online Mailbox Analysis Report"
    
    # Start building HTML
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$reportTitle - $UserAlias</title>
    <style>
        :root {
            --primary-gradient-start: #3498db;
            --primary-gradient-end: #2c3e50;
            --accent-color: #6b46c1;
            --bg-primary: #f5f7fa;
            --bg-secondary: #ffffff;
            --text-primary: #2c3e50;
            --text-secondary: #7f8c8d;
            --success-color: #27ae60;
            --warning-color: #f39c12;
            --danger-color: #e74c3c;
            --info-color: #3498db;
            --border-radius: 10px;
            --shadow: 0 4px 6px rgba(0,0,0,0.1);
            --shadow-hover: 0 8px 15px rgba(0,0,0,0.2);
        }
        
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            background: linear-gradient(135deg, var(--bg-primary) 0%, #e3e9f3 100%);
            min-height: 100vh;
            color: var(--text-primary);
            line-height: 1.6;
        }
        
        .main-container {
            display: flex;
            max-width: 100%;
            margin: 0;
            min-height: 100vh;
        }
        
        /* Sidebar Navigation */
        .sidebar {
            width: 280px;
            background: linear-gradient(180deg, var(--primary-gradient-start) 0%, var(--primary-gradient-end) 100%);
            color: white;
            position: fixed;
            left: 0;
            top: 0;
            height: 100vh;
            overflow-y: auto;
            box-shadow: 2px 0 10px rgba(0,0,0,0.1);
            z-index: 100;
        }
        
        .logo-section {
            padding: 25px 20px;
            background: rgba(255,255,255,0.1);
            border-bottom: 1px solid rgba(255,255,255,0.2);
            text-align: center;
        }
        
        .logo-text {
            font-size: 1.4rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 2px;
        }
        
        .sidebar-subtitle {
            font-size: 0.75rem;
            opacity: 0.9;
            margin-top: 5px;
        }
        
        .nav-menu {
            padding: 20px 0;
        }
        
        .nav-item {
            display: block;
            padding: 14px 20px;
            color: rgba(255,255,255,0.9);
            text-decoration: none;
            transition: all 0.3s ease;
            border-left: 3px solid transparent;
            cursor: pointer;
            font-size: 0.95rem;
        }
        
        .nav-item:hover {
            background: rgba(255,255,255,0.1);
            border-left-color: #ffffff;
            padding-left: 25px;
        }
        
        .nav-item.active {
            background: rgba(255,255,255,0.2);
            border-left-color: #ffffff;
            font-weight: 600;
        }
        
        .nav-icon {
            margin-right: 10px;
            display: inline-block;
            width: 20px;
        }
        
        /* Content Area */
        .content-wrapper {
            margin-left: 280px;
            width: calc(100% - 280px);
            min-height: 100vh;
            background: white;
        }
        
        .header {
            background: linear-gradient(135deg, var(--primary-gradient-start) 0%, var(--primary-gradient-end) 100%);
            color: white;
            padding: 30px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        
        .header h1 {
            font-size: 2rem;
            font-weight: 700;
            margin-bottom: 8px;
        }
        
        .subtitle {
            font-size: 1.1rem;
            opacity: 0.95;
        }
        
        .report-date {
            font-size: 0.9rem;
            opacity: 0.85;
            margin-top: 5px;
        }
        
        .content {
            padding: 30px;
        }
        
        .section {
            display: none;
            animation: fadeIn 0.5s ease-out;
        }
        
        .section.active {
            display: block;
        }
        
        @keyframes fadeIn {
            from {
                opacity: 0;
                transform: translateY(20px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }
        
        .section-header {
            margin-bottom: 25px;
            padding-bottom: 15px;
            border-bottom: 3px solid var(--primary-gradient-start);
        }
        
        .section-title {
            font-size: 1.8rem;
            color: var(--text-primary);
            font-weight: 600;
        }
        
        /* Info Grid and Cards */
        .info-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 20px;
            margin-bottom: 20px;
        }
        
        .info-card {
            background: white;
            padding: 20px;
            border-radius: var(--border-radius);
            box-shadow: var(--shadow);
            transition: all 0.3s ease;
            border-left: 4px solid var(--primary-gradient-start);
        }
        
        .info-card:hover {
            transform: translateY(-3px);
            box-shadow: var(--shadow-hover);
        }
        
        .info-label {
            font-weight: 600;
            color: var(--text-secondary);
            font-size: 0.85rem;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 8px;
        }
        
        .info-value {
            font-size: 1.2rem;
            color: var(--text-primary);
            font-weight: 500;
        }
        
        /* Security Score Display */
        .security-score-container {
            text-align: center;
            padding: 30px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border-radius: var(--border-radius);
            color: white;
            margin-bottom: 30px;
        }
        
        .security-score {
            font-size: 4rem;
            font-weight: bold;
            margin: 10px 0;
        }
        
        .security-grade {
            font-size: 1.5rem;
            margin: 10px 0;
        }
        
        .risk-level {
            display: inline-block;
            padding: 8px 20px;
            border-radius: 20px;
            font-weight: bold;
            margin-top: 10px;
        }
        
        .risk-low { background: var(--success-color); }
        .risk-medium { background: var(--warning-color); }
        .risk-high { background: var(--danger-color); }
        .risk-critical { background: #8b0000; }
        
        /* Tables */
        .table-container {
            overflow-x: auto;
            border-radius: var(--border-radius);
            box-shadow: var(--shadow);
            margin-top: 20px;
            background: white;
        }
        
        table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
        }
        
        th {
            background: linear-gradient(135deg, var(--primary-gradient-start), var(--primary-gradient-end));
            color: white;
            padding: 14px;
            text-align: left;
            font-weight: 600;
            font-size: 0.9rem;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            position: sticky;
            top: 0;
            z-index: 10;
        }
        
        td {
            padding: 14px;
            border-bottom: 1px solid #e9ecef;
            color: var(--text-primary);
        }
        
        tr:hover td {
            background: #f8f9fa;
        }
        
        /* Status Badges */
        .status-badge {
            display: inline-flex;
            align-items: center;
            padding: 6px 14px;
            border-radius: 20px;
            font-weight: 600;
            font-size: 0.85rem;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .status-enabled {
            background: rgba(39, 174, 96, 0.1);
            color: var(--success-color);
            border: 1px solid var(--success-color);
        }
        
        .status-disabled {
            background: rgba(231, 76, 60, 0.1);
            color: var(--danger-color);
            border: 1px solid var(--danger-color);
        }
        
        /* Severity Badges */
        .severity-low {
            background: rgba(39, 174, 96, 0.1);
            color: var(--success-color);
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 0.85rem;
            font-weight: 600;
        }
        
        .severity-medium {
            background: rgba(243, 156, 18, 0.1);
            color: var(--warning-color);
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 0.85rem;
            font-weight: 600;
        }
        
        .severity-high {
            background: rgba(231, 76, 60, 0.1);
            color: var(--danger-color);
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 0.85rem;
            font-weight: 600;
        }
        
        /* Alert Box */
        .alert {
            padding: 15px 20px;
            border-radius: var(--border-radius);
            margin: 20px 0;
            display: flex;
            align-items: center;
        }
        
        .alert-warning {
            background: rgba(243, 156, 18, 0.1);
            border-left: 4px solid var(--warning-color);
            color: var(--warning-color);
        }
        
        .alert-danger {
            background: rgba(231, 76, 60, 0.1);
            border-left: 4px solid var(--danger-color);
            color: var(--danger-color);
        }
        
        .alert-info {
            background: rgba(52, 152, 219, 0.1);
            border-left: 4px solid var(--info-color);
            color: var(--info-color);
        }
        
        /* Footer */
        .footer {
            background: linear-gradient(180deg, #f8f9fa 0%, #e9ecef 100%);
            color: var(--text-secondary);
            padding: 30px;
            text-align: center;
            border-top: 1px solid #dee2e6;
            position: relative;
            margin-top: 50px;
        }
    </style>
</head>
<body>
    <div class="main-container">
        <!-- Sidebar Navigation -->
        <div class="sidebar">
            <div class="logo-section">
                <div class="logo-text">NOVARTIS</div>
                <div class="sidebar-subtitle">Exchange Online Analysis v3.0</div>
            </div>
            <nav class="nav-menu">
                <a class="nav-item active" onclick="showSection('basic-info', this)">
                    <span class="nav-icon">&#128100;</span>
                    <span>Basic Information</span>
                </a>
                <a class="nav-item" onclick="showSection('security-overview', this)">
                    <span class="nav-icon">&#128274;</span>
                    <span>Security Overview</span>
                </a>
"@

    # Add navigation items conditionally
    if ($MailboxData.InboxRules) {
        $html += @"
                <a class="nav-item" onclick="showSection('inbox-rules', this)">
                    <span class="nav-icon">&#128221;</span>
                    <span>Inbox Rules</span>
                </a>
"@
    }
    
    if ($MailboxData.HoldInfo) {
        $html += @"
                <a class="nav-item" onclick="showSection('hold-info', this)">
                    <span class="nav-icon">&#9888;</span>
                    <span>Hold & Retention</span>
                </a>
"@
    }
    
    if ($MailboxData.Permissions) {
        $html += @"
                <a class="nav-item" onclick="showSection('permissions', this)">
                    <span class="nav-icon">&#128101;</span>
                    <span>Permissions & Delegates</span>
                </a>
"@
    }
    
    if ($MailboxData.Forwarding) {
        $html += @"
                <a class="nav-item" onclick="showSection('forwarding', this)">
                    <span class="nav-icon">&#10132;</span>
                    <span>Forwarding</span>
                </a>
"@
    }
    
    if ($MailboxData.StorageMetrics -and $MailboxData.StorageMetrics.Count -gt 0) {
        $html += @"
                <a class="nav-item" onclick="showSection('storage-metrics', this)">
                    <span class="nav-icon">&#128190;</span>
                    <span>Storage Metrics</span>
                </a>
"@
    }
    
    if ($MailboxData.FolderAnalysis -and $MailboxData.FolderAnalysis.Count -gt 0) {
        $html += @"
                <a class="nav-item" onclick="showSection('folder-analysis', this)">
                    <span class="nav-icon">&#128194;</span>
                    <span>Folder Analysis</span>
                </a>
"@
    }
    
    $html += @"
            </nav>
        </div>
        
        <!-- Main Content Area -->
        <div class="content-wrapper">
            <div class="header">
                <h1>Exchange Online Mailbox Analysis Report</h1>
                <div class="subtitle">User: $UserAlias@novartis.net</div>
                <div class="report-date">Generated: $reportDate</div>
            </div>
            
            <div class="content">
"@

    # Enhanced Basic Information Section
    if ($MailboxData.BasicInfo) {
        $basicInfo = $MailboxData.BasicInfo
        $html += @"
                <div id="basic-info" class="section active">
                    <div class="section-header">
                        <h2 class="section-title">Basic Information</h2>
                    </div>
                    <div class="info-grid">
                        <div class="info-card">
                            <div class="info-label">Display Name</div>
                            <div class="info-value">$($basicInfo.DisplayName)</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Email Address</div>
                            <div class="info-value">$($basicInfo.EmailAddress)</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Mailbox Type</div>
                            <div class="info-value">$($basicInfo.RecipientTypeDetails)</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Last Logon</div>
                            <div class="info-value">$($basicInfo.LastLogonTime)</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Archive Enabled</div>
                            <div class="info-value">
                                <span class="status-badge $(if ($basicInfo.ArchiveEnabled) { 'status-enabled' } else { 'status-disabled' })">
                                    $(if ($basicInfo.ArchiveEnabled) { 'Yes' } else { 'No' })
                                </span>
                            </div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Auto-Expanding Archive</div>
                            <div class="info-value">
                                <span class="status-badge $(if ($basicInfo.AutoExpandingArchiveEnabled) { 'status-enabled' } else { 'status-disabled' })">
                                    $(if ($basicInfo.AutoExpandingArchiveEnabled) { 'Enabled' } else { 'Disabled' })
                                </span>
                            </div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Total Mailboxes</div>
                            <div class="info-value">
                                Primary: $($basicInfo.MailboxCount.Primary)<br/>
                                Archive: $($basicInfo.MailboxCount.Archive)<br/>
                                Auxiliary: $($basicInfo.MailboxCount.Auxiliary)<br/>
                                <strong>Total: $($basicInfo.MailboxCount.Total)</strong>
                            </div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Retention Policy</div>
                            <div class="info-value">$($basicInfo.RetentionPolicy)</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Created Date</div>
                            <div class="info-value">$($basicInfo.WhenCreated)</div>
                        </div>
                    </div>
                </div>
"@
    }
    
    # Security Overview Section
    if ($MailboxData.SecurityAnalysis) {
        $security = $MailboxData.SecurityAnalysis
        $scoreColor = if ($security.OverallScore -ge 85) { "#27ae60" } elseif ($security.OverallScore -ge 70) { "#f39c12" } elseif ($security.OverallScore -ge 50) { "#e74c3c" } else { "#8b0000" }
        
        $html += @"
                <div id="security-overview" class="section">
                    <div class="section-header">
                        <h2 class="section-title">Security Overview</h2>
                    </div>
                    
                    <div class="security-score-container">
                        <div class="security-grade">Security Score</div>
                        <div class="security-score" style="color: $scoreColor;">$($security.OverallScore)/100</div>
                        <div class="risk-level risk-$(($security.RiskLevel).ToLower())">Risk Level: $($security.RiskLevel)</div>
                    </div>
                    
                    <h3 style="margin-top: 30px;">Category Scores</h3>
                    <div class="info-grid">
"@
        foreach ($category in $security.CategoryScores.Keys) {
            $categoryScore = $security.CategoryScores[$category]
            $categoryColor = if ($categoryScore -ge 85) { "#27ae60" } elseif ($categoryScore -ge 70) { "#f39c12" } else { "#e74c3c" }
            $html += @"
                        <div class="info-card">
                            <div class="info-label">$category</div>
                            <div class="info-value" style="color: $categoryColor; font-size: 1.5rem;">$categoryScore/100</div>
                        </div>
"@
        }
        $html += @"
                    </div>
"@
        
        if ($security.SecurityIssues.Count -gt 0) {
            $html += @"
                    <h3 style="margin-top: 30px;">Security Issues Detected</h3>
                    <div class="table-container">
                        <table>
                            <thead>
                                <tr>
                                    <th>Category</th>
                                    <th>Issue</th>
                                    <th>Severity</th>
                                    <th>Impact on Score</th>
                                </tr>
                            </thead>
                            <tbody>
"@
            foreach ($issue in $security.SecurityIssues) {
                $severityClass = "severity-$(($issue.Severity).ToLower())"
                $html += @"
                                <tr>
                                    <td>$($issue.Category)</td>
                                    <td>$($issue.Issue)</td>
                                    <td><span class="$severityClass">$($issue.Severity)</span></td>
                                    <td>$($issue.Impact) points</td>
                                </tr>
"@
            }
            $html += @"
                            </tbody>
                        </table>
                    </div>
"@
        }
        
        if ($security.Recommendations.Count -gt 0) {
            $html += @"
                    <h3 style="margin-top: 30px;">Recommendations</h3>
                    <ul style="margin-left: 20px;">
"@
            foreach ($rec in $security.Recommendations) {
                $html += @"
                        <li style="margin: 10px 0;">$rec</li>
"@
            }
            $html += @"
                    </ul>
"@
        }
        
        $html += @"
                </div>
"@
    }
    
    # Inbox Rules Section
    if ($MailboxData.InboxRules) {
        $rules = $MailboxData.InboxRules
        $html += @"
                <div id="inbox-rules" class="section">
                    <div class="section-header">
                        <h2 class="section-title">Inbox Rules Analysis</h2>
                    </div>
                    
                    <div class="info-grid">
                        <div class="info-card">
                            <div class="info-label">Total Rules</div>
                            <div class="info-value">$($rules.TotalRules)</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Enabled Rules</div>
                            <div class="info-value">$($rules.EnabledRules)</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Forwarding Rules</div>
                            <div class="info-value">$($rules.ForwardingRules)</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Deletion Rules</div>
                            <div class="info-value">$($rules.DeletingRules)</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Moving Rules</div>
                            <div class="info-value">$($rules.MovingRules)</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Suspicious Rules</div>
                            <div class="info-value" style="color: $(if ($rules.SuspiciousRules.Count -gt 0) { '#e74c3c' } else { '#27ae60' });">
                                $($rules.SuspiciousRules.Count)
                            </div>
                        </div>
                    </div>
"@
        
        if ($rules.SuspiciousRules.Count -gt 0) {
            $html += @"
                    <div class="alert alert-danger">
                        <strong>⚠️ Warning:</strong> $($rules.SuspiciousRules.Count) suspicious rules detected that may pose security risks.
                    </div>
                    
                    <h3 style="margin-top: 30px;">Suspicious Rules Details</h3>
                    <div class="table-container">
                        <table>
                            <thead>
                                <tr>
                                    <th>Rule Name</th>
                                    <th>Status</th>
                                    <th>Risk Level</th>
                                    <th>Actions</th>
                                    <th>Suspicious Reasons</th>
                                </tr>
                            </thead>
                            <tbody>
"@
            foreach ($rule in $rules.SuspiciousRules) {
                $riskClass = "severity-$(($rule.RiskLevel).ToLower())"
                $html += @"
                                <tr>
                                    <td><strong>$($rule.Name)</strong></td>
                                    <td>$(if ($rule.Enabled) { '<span class="status-enabled">Enabled</span>' } else { '<span class="status-disabled">Disabled</span>' })</td>
                                    <td><span class="$riskClass">$($rule.RiskLevel)</span></td>
                                    <td>$($rule.Actions -join ', ')</td>
                                    <td>$($rule.SuspiciousReasons -join '; ')</td>
                                </tr>
"@
            }
            $html += @"
                            </tbody>
                        </table>
                    </div>
"@
        }
        
        if ($rules.AllRules.Count -gt 0) {
            $html += @"
                    <h3 style="margin-top: 30px;">All Inbox Rules</h3>
                    <div class="table-container">
                        <table>
                            <thead>
                                <tr>
                                    <th>Rule Name</th>
                                    <th>Enabled</th>
                                    <th>Priority</th>
                                    <th>Conditions</th>
                                    <th>Actions</th>
                                </tr>
                            </thead>
                            <tbody>
"@
            foreach ($rule in $rules.AllRules) {
                $html += @"
                                <tr>
                                    <td><strong>$($rule.Name)</strong></td>
                                    <td>$(if ($rule.Enabled) { '<span class="status-enabled">Yes</span>' } else { '<span class="status-disabled">No</span>' })</td>
                                    <td>$($rule.Priority)</td>
                                    <td>$(if ($rule.Conditions.Count -gt 0) { $rule.Conditions -join '; ' } else { 'No conditions' })</td>
                                    <td>$(if ($rule.Actions.Count -gt 0) { $rule.Actions -join '; ' } else { 'No actions' })</td>
                                </tr>
"@
            }
            $html += @"
                            </tbody>
                        </table>
                    </div>
"@
        }
        
        $html += @"
                </div>
"@
    }
    
    # Enhanced Permissions Section with Folder Permissions
    if ($MailboxData.Permissions) {
        $perms = $MailboxData.Permissions
        $html += @"
                <div id="permissions" class="section">
                    <div class="section-header">
                        <h2 class="section-title">Permissions & Delegates</h2>
                    </div>
                    
                    <div class="info-grid">
                        <div class="info-card">
                            <div class="info-label">Total Delegates</div>
                            <div class="info-value">$($perms.TotalDelegateCount)</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Full Access</div>
                            <div class="info-value">$($perms.FullAccessDelegates.Count)</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Send As</div>
                            <div class="info-value">$($perms.SendAsDelegates.Count)</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Send on Behalf</div>
                            <div class="info-value">$($perms.SendOnBehalfDelegates.Count)</div>
                        </div>
                    </div>
"@
        
        # Full Access Delegates
        if ($perms.FullAccessDelegates.Count -gt 0) {
            $html += @"
                    <h3 style="margin-top: 30px;">Full Access Delegates</h3>
                    <div class="table-container">
                        <table>
                            <thead>
                                <tr>
                                    <th>User</th>
                                    <th>Access Rights</th>
                                    <th>Inherited</th>
                                </tr>
                            </thead>
                            <tbody>
"@
            foreach ($delegate in $perms.FullAccessDelegates) {
                $html += @"
                                <tr>
                                    <td>$($delegate.User)</td>
                                    <td>$($delegate.AccessRights)</td>
                                    <td>$($delegate.IsInherited)</td>
                                </tr>
"@
            }
            $html += @"
                            </tbody>
                        </table>
                    </div>
"@
        }
        
        # Folder Permissions
        if ($perms.FolderPermissions.Count -gt 0) {
            $html += @"
                    <h3 style="margin-top: 30px;">Folder Permissions</h3>
"@
            foreach ($folderName in $perms.FolderPermissions.Keys) {
                $folderPerm = $perms.FolderPermissions[$folderName]
                if ($folderPerm.TotalPermissions -gt 0 -or $folderPerm.DefaultAccess -ne "None" -or $folderPerm.AnonymousAccess -ne "None") {
                    $html += @"
                    <h4 style="margin-top: 20px; color: var(--primary-gradient-start);">$folderName Folder</h4>
                    <div class="info-grid">
                        <div class="info-card">
                            <div class="info-label">Default Access</div>
                            <div class="info-value">$($folderPerm.DefaultAccess)</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Anonymous Access</div>
                            <div class="info-value">$($folderPerm.AnonymousAccess)</div>
                        </div>
                        <div class="info-card">
                            <div class="info-label">Custom Permissions</div>
                            <div class="info-value">$($folderPerm.TotalPermissions)</div>
                        </div>
                    </div>
"@
                    
                    if ($folderPerm.Permissions.Count -gt 0) {
                        $html += @"
                    <div class="table-container">
                        <table>
                            <thead>
                                <tr>
                                    <th>User</th>
                                    <th>Access Rights</th>
                                </tr>
                            </thead>
                            <tbody>
"@
                        foreach ($perm in $folderPerm.Permissions) {
                            $html += @"
                                <tr>
                                    <td>$($perm.User)</td>
                                    <td>$($perm.AccessRights)</td>
                                </tr>
"@
                        }
                        $html += @"
                            </tbody>
                        </table>
                    </div>
"@
                    }
                }
            }
        }
        
        $html += @"
                </div>
"@
    }
    
    # Add other sections (Hold Info, Forwarding, Storage Metrics, Folder Analysis)
    # ... [Previous sections remain the same]
    
    # Close HTML
    $html += @"
            </div>
            
            <!-- Footer -->
            <div class="footer">
                <div>© 2024 Novartis eDLS Team - Exchange Online Mailbox Analysis Tool v3.0</div>
                <div>Generated on $reportDate | Confidential</div>
            </div>
        </div>
    </div>
    
    <script>
        function showSection(sectionId, navElement) {
            // Hide all sections
            var sections = document.querySelectorAll('.section');
            sections.forEach(function(section) {
                section.classList.remove('active');
            });
            
            // Remove active class from all nav items
            var navItems = document.querySelectorAll('.nav-item');
            navItems.forEach(function(item) {
                item.classList.remove('active');
            });
            
            // Show selected section
            var selectedSection = document.getElementById(sectionId);
            if (selectedSection) {
                selectedSection.classList.add('active');
            }
            
            // Add active class to clicked nav item
            navElement.classList.add('active');
        }
    </script>
</body>
</html>
"@
    
    return $html
}

# Export report function
function Export-Report {
    param(
        [string]$HtmlContent,
        [string]$UserAlias
    )
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $fileName = "MailboxAnalysis_${UserAlias}_${timestamp}.html"
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $fullPath = Join-Path -Path $desktopPath -ChildPath $fileName
    
    try {
        $HtmlContent | Out-File -FilePath $fullPath -Encoding UTF8
        Write-Host "Report saved to: $fullPath" -ForegroundColor Green
        return $fullPath
    }
    catch {
        Write-Host "Error saving report: $_" -ForegroundColor Red
        
        # Try alternative location
        $tempPath = Join-Path -Path $env:TEMP -ChildPath $fileName
        try {
            $HtmlContent | Out-File -FilePath $tempPath -Encoding UTF8
            Write-Host "Report saved to temp location: $tempPath" -ForegroundColor Yellow
            return $tempPath
        }
        catch {
            Write-Host "Failed to save report to any location" -ForegroundColor Red
            return $null
        }
    }
}

# Main execution function
function Main {
    Clear-Host
    Show-Banner
    
    # Get user inputs
    Write-Host "Please provide the following information:" -ForegroundColor Cyan
    Write-Host ""
    
    $userAlias = Get-UserInput -Prompt "Enter User Alias (e.g., john.doe)"
    $userPrincipalName = "$userAlias@novartis.net"
    
    Write-Host ""
    Write-Host "Configuration Summary:" -ForegroundColor Yellow
    Write-Host "  Target Mailbox: $userPrincipalName" -ForegroundColor Gray
    Write-Host ""
    
    $confirm = Get-UserInput -Prompt "Do you want to proceed? (Y/N)" -DefaultValue "Y"
    if ($confirm -ne 'Y') {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        return
    }
    
    # Connect to Exchange Online
    if (-not (Connect-ExchangeOnlineSession)) {
        Write-Host "Cannot proceed without Exchange Online connection." -ForegroundColor Red
        return
    }
    
    # Progress tracking
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    # Analyze mailbox
    Write-Host "`nStarting comprehensive mailbox analysis..." -ForegroundColor Yellow
    $mailboxData = Get-MailboxAnalysis -UserPrincipalName $userPrincipalName
    
    if ($mailboxData.Error) {
        Write-Host "`nError occurred during analysis:" -ForegroundColor Red
        Write-Host $mailboxData.Error -ForegroundColor Red
        
        $retry = Get-UserInput -Prompt "Do you want to retry? (Y/N)" -DefaultValue "N"
        if ($retry -eq 'Y') {
            $mailboxData = Get-MailboxAnalysis -UserPrincipalName $userPrincipalName
        }
    }
    
    # Generate HTML report
    Write-Host "`nGenerating enhanced HTML report..." -ForegroundColor Yellow
    $htmlReport = Generate-HTMLReport -MailboxData $mailboxData -UserAlias $userAlias
    
    # Save report
    $reportPath = Export-Report -HtmlContent $htmlReport -UserAlias $userAlias
    
    if ($reportPath) {
        # Open report in browser
        $openReport = Get-UserInput -Prompt "`nDo you want to open the report in your browser? (Y/N)" -DefaultValue "Y"
        if ($openReport -eq 'Y') {
            try {
                Start-Process $reportPath
                Write-Host "Report opened in default browser" -ForegroundColor Green
            }
            catch {
                Write-Host "Could not open report automatically. Please open manually from: $reportPath" -ForegroundColor Yellow
            }
        }
    }
    
    $stopwatch.Stop()
    $duration = $stopwatch.Elapsed
    
    # Disconnect from Exchange Online
    Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Yellow
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "Disconnected from Exchange Online" -ForegroundColor Green
    }
    catch {
        # Silently continue if disconnect fails
    }
    
    # Summary
    Write-Host "`n" -NoNewline
    Write-Host "==========================================================" -ForegroundColor Cyan
    Write-Host "                  Analysis Complete!                      " -ForegroundColor Green
    Write-Host "==========================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Summary:" -ForegroundColor Yellow
    Write-Host "  • User Analyzed: $userAlias" -ForegroundColor Gray
    Write-Host "  • Report Location: $reportPath" -ForegroundColor Gray
    Write-Host "  • Total Time: $($duration.Minutes) minutes $($duration.Seconds) seconds" -ForegroundColor Gray
    Write-Host ""
    
    # Key findings summary
    if ($mailboxData.BasicInfo) {
        Write-Host "Key Findings:" -ForegroundColor Yellow
        Write-Host "  • Archive Status: $(if ($mailboxData.BasicInfo.ArchiveEnabled) { 'Enabled' } else { 'Disabled' })" -ForegroundColor Gray
        Write-Host "  • Auto-Expanding: $(if ($mailboxData.BasicInfo.AutoExpandingArchiveEnabled) { 'Yes' } else { 'No' })" -ForegroundColor Gray
        Write-Host "  • Total Mailboxes: $($mailboxData.BasicInfo.MailboxCount.Total)" -ForegroundColor Gray
        
        if ($mailboxData.InboxRules) {
            Write-Host "  • Inbox Rules: $($mailboxData.InboxRules.TotalRules) total, $($mailboxData.InboxRules.SuspiciousRules.Count) suspicious" -ForegroundColor $(if ($mailboxData.InboxRules.SuspiciousRules.Count -gt 0) { "Yellow" } else { "Gray" })
        }
        
        if ($mailboxData.SecurityAnalysis) {
            $statusColor = switch ($mailboxData.SecurityAnalysis.RiskLevel) {
                "Low" { "Green" }
                "Medium" { "Yellow" }
                "High" { "Red" }
                "Critical" { "Red" }
                default { "Gray" }
            }
            Write-Host "  • Security Score: $($mailboxData.SecurityAnalysis.OverallScore)/100 - Risk: $($mailboxData.SecurityAnalysis.RiskLevel)" -ForegroundColor $statusColor
        }
    }
    
    Write-Host ""
    Write-Host "Thank you for using the Novartis Exchange Online Mailbox Analysis Tool v3.0!" -ForegroundColor Cyan
    Write-Host ""
}

# Run the main function
Main
