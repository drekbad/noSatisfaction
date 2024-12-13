Param(
    [string]$DatabasePath = ".\engagements.json",
    [switch]$Debug
)

# Predefined engagement categories
$EngagementOptions = @(
    "External",
    "Internal",
    "WiFi",
    "FW",
    "Active Directory Assessment",
    "Password Audit",
    "Web App Pentest",
    "Mobile App Pentest",
    "365 / Cloud Pentest"
)

function Count-BusinessDays($StartDate, $EndDate) {
    $start = [DateTime]$StartDate
    $end = [DateTime]$EndDate
    $count = 0
    for ($d = $start; $d -le $end; $d = $d.AddDays(1)) {
        if ($d.DayOfWeek -in "Monday","Tuesday","Wednesday","Thursday","Friday") {
            $count++
        }
    }
    return $count
}

function Validate-Date($prompt) {
    while ($true) {
        $input = Read-Host "$($prompt):"
        $dt = [DateTime]::MinValue
        if ([DateTime]::TryParse($input, [ref]$dt)) {
            if ($dt.Month -ge 1 -and $dt.Month -le 12) {
                return $dt
            } else {
                Write-Host "Month out of range. Please use mm/dd/yy format."
            }
        } else {
            Write-Host "Invalid date format. Please use mm/dd/yy format."
        }
    }
}

function Get-Database {
    if (Test-Path $DatabasePath) {
        $content = Get-Content $DatabasePath -Raw
        if ([string]::IsNullOrWhiteSpace($content)) {
            return [PSCustomObject]@{
                Metadata = [PSCustomObject]@{ TotalRecords = 0 }
                Engagements = @()
            }
        }
        try {
            $data = $content | ConvertFrom-Json
        } catch {
            return [PSCustomObject]@{
                Metadata = [PSCustomObject]@{ TotalRecords = 0 }
                Engagements = @()
            }
        }

        if ($null -eq $data) {
            return [PSCustomObject]@{
                Metadata = [PSCustomObject]@{ TotalRecords = 0 }
                Engagements = @()
            }
        }

        # Check if data is old format (just array)
        if ($data -is [System.Collections.IEnumerable] -and $data -isnot [string] -and
            (-not ($data.PSObject.Properties.Name -contains "Metadata")) -and
            (-not ($data.PSObject.Properties.Name -contains "Engagements"))) {
            # Old format
            $engs = @($data)
            $obj = [PSCustomObject]@{
                Metadata = [PSCustomObject]@{ TotalRecords = $engs.Count }
                Engagements = $engs
            }
            $data = $obj
        }

        # Ensure Metadata and Engagements
        if (-not $data.PSObject.Properties.Name -contains 'Metadata' -or -not $data.PSObject.Properties.Name -contains 'Engagements') {
            $engs = @($data)
            $data = [PSCustomObject]@{
                Metadata = [PSCustomObject]@{ TotalRecords = $engs.Count }
                Engagements = $engs
            }
        }

        $data.Engagements = @($data.Engagements)

        # If TotalRecords is null, fix it
        if ($null -eq $data.Metadata.TotalRecords) {
            $data.Metadata.TotalRecords = $data.Engagements.Count
        }

        if ($data.Engagements.Count -ne $data.Metadata.TotalRecords) {
            Write-Host "Warning: The recorded total ($($data.Metadata.TotalRecords)) does not match the number of engagement records ($($data.Engagements.Count))."
        }

        # Force DomainAdminObtained to boolean if needed
        foreach ($eng in $data.Engagements) {
            if ($eng.PSObject.Properties.Name -notcontains "DomainAdminObtained") {
                $eng | Add-Member -MemberType NoteProperty -Name DomainAdminObtained -Value $false
            } else {
                if ($eng.DomainAdminObtained -isnot [boolean]) {
                    # If it's not boolean, try to parse it as string
                    $strVal = $eng.DomainAdminObtained.ToString().ToLower()
                    if ($strVal -eq 'true') {
                        $eng.DomainAdminObtained = $true
                    } else {
                        $eng.DomainAdminObtained = $false
                    }
                }
            }
        }

        return $data
    } else {
        return [PSCustomObject]@{
            Metadata = [PSCustomObject]@{ TotalRecords = 0 }
            Engagements = @()
        }
    }
}

function Save-Database($data) {
    $json = $data | ConvertTo-Json -Depth 10
    $json | Out-File $DatabasePath -Force
}

function Select-Client {
    Param($prompt = "Select a client")

    $db = Get-Database
    $engs = @($db.Engagements)

    if ($engs.Count -eq 0) {
        Write-Host "No records found."
        return $null
    }

    $clients = @($engs | Sort-Object ClientName)
    $lastMatches = $null

    while ($true) {
        Write-Host "$($prompt): Enter a client name (or type :list to see all):"
        $inputChoice = Read-Host

        if ([string]::IsNullOrWhiteSpace($inputChoice)) {
            return $null
        }

        if ($inputChoice -eq ":list") {
            if ($clients.Count -gt 0) {
                for ($i = 0; $i -lt $clients.Count; $i++) {
                    Write-Host "$($i+1). $($clients[$i].ClientName)"
                }
            } else {
                Write-Host "No clients found."
            }
            $lastMatches = $null
            continue
        }

        if ($inputChoice -match '^\d+$') {
            $index = [int]$inputChoice - 1

            if ($lastMatches -and $lastMatches.Count -gt 0) {
                if ($index -ge 0 -and $index -lt $lastMatches.Count) {
                    return $lastMatches[$index].ClientName
                } else {
                    Write-Host "Invalid number selected."
                }
            } else {
                if ($index -ge 0 -and $index -lt $clients.Count) {
                    return $clients[$index].ClientName
                } else {
                    Write-Host "Invalid number selected."
                }
            }
            continue
        }

        $lcInput = $inputChoice.ToLower()
        $matches = @($clients | Where-Object { $_.ClientName.ToLower().Contains($lcInput) })

        if ($matches.Count -eq 0) {
            Write-Host "No client found matching '$inputChoice'."
            $lastMatches = $null
        } elseif ($matches.Count -eq 1) {
            return $matches[0].ClientName
        } else {
            Write-Host "Multiple matches found:"
            for ($i = 0; $i -lt $matches.Count; $i++) {
                Write-Host "$($i+1). $($matches[$i].ClientName)"
            }
            $lastMatches = $matches
        }
    }
}

function Get-EngagementType {
    Write-Host "Select all applicable categories (comma separated):"
    for ($i = 0; $i -lt $EngagementOptions.Count; $i++) {
        Write-Host "$($i+1). $($EngagementOptions[$i])"
    }
    Write-Host "$($EngagementOptions.Count + 1). Create a new custom category"

    $selection = Read-Host "Enter numbers (e.g. 1,2,5)"

    $categories = @()
    foreach ($choice in $selection.Split(',')) {
        $trimmed = $choice.Trim()
        if ($trimmed -match '^\d+$') {
            $index = [int]$trimmed
            if ($index -eq ($EngagementOptions.Count + 1)) {
                $newCat = Read-Host "Enter new category name"
                if ($newCat) {
                    $categories += $newCat.Trim()
                }
            } elseif ($index -ge 1 -and $index -le $EngagementOptions.Count) {
                $categories += $EngagementOptions[$index - 1].Trim()
            } else {
                Write-Host "Invalid selection: $index"
            }
        } else {
            Write-Host "Invalid input: $trimmed"
        }
    }

    $categories = @($categories)
    return $categories
}

function Parse-Rating($ratingInput) {
    [double]$rating = $ratingInput
    if ($ratingInput -match "\.") {
        $decimalPart = $ratingInput.Split('.')[1]
        if ($decimalPart.Length -gt 1) {
            $rating = [math]::Round($rating, 1)
        }
    }
    return $rating
}

function Add-Record {
    $db = Get-Database

    Write-Host "Adding a new engagement record..."

    $ClientName = Read-Host "Enter the Client Name"
    $EngagementType = Get-EngagementType
    $EngagementType = @($EngagementType)
    $Date = (Get-Date).ToString("yyyy-MM-dd")

    $domainInput = Read-Host "Was Domain Admin Obtained? (y/n)"
    $DomainAdminObtained = ($domainInput -eq 'y')
    $DomainAdminObtained = [bool]$DomainAdminObtained

    $NumberOfUsers = [int](Read-Host "Number of Users in the Forest")
    $NumberOfLiveHosts = [int](Read-Host "Number of Live Hosts Discovered")
    $CompromisedUsersCount = [int](Read-Host "Number of Users Compromised")
    $SensitiveDataObtained = ((Read-Host "Was highly sensitive data obtained? (y/n)") -eq 'y')

    $ProjectedHours = [int](Read-Host "Projected Hours for the project")
    $HoursSpent = [int](Read-Host "Hours Spent on closeout")
    $hoursDiff = $HoursSpent - $ProjectedHours

    $startDate = Validate-Date "Enter the Start Date (mm/dd/yy)"
    $endDate = Validate-Date "Enter the End Date (mm/dd/yy)"
    $businessDays = Count-BusinessDays $startDate $endDate

    Write-Host "Gathering client feedback."
    $CustomFeedback = Read-Host "Use standard 5 questions? (y/n)"
    $ClientFeedbackQuestions = @()

    if ($CustomFeedback -eq 'y') {
        $questions = @(
            "How satisfied were you with the engagement scope?",
            "Did we meet your expectations?",
            "Would you recommend our services?",
            "How did our communication feel throughout the project?",
            "Any areas we could improve?"
        )
        foreach ($q in $questions) {
            $a = Read-Host $q
            $ClientFeedbackQuestions += [PSCustomObject]@{
                Question = $q
                Answer   = $a
            }
        }
    } else {
        for ($i = 1; $i -le 5; $i++) {
            $q = Read-Host "Enter Question #$i"
            $a = Read-Host "Answer"
            $ClientFeedbackQuestions += [PSCustomObject]@{
                Question = $q
                Answer   = $a
            }
        }
    }

    $ratingInput = Read-Host "Please provide the client rating out of 5 (e.g., 4.5, 4.9, 4.85)"
    $ClientRating = 0
    if ([double]::TryParse($ratingInput, [ref]$null)) {
        [double]$rating = [double]$ratingInput
        $dotIndex = $ratingInput.IndexOf('.')
        if ($dotIndex -gt -1) {
            $decimalPart = $ratingInput.Substring($dotIndex+1)
            if ($decimalPart.Length -gt 1) {
                $rating = [math]::Round($rating, 1)
            }
        }
        $ClientRating = $rating
    } else {
        Write-Host "Invalid rating input, defaulting to 0"
    }

    $record = [PSCustomObject]@{
        ClientName              = $ClientName
        EngagementType          = $EngagementType
        Date                    = $Date
        DomainAdminObtained     = $DomainAdminObtained
        NumberOfUsers           = $NumberOfUsers
        NumberOfLiveHosts       = $NumberOfLiveHosts
        CompromisedUsersCount   = $CompromisedUsersCount
        SensitiveDataObtained   = $SensitiveDataObtained
        ClientFeedbackQuestions = $ClientFeedbackQuestions
        ClientRating            = $ClientRating
        ProjectedHours          = $ProjectedHours
        HoursSpent              = $HoursSpent
        HoursDifference         = $hoursDiff
        StartDate               = $startDate.ToString("MM/dd/yy")
        EndDate                 = $endDate.ToString("MM/dd/yy")
        BusinessDaysCount       = $businessDays
    }

    $db.Engagements = @($db.Engagements)
    $db.Engagements += $record
    $db.Metadata.TotalRecords = $db.Engagements.Count
    Save-Database $db

    $db = Get-Database
    Write-Host "Record added successfully."
}

function ModifyOrDelete-Record {
    $selectedClient = Select-Client "Select a client to modify/delete"
    if (-not $selectedClient) {
        return
    }

    $db = Get-Database
    $db.Engagements = @($db.Engagements)
    $record = $db.Engagements | Where-Object { $_.ClientName -eq $selectedClient }

    if (-not $record) {
        Write-Host "No record found."
        return
    }

    Write-Host "Selected Client: $($record.ClientName)"

    Write-Host "What would you like to do?"
    Write-Host "1. Modify fields"
    Write-Host "2. Delete the record"
    $action = Read-Host "Choose an option (1 or 2)"

    if ($action -eq '2') {
        $db.Engagements = $db.Engagements | Where-Object { $_.ClientName -ne $selectedClient }
        $db.Metadata.TotalRecords = $db.Engagements.Count
        Save-Database $db
        Write-Host "Record deleted."
        return
    }

    Write-Host "Leave field blank if you do not wish to change it."

    $newClientName = Read-Host "New Client Name (current: $($record.ClientName))"
    if ($newClientName) { $record.ClientName = $newClientName }

    Write-Host "Modify Engagement Type?"
    $changeType = Read-Host "y/n"
    if ($changeType -eq 'y') {
        $et = Get-EngagementType
        $et = @($et)
        $record.EngagementType = $et | ForEach-Object { $_.Trim() }
    }

    $newDomainAdmin = Read-Host "Domain Admin Obtained? (y/n) (current: $($record.DomainAdminObtained))"
    if ($newDomainAdmin -eq 'y') { $record.DomainAdminObtained = $true }
    elseif ($newDomainAdmin -eq 'n') { $record.DomainAdminObtained = $false }

    $newUsers = Read-Host "Number of Users (current: $($record.NumberOfUsers))"
    if ($newUsers) { $record.NumberOfUsers = [int]$newUsers }

    $newHosts = Read-Host "Number of Live Hosts (current: $($record.NumberOfLiveHosts))"
    if ($newHosts) { $record.NumberOfLiveHosts = [int]$newHosts }

    $newCompromised = Read-Host "Compromised Users Count (current: $($record.CompromisedUsersCount))"
    if ($newCompromised) { $record.CompromisedUsersCount = [int]$newCompromised }

    $newSensitiveData = Read-Host "Sensitive Data Obtained? (y/n) (current: $($record.SensitiveDataObtained))"
    if ($newSensitiveData -eq 'y') { $record.SensitiveDataObtained = $true }
    elseif ($newSensitiveData -eq 'n') { $record.SensitiveDataObtained = $false }

    $newRating = Read-Host "Client Rating out of 5 (current: $($record.ClientRating))"
    if ($newRating) { $record.ClientRating = Parse-Rating($newRating) }

    $newProjected = Read-Host "Projected Hours (current: $($record.ProjectedHours))"
    if ($newProjected) { $record.ProjectedHours = [int]$newProjected }

    $newSpent = Read-Host "Hours Spent (current: $($record.HoursSpent))"
    if ($newSpent) { $record.HoursSpent = [int]$newSpent }
    $record.HoursDifference = $record.HoursSpent - $record.ProjectedHours

    $changeDates = Read-Host "Change Start/End Dates? (y/n)"
    if ($changeDates -eq 'y') {
        $newStart = Validate-Date "Enter the new Start Date (mm/dd/yy) (current: $($record.StartDate))"
        $newEnd = Validate-Date "Enter the new End Date (mm/dd/yy) (current: $($record.EndDate))"
        $record.StartDate = $newStart.ToString("MM/dd/yy")
        $record.EndDate = $newEnd.ToString("MM/dd/yy")
        $record.BusinessDaysCount = Count-BusinessDays $newStart $newEnd
    }

    $db.Engagements = $db.Engagements | Where-Object { $_.ClientName -ne $selectedClient }
    $db.Engagements += $record
    $db.Metadata.TotalRecords = $db.Engagements.Count
    Save-Database $db

    Write-Host "Record updated successfully."
}

function View-Record {
    $selectedClient = Select-Client "Select a client to view"
    if (-not $selectedClient) {
        return
    }

    $db = Get-Database
    $record = $db.Engagements | Where-Object { $_.ClientName -eq $selectedClient }

    if (-not $record) {
        Write-Host "No record found."
        return
    }

    $record | ConvertTo-Json -Depth 10 | Out-Host
}

function Query-Data {
    $db = Get-Database
    $db.Engagements = @($db.Engagements)
    if ($db.Engagements.Count -eq 0) {
        Write-Host "No records found."
        return
    }

    Write-Host "Available Queries:"
    Write-Host "1. Complete Metrics Report"
    Write-Host "2. Show average metrics (e.g. average number of compromised users)."
    Write-Host "3. List all clients and ratings."
    $choice = Read-Host "Select a query option (1-3)"

    switch ($choice) {
        1 {
            $internal = @($db.Engagements | Where-Object {
                (,@($_.EngagementType) | ForEach-Object { $_.ToLower().Trim() }) -contains "internal"
            })
            $intCount = [int]$internal.Count

            # Use -eq $true since we forced boolean conversion
            $daSet = $internal | Where-Object { $_.DomainAdminObtained -eq $true }
            $daCount = $daSet.Count

            if ($intCount -eq 0) {
                $internalDAValue = "No internal engagements"
            } else {
                $percentage = [math]::Round(($daCount / $intCount)*100,2)
                $internalDAValue = "$percentage% ($daCount/$intCount)"
            }

            $hasComp = $db.Engagements | Where-Object { $_.NumberOfUsers -gt 0 -and $_.CompromisedUsersCount -ge 0 }
            if ($hasComp.Count -eq 0) {
                $avgPercentCompromisedValue = "N/A"
            } else {
                $percentages = $hasComp | ForEach-Object {
                    if ($_.NumberOfUsers -gt 0) {
                        (($_.CompromisedUsersCount / $_.NumberOfUsers)*100)
                    }
                }
                $percentages = $percentages | Where-Object { $_ -ne $null }
                if ($percentages.Count -gt 0) {
                    $avgPercentCompromised = [math]::Round(($percentages | Measure-Object -Average).Average,2)
                    $avgPercentCompromisedValue = "$avgPercentCompromised%"
                } else {
                    $avgPercentCompromisedValue = "N/A"
                }
            }

            $hasHours = $db.Engagements | Where-Object { $_.ProjectedHours -ge 0 -and $_.HoursSpent -ge 0 }
            if ($hasHours.Count -eq 0) {
                $timeAllotmentValue = "N/A"
            } else {
                $differences = $hasHours | ForEach-Object { $_.HoursSpent - $_.ProjectedHours }
                $avgDiff = [math]::Round(($differences | Measure-Object -Average).Average,2)
                if ($avgDiff -gt 0) {
                    $timeAllotmentValue = "Over by $avgDiff h"
                } elseif ($avgDiff -lt 0) {
                    $timeAllotmentValue = "Under by $([math]::Abs($avgDiff)) h"
                } else {
                    $timeAllotmentValue = "Exactly on target!"
                }
            }

            $hasDates = $db.Engagements | Where-Object { $_.BusinessDaysCount -gt 0 }
            if ($hasDates.Count -eq 0) {
                $avgProjectLengthValue = "N/A"
            } else {
                $avgDays = [math]::Round(($hasDates.BusinessDaysCount | Measure-Object -Average).Average,0)
                $weeks = [math]::Floor($avgDays / 5)
                $days = $avgDays % 5
                $avgProjectLengthValue = "$avgDays days ($weeks wks, $days days)"
            }

            $avgRating = ($db.Engagements.ClientRating | Measure-Object -Average).Average
            if ($avgRating -eq $null) {
                $avgRatingValue = "N/A"
            } else {
                $avgRating = [math]::Round($avgRating,2)
                $avgRatingValue = "$avgRating / 5"
            }

            $metrics = @(
                [PSCustomObject]@{Label="Internal pentests with DA result:"; Value=$internalDAValue},
                [PSCustomObject]@{Label="Avg. compromised users %:"; Value=$avgPercentCompromisedValue},
                [PSCustomObject]@{Label="Average Time Allotment Accuracy:"; Value=$timeAllotmentValue},
                [PSCustomObject]@{Label="Average Project Length:"; Value=$avgProjectLengthValue},
                [PSCustomObject]@{Label="Average Client Rating:"; Value=$avgRatingValue}
            )

            $maxLabelLength = ($metrics.Label | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum

            Write-Host "`nComplete Metrics Report:`n"
            foreach ($m in $metrics) {
                $paddedLabel = $m.Label.PadRight($maxLabelLength)
                Write-Host "$paddedLabel  $($m.Value)"
            }

            # If single internal engagement scenario and debug requested, print more details
            if ($Debug -and $intCount -eq 1) {
                Write-Host "`nDEBUG: Single internal engagement scenario detected."
                Write-Host "DEBUG: intCount=$intCount daCount=$daCount"
                Write-Host "DEBUG: internal: $($internal | ConvertTo-Json)"

                if ($internal.Count -eq 1) {
                    Write-Host "DEBUG: DomainAdminObtained=$($internal[0].DomainAdminObtained)"
                    Write-Host "DEBUG: DomainAdminObtained Type: $($internal[0].DomainAdminObtained.GetType().FullName)"
                }

                Write-Host "DEBUG: daSet count: $($daSet.Count)"
                Write-Host "DEBUG: daSet: $($daSet | ConvertTo-Json)"
            }

            # Export to CSV for formatting in Excel
            $reportFile = "metrics_report.csv"
            $metrics | Export-Csv -Path $reportFile -NoTypeInformation -Encoding UTF8
            Write-Host "`nMetrics report written to $reportFile"
        }

        2 {
            # Show average metrics
            $avgUsers = [math]::Round(($db.Engagements.NumberOfUsers | Measure-Object -Average).Average,2)
            $avgHosts = [math]::Round(($db.Engagements.NumberOfLiveHosts | Measure-Object -Average).Average,2)
            $avgCompromised = [math]::Round(($db.Engagements.CompromisedUsersCount | Measure-Object -Average).Average,2)
            Write-Host "Average Number of Users: $avgUsers"
            Write-Host "Average Number of Live Hosts: $avgHosts"
            Write-Host "Average Compromised Users: $avgCompromised"
        }

        3 {
            # List all clients and ratings
            $sorted = @($db.Engagements | Sort-Object ClientName)
            foreach ($s in $sorted) {
                Write-Host "Client: $($s.ClientName), Rating: $($s.ClientRating)/5"
            }
        }

        default {
            Write-Host "Invalid choice."
        }
    }
}

function Main-Menu {
    Clear-Host
    Write-Host "=== Engagement Tracking ==="
    Write-Host "1. Add a New Record"
    Write-Host "2. Modify/Delete a Record"
    Write-Host "3. View a Client Record"
    Write-Host "4. Query the Data"
    Write-Host "5. Exit"
    $choice = Read-Host "Select an option (1-5)"

    switch ($choice) {
        1 { Add-Record }
        2 { ModifyOrDelete-Record }
        3 { View-Record }
        4 { Query-Data }
        5 { exit }
        default { Write-Host "Invalid choice. Try again." }
    }
}

while ($true) {
    Main-Menu
    Write-Host "Press Enter to continue..."
    [void][System.Console]::ReadKey()
}
