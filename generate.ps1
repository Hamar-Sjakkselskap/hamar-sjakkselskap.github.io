$files = dir .\resultater

$tournaments = @(
    @{Name = "Klubbmesterskapet 2020"; Group = "Gruppe A"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Klubbmesterskapet2020-HamarSjakkselskap&group=A"; Active = $true}
    @{Name = "Klubbmesterskapet 2020"; Group = "Gruppe B"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Klubbmesterskapet2020-HamarSjakkselskap&group=B"; Active = $true}
    @{Name = "Seriespill 3.div"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Ostlandsserien201920204div-NorgesSjakkforbund&group=3.%20div%20B"; Active = $true}
    @{Name = "Seriespill 4.div"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Ostlandsserien201920204div-NorgesSjakkforbund&group=4.%20div%20A"; Active = $true}

    @{Name = "Hamarturneringen 2019"; Group = "Gruppe A"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Hamarturneringen2019-HamarSjakkselskap&group=A"; Active = $false}
    @{Name = "Hamarturneringen 2019"; Group = "Gruppe B"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Hamarturneringen2019-HamarSjakkselskap&group=B"; Active = $false}
)


# Get stats from string
function Get-StatsFromContent([string]$content) {
    $stats = @{
        Winner = $null
        Participants = 0
        Title = $null
        BestRatingProgress = $null
    }

    $html = New-Object -ComObject HTMLFile
    $html.IHTMLDocument2_write($content)
    $tsTabell = $html.body.getElementsByClassName("ts_tabell")
    if($tsTabell.length -gt 0) {
        $stats.Participants = $tsTabell.length
        $stats.Winner = $tsTabell[0].childNodes[1].innerText
    } 

    $xl45 = $html.body.getElementsByClassName("xl45")
    if($xl45.length -gt 0) {
        $stats.Winner = $xl45[0].innerText
        $stats.Participants = $xl45.length
    }

    $h1s = $html.body.getElementsByTagName("h1")
    if($h1s.length -gt 0) {
        $stats.Title = $h1s[0].innerText
    }

    # Determine who has best rating progress
    $trElements = $html.body.getElementsByTagName("tr")
    $bestRatingProgress = -999
    $trElements | ? className -eq 'ts_tabell' | foreach {
        $lastTd = $_.getElementsByTagName("td")  | select -last 1
        $allTds = $_.getElementsByTagName("td") 
        $txt = $lastTd.innerText.Trim()
        if($txt -match "^[0-9]+ \([+-]?[0-9]+\)$") {
            $prog = [int] $txt.Split("(")[1].Replace(")","")
            if($prog -gt $bestRatingProgress) {
                $bestRatingProgress = $prog
                $stats.BestRatingProgress = $allTds[1].InnerText + " " + $txt.Split(" ")[1]
            }
        }
    }


    return $stats
}

# Get stats from file
function Get-StatsFromFile([string]$file) {
    Write-verbose "Getting stats for $file" -Verbose
    $content = Get-Content -Raw $file -Encoding Default
    Get-StatsFromContent $content
}

# Get stats from turneringsservice
function Get-StatsFromTurneringsservice([string]$uri) {
    Write-verbose "Getting stats for $uri" -Verbose
    $content = Invoke-WebRequest -Uri $uri -UseBasicParsing -ErrorAction Stop
    Get-StatsFromContent  $content.Content
}

# Parse all files into objects with statistics
$objects = $files | foreach {
    $obj = @{
        Type = "Ukjent"
        File = "resultater/$($_.Name)"
        Date = [Datetime]::ParseExact(([Regex]::Matches($_.BaseName, "[0-9]{6}") | Select -First 1 | Select -exp Value),"yyMMdd", $null)
        Group = $null
        Stats = $null
    }

    if($_.BaseName -like "*-*") {
        $obj.Group = "Gruppe " + ($_.BaseName -split "-" | Select-Object -Last 1)
    }

    if($_.Name -match "Hu[0-9]{6}(-[A-C1-9])?.htm") {
        $obj.Type = "Hurtigsjakk"
        $obj.Stats = Get-StatsFromFile $obj.File
    } elseif($_.Name -match "Ly[0-9]{6}(-[A-C1-9])?.htm") {
        $obj.Type = "Lynsjakk"
        $obj.Stats = Get-StatsFromFile $obj.File
    } elseif($_.Name -match "Hc[0-9]{6}(-[A-C1-9])?.htm") {
        $obj.Type = "Hurtigsjakk med tidshandicap"
        $obj.Stats = Get-StatsFromFile $obj.File
    }

    [PSCustomObject] $obj
} | sort Date, Group



$text = @() # "# Resultater - Hamar sjakkselskap")


$text += $tournaments | ? active | Foreach -Begin {
    ""
    "# Aktive turneringer i turneringsservice"
    ""
    "| Navn | Gruppe | Deltagere | Leder |"
    "|-|-|-|-|"
} -Process {
    $stats = Get-StatsFromTurneringsservice -uri $_.url
    "|[{0}]({1})|{2}|{3}|{4}|" -f $_.Name, $_.Url, $_.Group, $Stats.Participants, $Stats.Winner
} -End {
    "|[Arkiv](turneringer.md)||||"
}



# Generate summary file
$text += $objects | Group Type | ? Name -in "Lynsjakk","Hurtigsjakk" | Foreach {
    ""
    "# $($_.Name)"
    ""
    "| Dato | Gruppe | Deltagere | Vinner | Beste ratingfremgang |"
    "|-|-|-|-|-|"

    $_.Group | Sort @{Expression = {$_.Date}; Ascending = $false}, Group | Select -First 5 | Foreach {
        "|[{0}]({1})|{2}|{3}|{4}|{5}|" -f $_.Date.ToString("yyyy-MM-dd"), $_.File, $_.Group, $_.Stats.Participants, $_.Stats.Winner, $_.Stats.BestRatingProgress
    }

    "|[Alle]({0}.md)||||" -f $_.Name
}

$text += ""
$text += "[Alle resultater](arkiv.md)"
$text | Set-Content "index.md" -Encoding UTF8

# Generate per type files
$objects | Group Type | Foreach {
    $text = @(
        "# $($_.Name)"
        ""
        "| Dato | Gruppe | Deltakere | Vinner | Beste ratingfremgang |"
        "|-|-|-|-|-|"
    )

    $text += $_.Group | sort @{Expression = {$_.Date}; Ascending = $false}, Group | Foreach {
        "|[{0}]({1})|{2}|{3}|{4}|{5}|" -f $_.Date.ToString("yyyy-MM-dd"), $_.File, $_.Group, $_.Stats.Participants, $_.Stats.Winner, $_.Stats.BestRatingProgress
    }

    $text | Set-Content "$($_.Name).md" -Encoding UTF8
}

$tournaments | Foreach -Begin {
    $text = @(
        "# Turneringer"
        ""
        "| Navn | Gruppe | Deltagere | Leder |"
        "|-|-|-|-|"
    )
} -Process {
    if($_.active) {
        $stats = Get-StatsFromTurneringsservice -uri $_.url
    } else {
        $stats = @{}
    }
    $text += "|[{0}]({1})|{2}|{3}|{4}|" -f $_.Name, $_.Url, $_.Group, $Stats.Participants, $Stats.Winner
} -End {
    $text | Set-Content "turneringer.md" -Encoding UTF8
}


$objects | 
    sort @{Expression = {$_.Date}; Ascending = $false}, Group |
    Foreach-Object -Begin {
        $text = @(
            "# Alle turneringer"
            ""
            "| Dato | Turnering | Deltagere | Vinner | Beste ratingfremgang |"
            "|-|-|-|-|-|"
        )
    } -Process {
        if($_.Type -ne "Ukjent") {
            $Name = $_.Type
            if($_.Group) {
                $Name += " - " + $_.Group
            }
        } else {
            $Name = $_.Stats.Name
        }
        $text += "|[{1}]({0})|[{2}]({0})|{3}|{4}|{5}|" -f $_.File, $_.Date.ToString("yyyy-MM-dd"), $Name , $_.Stats.Participants, $_.Stats.Winner, $_.Stats.BestRatingProgress
    } -End {
        $text | Set-Content "arkiv.md" -Encoding UTF8
    }
