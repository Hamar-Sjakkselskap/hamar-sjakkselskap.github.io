[CmdletBinding()]

Param()

$Error.Clear()
Push-Location $PSScriptRoot
Install-Module PowerHTML  -Scope CurrentUser -Force -Confirm:$false

# git pull | Out-Null

$files = Get-ChildItem .\resultater

$tournaments = @(
    @{Name = "Hamarturneringen 2025"; Group = "Gruppe A"; Url = "https://tournamentservice.com/standings.aspx?TID=Hamarturneringen2025-HamarSjakkselskap&group=A"; Active = $true}
    @{Name = "Hamarturneringen 2025"; Group = "Gruppe B"; Url = "https://tournamentservice.com/standings.aspx?TID=Hamarturneringen2025-HamarSjakkselskap&group=B"; Active = $true} 

    @{Name = "Klubbmesterskapet 2025"; Group = "Gruppe A"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Klubbmesterskapet2025-HamarSjakkselskap&group=A"; Active = $false}
    @{Name = "Klubbmesterskapet 2025"; Group = "Gruppe B"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Klubbmesterskapet2025-HamarSjakkselskap&group=B"; Active = $false} 

    @{Name = "Julelyn 2024"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Julelynsjakk2024-HamarSjakkselskap"; Active = $false}

    @{Name = "Hamarturneringen 2024"; Group = "Gruppe A"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Hamarturneringen2024-HamarSjakkselskap&group=A"; Active = $false}
    @{Name = "Hamarturneringen 2024"; Group = "Gruppe B"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Hamarturneringen2024-HamarSjakkselskap&group=B"; Active = $false} 

    @{Name = "Klubbmesterskapet 2024"; Group = "Gruppe A"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Klubbmesterskapet2024-HamarSjakkselskap&group=A"; Active = $false}
    @{Name = "Klubbmesterskapet 2024"; Group = "Gruppe B"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Klubbmesterskapet2024-HamarSjakkselskap&group=B"; Active = $false} 
    
    @{Name = "Julelyn 2023"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Julelynsjakk2023-HamarSjakkselskap"; Active = $false}

    @{Name = "Klubbmesterskapet 2023"; Group = "Gruppe A"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Klubbmesterskapet2023-HamarSjakkselskap&group=A"; Active = $false}
    @{Name = "Klubbmesterskapet 2023"; Group = "Gruppe B"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Klubbmesterskapet2023-HamarSjakkselskap&group=B"; Active = $false}

    @{Name = "Hamarturneringen 2023"; Group = "Gruppe A"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Hamarturneringen2023-HamarSjakkselskap&group=A"; Active = $false}
    @{Name = "Hamarturneringen 2023"; Group = "Gruppe B"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Hamarturneringen2023-HamarSjakkselskap&group=B"; Active = $false}

    @{Name = "Klubbmesterskapet 2022"; Group = "Gruppe A"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Klubbmesterskapet2022-HamarSjakkselskap&group=A"; Active = $false}
    @{Name = "Klubbmesterskapet 2022"; Group = "Gruppe B"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Klubbmesterskapet2022-HamarSjakkselskap&group=B"; Active = $false}

    @{Name = "Hamarturneringen 2021"; Group = "Gruppe A"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Hamarturneringen2021-HamarSjakkselskap&group=A"; Active = $false}
    @{Name = "Hamarturneringen 2021"; Group = "Gruppe B"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Hamarturneringen2021-HamarSjakkselskap&group=B"; Active = $false}

    @{Name = "Klubbmesterskapet 2020"; Group = "Gruppe A"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Klubbmesterskapet2020-HamarSjakkselskap&group=A"; Active = $false}
    @{Name = "Klubbmesterskapet 2020"; Group = "Gruppe B"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Klubbmesterskapet2020-HamarSjakkselskap&group=B"; Active = $false}
    @{Name = "Seriespill 3.div"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Ostlandsserien201920204div-NorgesSjakkforbund&group=3.%20div%20B"; Active = $false}
    @{Name = "Seriespill 4.div"; Url = "http://turneringsservice.sjakklubb.no/standings.aspx?TID=Ostlandsserien201920204div-NorgesSjakkforbund&group=4.%20div%20A"; Active = $false}

    @{Name = "Hamar Vinterlyn 2020"; Url="http://turneringsservice.sjakklubb.no/standings.aspx?TID=HamarVinterlyn2020-HamarSjakkselskap&fbclid=IwAR3vN8uLxFtBOng25gkQqFaRQIip1GKcLZ3tcHgjw9P14qMbxykzTR8WTk0"; Active = $false}
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
    
    # Read as html object
    $html = ConvertFrom-Html -Content $content

    $tableStandings = $html.SelectNodes(".//*[contains(@class,'table-standings')]")
    if($tableStandings) {
        $tableRows = $tableStandings.SelectNodes(".//tr[not(contains(@class,'ts_top'))]")
        $stats.Participants = $tableRows | Measure-Object | Select-Object -ExpandProperty Count
        $stats.Winner = $tableRows[0].SelectNodes(".//td")[1].InnerText
    } 

    $tsTabell = $html.SelectNodes(".//*[contains(@class,'ts_tabell')]")
    if($tsTabell) {
        $stats.Participants = $tsTabell | Measure-Object | Select-Object -ExpandProperty Count
        $stats.Winner = ($tsTabell[0].ChildNodes | ? name -eq td | select -index 1).InnerText
    } 

    $xl45 = $html.SelectNodes(".//*[contains(@class,'xl45')]")
    if($xl45) {
        $stats.Winner = $xl45[0].innerText
        $stats.Participants = $xl45 | Measure-Object | Select-Object -ExpandProperty Count
    }

    $h1s = $html.SelectNodes("//h1")
    if($h1s) {
        $stats.Title = $h1s[0].innerText
    } else {
        $xl62 = $html.SelectNodes(".//*[contains(@class,'xl62')]")
        if($xl62) {
            $stats.Title = $xl62.innerText.ToLower()
        }
    }

    # Determine who has best rating progress
    $trElements = $html.SelectNodes(".//*[contains(@class,'ts_tabell')]")
    $bestRatingProgress = -999
    if($trElements) {
        $trElements | ForEach-Object {
            $trElement = $_
            $lastTd = $trElement.ChildNodes | ? name -eq td | select -last 1
            $allTds = $trElement.ChildNodes | ? name -eq td
            $txt = $lastTd.innerText.Replace("&nbsp;"," ").Trim()
            if($txt -match "^[0-9]+ \([ ]?[+-]?[0-9]+\)$") {
                $prog = [int] $txt.Split("(")[1].Replace(")","").Trim()
                if($prog -gt $bestRatingProgress) {
                    $bestRatingProgress = $prog
                    $stats.BestRatingProgress = $allTds[1].InnerText + " " + $txt.Split(" ",2)[1].Replace(" ","")
                }
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
$turneringsserviceCache = @{}
function Get-StatsFromTurneringsservice([string]$uri) {
    if($turneringsserviceCache.ContainsKey($uri)) {
        Write-Verbose "Cached stats for $uri" -Verbose
        $content = $turneringsserviceCache[$uri]
    } else {
        Write-verbose "Getting stats for $uri" -Verbose
        $content = Invoke-WebRequest -Uri $uri -UseBasicParsing -ErrorAction Stop
        $turneringsserviceCache[$uri] = $content
    }
    Get-StatsFromContent  $content.Content
}

# Parse all files into objects with statistics
$objects = $files | ForEach-Object {
    $file = $_
    $obj = @{
        Type = "Ukjent"
        File = "resultater/$($file.Name)"
        Date = [Datetime]::ParseExact(([Regex]::Matches($file.BaseName, "[0-9]{6}") | Select-Object -First 1 | Select-Object -exp Value),"yyMMdd", $null)
        Group = $null
        Stats = $null
    }

    if($file.BaseName -like "*-*") {
        $obj.Group = "Gruppe " + ($file.BaseName -split "-" | Select-Object -Last 1)
    }

    if($file.Name -match "Hu[0-9]{6}(-[A-C1-9])?.htm") {
        $obj.Type = "Hurtigsjakk"
        $obj.Stats = Get-StatsFromFile $obj.File
    } elseif($file.Name -match "Ly[0-9]{6}(-[A-C1-9])?.htm") {
        $obj.Type = "Lynsjakk"
        $obj.Stats = Get-StatsFromFile $obj.File
    } elseif($file.Name -match "Hc[0-9]{6}(-[A-C1-9])?.htm") {
        $obj.Type = "Hurtigsjakk med tidshandicap"
        $obj.Stats = Get-StatsFromFile $obj.File
    }

    [PSCustomObject] $obj
} | Sort-Object Date, Group



$text = @() # "# Resultater - Hamar sjakkselskap")


$text += $tournaments | Where-Object active | ForEach-Object -Begin {
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
$text += $objects | Group-Object Type | Where-Object Name -in "Lynsjakk","Hurtigsjakk" | ForEach-Object {
    ""
    "# $($_.Name)"
    ""
    "| Dato | Gruppe | Deltagere | Vinner | Beste ratingfremgang |"
    "|-|-|-|-|-|"

    $_.Group | Sort-Object @{Expression = {$_.Date}; Ascending = $false}, Group | Select-Object -First 5 | ForEach-Object {
        "|[{0}]({1})|{2}|{3}|{4}|{5}|" -f $_.Date.ToString("yyyy-MM-dd"), $_.File, $_.Group, $_.Stats.Participants, $_.Stats.Winner, $_.Stats.BestRatingProgress
    }

    "|[Alle]({0}.md)||||" -f $_.Name
}

$text += ""
$text += "[Alle resultater](arkiv.md)"
$text | Set-Content "index.md" -Encoding UTF8

# Generate per type files
$objects | Group-Object Type | ForEach-Object {
    $text = @(
        "# $($_.Name)"
        ""
        "| Dato | Gruppe | Deltakere | Vinner | Beste ratingfremgang |"
        "|-|-|-|-|-|"
    )

    $text += $_.Group | Sort-Object @{Expression = {$_.Date}; Ascending = $false}, Group | ForEach-Object {
        "|[{0}]({1})|{2}|{3}|{4}|{5}|" -f $_.Date.ToString("yyyy-MM-dd"), $_.File, $_.Group, $_.Stats.Participants, $_.Stats.Winner, $_.Stats.BestRatingProgress
    }

    $text | Set-Content "$($_.Name).md" -Encoding UTF8
}

$tournaments | ForEach-Object -Begin {
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
    Sort-Object @{Expression = {$_.Date}; Ascending = $false}, Group |
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

if($Error) {
    Write-Warning "Fix feil under kj�ring"
    Read-Host -Prompt "Trykk enter for � sende resultater uansett - kan bli meget feil!"    
} 

# stage all changes
# git add -A
# git commit -m 'Nye resultater'
# git push

# Read-Host -Prompt "Sendt, lukk vindu eller klikk enter" 
