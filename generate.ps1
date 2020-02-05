$files = dir .\resultater

function Get-Stats([string]$file) {
    Write-verbose "Getting winner for $file" -Verbose
    $content = Get-Content -Raw $file -Encoding Default
    
    $stats = @{
        Winner = $null
        Participants = 0
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

    return $stats
}

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
        $obj.Stats = Get-Stats $obj.File
    } elseif($_.Name -match "Ly[0-9]{6}(-[A-C1-9])?.htm") {
        $obj.Type = "Lynsjakk"
        $obj.Stats = Get-Stats $obj.File
    } elseif($_.Name -match "Hc[0-9]{6}(-[A-C1-9])?.htm") {
        $obj.Type = "Hamarcup"
        $obj.Stats = Get-Stats $obj.File
    }

    [PSCustomObject] $obj
} | sort Date, Group



$text = @("# Resultater")

# Generate summary file
$text += $objects | Group Type | Foreach {
    ""
    "## $($_.Name)"
    ""
    "| Dato | Gruppe | Deltagere | Vinner |"
    "|-|-|-|-|"

    $_.Group | Sort @{Expression = {$_.Date}; Ascending = $false}, Group | Select -First 5 | Foreach {
        "|[{0}]({1})|{2}|{3}|{4}|" -f $_.Date.ToString("yyyy-MM-dd"), $_.File, $_.Group, $_.Stats.Participants, $_.Stats.Winner
    }

    "|[Alle]({0}.md)||||" -f $_.Name
}

$text | Set-Content "index.md" -Encoding UTF8

# Generate per type files
$objects | Group Type | Foreach {
    $text = @(
        "# $($_.Name)"
        ""
        "| Dato | Gruppe | Deltakere | Vinner |"
        "|-|-|-|-|"
    )

    $text += $_.Group | sort @{Expression = {$_.Date}; Ascending = $false}, Group | Foreach {
        "|[{0}]({1})|{2}|{3}|{4}|" -f $_.Date.ToString("yyyy-MM-dd"), $_.File, $_.Group, $_.Stats.Participants, $_.Stats.Winner
    }

    $text | Set-Content "$($_.Name).md" -Encoding UTF8
}