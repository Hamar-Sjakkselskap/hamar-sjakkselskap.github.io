$files = dir .\resultater
$objects = $files | foreach {
    $obj = @{
        Type = "Ukjent"
        File = "resultater/$($_.Name)"
        Date = [Datetime]::ParseExact(([Regex]::Matches($_.BaseName, "[0-9]{6}") | Select -First 1 | Select -exp Value),"yyMMdd", $null)
        Group = $null
    }

    if($_.BaseName -like "*-*") {
        $obj.Group = $_.BaseName -split "-" | Select-Object -Last 1
    }

    if($_.Name -match "Hu[0-9]{6}(-[A-C1-9])?.htm") {
        $obj.Type = "Hurtigsjakk"
    } elseif($_.Name -match "Ly[0-9]{6}(-[A-C1-9])?.htm") {
        $obj.Type = "Lynsjakk"
    } elseif($_.Name -match "Hc[0-9]{6}(-[A-C1-9])?.htm") {
        $obj.Type = "Hamarcup"
    }

    [PSCustomObject] $obj
}

$text = @("# Resultater")

# Generate summary file
$text += $objects | Group Type | Foreach {
    ""
    "## $($_.Name)"

    $_.Group | Sort Date -Descending | Select -First 3 | Foreach {
        $Group = ""
        if($_.Group) {
            $Group = " - Gruppe $($_.Group)"
        }
        " - [{0}{2}]({1})" -f $_.Date.ToString("yyyy-MM-dd"), $_.File, $Group
    }

    " - [Alle]({0}.md)" -f $_.Name
}

$text | Set-Content "index.md"

# Generate per type files
$objects | Group Type | Foreach {
    $text = @("# $($_.Name)")

    $text += $_.Group | Sort Date -Descending | Foreach {
        $Group = ""
        if($_.Group) {
            $Group = " - Gruppe $($_.Group)"
        }
        " - [{0}{2}]({1})" -f $_.Date.ToString("yyyy-MM-dd"), $_.File, $Group
    }

    $text | Set-Content "$($_.Name).md"
}