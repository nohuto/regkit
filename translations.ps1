# Copyright (C) 2026 nohuto
# This is used to autogenerate the template & merge new text into existing po files

param([Parameter(Mandatory)][string]$gettext, [string]$language)

$src = "$PSScriptRoot\src"
$lang = "$PSScriptRoot\assets\lang"
$pot = "$lang\regkit.pot"
$temp = Join-Path ([IO.Path]::GetTempPath()) "regkit-translations-$PID"
$utf8 = New-Object Text.UTF8Encoding $false
$header = "# RegKit translation template.`n# Copyright (C) 2026 nohuto`n# This file is distributed under the same license as RegKit.`n#`n"

function invokeTool([string]$tool, [string[]]$arguments) {
    & (Join-Path $gettext "$tool.exe") @arguments 2>&1 | % { "$_" }
    if ($LASTEXITCODE) { throw "$tool - $LASTEXITCODE" }
}

function addText($texts, [string]$text, [string]$reference) {
    if (!$texts.Contains($text)) { $texts[$text] = [Activator]::CreateInstance([System.Collections.Generic.List[string]]) }
    $texts[$text].Add($reference)
}

function getLang([string]$code) {
    try { $culture = [Globalization.CultureInfo]::GetCultureInfo($code) } catch { $culture = $null }
    if (!$culture -or !$culture.Name -or $culture.LCID -eq 4096) { throw "'$code' isnt a win language name (e.g. fr, pt-BR)" }
    if ($culture.TwoLetterISOLanguageName -eq 'en') { throw "English is built in and needs no .po file" }
    $culture.Name
}

# get dialog captions & control labels from rc files
function addRcTexts($texts) {
    foreach ($file in gci "$PSScriptRoot\resources\*.rc" -ErrorAction Stop) {
        $lines = [IO.File]::ReadAllLines($file.FullName)
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match '^\s*(CAPTION|LTEXT|RTEXT|CTEXT|PUSHBUTTON|DEFPUSHBUTTON|CHECKBOX|AUTOCHECKBOX|RADIOBUTTON|AUTORADIOBUTTON|GROUPBOX|CONTROL)\s+"((?:[^"]|"")*[A-Za-z](?:[^"]|"")*)"') {
                addText $texts ($Matches[2] -replace '""', '\"') "resources/$($file.Name):$($i + 1)"
            }
        }
    }
}

# get display labels from NUL separated file filters
function addFilterTexts($texts, $files) {
    foreach ($file in $files) {
        $content = [IO.File]::ReadAllText($file)
        if (!$content.Contains('\0')) { continue }
        $groups = [Activator]::CreateInstance([System.Collections.Generic.List[object]])
        $end = -1
        foreach ($match in [regex]::Matches($content, 'L"((?:[^"\\\r\n]|\\.)*)"')) {
            if ($groups.Count -and [string]::IsNullOrWhiteSpace($content.Substring($end, $match.Index - $end))) { $groups[$groups.Count - 1].text += $match.Groups[1].Value }
            else { $groups.Add([pscustomobject]@{ start = $match.Index; text = $match.Groups[1].Value }) }
            $end = $match.Index + $match.Length
        }
        $path = $file.Substring($src.Length + 1).Replace('\', '/')
        foreach ($group in $groups | ? { $_.text.Contains('\0') }) {
            $reference = "${path}:$($content.Substring(0, $group.start).Split("`n").Count)"
            $parts = $group.text -split '\\0'
            for ($i = 0; $i -lt $parts.Count; $i += 2) {
                if (($parts[$i] -replace '\\.', '') -match '[A-Za-z]') { addText $texts $parts[$i] $reference }
            }
        }
    }
}

function updateTemplate {
    $files = (gci $src -Recurse -Include *.cpp, *.h -ErrorAction Stop).FullName
    [IO.File]::WriteAllLines("$temp\files.txt", [string[]]($files | % { $_.Substring($src.Length + 1).Replace('\', '/') }), $utf8)

    invokeTool xgettext '--language=C++', '--from-code=UTF-8', '--keyword=Tr', '--keyword=TrNoop', '--keyword=TrDetail:1', '--keyword=TrLabel:1', '--add-comments=TRANSLATORS', '--no-wrap', '--package-name=RegKit', '--msgid-bugs-address=https://github.com/nohuto/regkit/issues', '-D', $src, '-f', "$temp\files.txt", '-o', "$temp\code.pot"

    $texts = [ordered]@{}
    addRcTexts $texts
    addFilterTexts $texts $files
    $lines = [Collections.Generic.List[string]]@('msgid ""', 'msgstr ""', '"Content-Type: text/plain; charset=UTF-8\n"', '')
    foreach ($entry in $texts.GetEnumerator()) { $lines.AddRange([string[]]@("#: $($entry.Value -join ' ')", "msgid `"$($entry.Key)`"", 'msgstr ""', '')) }
    [IO.File]::WriteAllLines("$temp\extra.pot", $lines, $utf8)

    # combine everything into template
    invokeTool msgcat '--use-first', '--no-wrap', '-o', "$temp\regkit.pot", "$temp\code.pot", "$temp\extra.pot"
    $content = [IO.File]::ReadAllText("$temp\regkit.pot")
    $start = $content.IndexOf('#, fuzzy')
    if ($start -lt 0) { throw "msgcat wrote no template header" }
    [IO.File]::WriteAllText($pot, $header + $content.Substring($start), $utf8)
}

$failed = $false

try {
    foreach ($tool in 'xgettext', 'msgcat', 'msginit', 'msgmerge', 'msgfmt') {
        if (!(Test-Path -LiteralPath (Join-Path $gettext "$tool.exe"))) { throw "$tool.exe wasnt found in '$gettext', pass the bin folder of GNU gettext" }
    }
    if ($language) { $language = getLang $language }
    ni -ItemType Directory -Force $temp -ErrorAction Stop | Out-Null
    updateTemplate

    # create lang catalog if theres none
    if ($language) {
        if (Test-Path -LiteralPath "$lang\$language.po") { Write-Host "$language.po already exists, its updated instead" }
        else { invokeTool msginit '--no-translator', '--no-wrap', "--locale=$($language.Replace('-', '_'))", '-i', $pot, '-o', "$lang\$language.po" }
    }
    
    foreach ($po in gci "$lang\*.po") {
        try {
            invokeTool msgmerge '--quiet', '--update', '--backup=none', '--no-wrap', $po.FullName, $pot
            invokeTool msgfmt '--check-format', '--statistics', '-o', 'NUL', $po.FullName | % { "$($po.Name): $_" }
        }
        catch {
            Write-Host "$($po.Name): $($_.Exception.Message)" -ForegroundColor Red
            $failed = $true
        }
    }
}

catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    $failed = $true
}

finally {
    ri -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue
}

if ($failed) { exit 1 }
