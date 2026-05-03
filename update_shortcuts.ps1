$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$wshell = New-Object -ComObject WScript.Shell

function Update-RotationData {
    param([string]$filePath, [hashtable]$toAdd, [string[]]$toRemove)
    $data = @{}
    if (Test-Path -LiteralPath $filePath) {
        Get-Content -LiteralPath $filePath | ForEach-Object {
            if ($_ -match "^(.*?):(\d+)$") {
                $data[$matches[1]] = $matches[2]
            }
        }
    }
    foreach ($key in $toRemove) { $data.Remove($key) }
    foreach ($key in $toAdd.Keys) { $data[$key] = $toAdd[$key] }
    if ($data.Count -eq 0) {
        if (Test-Path -LiteralPath $filePath) { Remove-Item -LiteralPath $filePath }
    } else {
        $content = @()
        foreach ($key in $data.Keys) { $content += "$key:$($data[$key])" }
        $content | Out-File -FilePath $filePath -Encoding utf8
    }
}

# --- Moving Videos and Thumbnails ---
$sourceVideoPath = Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'test1') -ChildPath 'test1 - 7285941165235522862.mp4'
$targetVideoPath = Join-Path -Path Join-Path -Path $PSScriptRoot -ChildPath 'test2' -ChildPath 'test1 - 7285941165235522862.mp4'
if (Test-Path -LiteralPath $sourceVideoPath) {
    Move-Item -LiteralPath $sourceVideoPath -Destination $targetVideoPath -Force
    Write-Host "Moved video to: test2\test1 - 7285941165235522862.mp4"
}
$sourceProjectPath = Join-Path -Path $PSScriptRoot -ChildPath 'test1'
$targetProjectPath = Join-Path -Path $PSScriptRoot -ChildPath 'test2'
$thumbDirs = @('Thumbnails', 'Edit Thumbnails')
foreach ($dirName in $thumbDirs) {
    $srcDirPath = Join-Path -Path $sourceProjectPath -ChildPath $dirName
    $tgtDirPath = Join-Path -Path $targetProjectPath -ChildPath $dirName
    if (-not (Test-Path -LiteralPath $tgtDirPath)) { New-Item -ItemType Directory -Path $tgtDirPath -Force | Out-Null }
    if (Test-Path -LiteralPath $srcDirPath) {
        Get-ChildItem -Path $srcDirPath -Filter 'test1 - 7285941165235522862_*' | ForEach-Object { Move-Item -LiteralPath $_.FullName -Destination $tgtDirPath -Force }
        Write-Host "Moved thumbnails for test1 - 7285941165235522862.mp4 to test2\$dirName"
    }
}
$srcScPath = Join-Path -Path $sourceProjectPath -ChildPath 'sc'
$tgtScPath = Join-Path -Path $targetProjectPath -ChildPath 'sc'
$scFileName = 'test1 - 7285941165235522862.mp4.lnk'
$srcScFile = Join-Path -Path $srcScPath -ChildPath $scFileName
if (Test-Path -LiteralPath $srcScFile) {
    if (-not (Test-Path -LiteralPath $tgtScPath)) { New-Item -ItemType Directory -Path $tgtScPath -Force | Out-Null }
    $tgtScFile = Join-Path -Path $tgtScPath -ChildPath $scFileName
    Move-Item -LiteralPath $srcScFile -Destination $tgtScFile -Force
    $shortcut = $wshell.CreateShortcut($tgtScFile)
    $shortcut.TargetPath = $targetVideoPath
    $shortcut.Save()
    Write-Host "Moved and updated subfolder shortcut for test1 - 7285941165235522862.mp4"
}
$rootScPath = Join-Path -Path $PSScriptRoot -ChildPath 'sc'
$rootScFile = Join-Path -Path $rootScPath -ChildPath $scFileName
if (Test-Path -LiteralPath $rootScFile) {
    $shortcut = $wshell.CreateShortcut($rootScFile)
    $shortcut.TargetPath = $targetVideoPath
    $shortcut.Save()
    Write-Host "Updated root shortcut target for test1 - 7285941165235522862.mp4"
}

$sourceVideoPath = Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'test2') -ChildPath 'test2 - 7603115258047237398.mp4'
$targetVideoPath = Join-Path -Path Join-Path -Path $PSScriptRoot -ChildPath 'test3' -ChildPath 'test2 - 7603115258047237398.mp4'
if (Test-Path -LiteralPath $sourceVideoPath) {
    Move-Item -LiteralPath $sourceVideoPath -Destination $targetVideoPath -Force
    Write-Host "Moved video to: test3\test2 - 7603115258047237398.mp4"
}
$sourceProjectPath = Join-Path -Path $PSScriptRoot -ChildPath 'test2'
$targetProjectPath = Join-Path -Path $PSScriptRoot -ChildPath 'test3'
$thumbDirs = @('Thumbnails', 'Edit Thumbnails')
foreach ($dirName in $thumbDirs) {
    $srcDirPath = Join-Path -Path $sourceProjectPath -ChildPath $dirName
    $tgtDirPath = Join-Path -Path $targetProjectPath -ChildPath $dirName
    if (-not (Test-Path -LiteralPath $tgtDirPath)) { New-Item -ItemType Directory -Path $tgtDirPath -Force | Out-Null }
    if (Test-Path -LiteralPath $srcDirPath) {
        Get-ChildItem -Path $srcDirPath -Filter 'test2 - 7603115258047237398_*' | ForEach-Object { Move-Item -LiteralPath $_.FullName -Destination $tgtDirPath -Force }
        Write-Host "Moved thumbnails for test2 - 7603115258047237398.mp4 to test3\$dirName"
    }
}
$srcScPath = Join-Path -Path $sourceProjectPath -ChildPath 'sc'
$tgtScPath = Join-Path -Path $targetProjectPath -ChildPath 'sc'
$scFileName = 'test2 - 7603115258047237398.mp4.lnk'
$srcScFile = Join-Path -Path $srcScPath -ChildPath $scFileName
if (Test-Path -LiteralPath $srcScFile) {
    if (-not (Test-Path -LiteralPath $tgtScPath)) { New-Item -ItemType Directory -Path $tgtScPath -Force | Out-Null }
    $tgtScFile = Join-Path -Path $tgtScPath -ChildPath $scFileName
    Move-Item -LiteralPath $srcScFile -Destination $tgtScFile -Force
    $shortcut = $wshell.CreateShortcut($tgtScFile)
    $shortcut.TargetPath = $targetVideoPath
    $shortcut.Save()
    Write-Host "Moved and updated subfolder shortcut for test2 - 7603115258047237398.mp4"
}
$rootScPath = Join-Path -Path $PSScriptRoot -ChildPath 'sc'
$rootScFile = Join-Path -Path $rootScPath -ChildPath $scFileName
if (Test-Path -LiteralPath $rootScFile) {
    $shortcut = $wshell.CreateShortcut($rootScFile)
    $shortcut.TargetPath = $targetVideoPath
    $shortcut.Save()
    Write-Host "Updated root shortcut target for test2 - 7603115258047237398.mp4"
}

$sourceVideoPath = Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'test2') -ChildPath 'test2 - 7603117255995559190.mp4'
$targetVideoPath = Join-Path -Path Join-Path -Path $PSScriptRoot -ChildPath 'test3' -ChildPath 'test2 - 7603117255995559190.mp4'
if (Test-Path -LiteralPath $sourceVideoPath) {
    Move-Item -LiteralPath $sourceVideoPath -Destination $targetVideoPath -Force
    Write-Host "Moved video to: test3\test2 - 7603117255995559190.mp4"
}
$sourceProjectPath = Join-Path -Path $PSScriptRoot -ChildPath 'test2'
$targetProjectPath = Join-Path -Path $PSScriptRoot -ChildPath 'test3'
$thumbDirs = @('Thumbnails', 'Edit Thumbnails')
foreach ($dirName in $thumbDirs) {
    $srcDirPath = Join-Path -Path $sourceProjectPath -ChildPath $dirName
    $tgtDirPath = Join-Path -Path $targetProjectPath -ChildPath $dirName
    if (-not (Test-Path -LiteralPath $tgtDirPath)) { New-Item -ItemType Directory -Path $tgtDirPath -Force | Out-Null }
    if (Test-Path -LiteralPath $srcDirPath) {
        Get-ChildItem -Path $srcDirPath -Filter 'test2 - 7603117255995559190_*' | ForEach-Object { Move-Item -LiteralPath $_.FullName -Destination $tgtDirPath -Force }
        Write-Host "Moved thumbnails for test2 - 7603117255995559190.mp4 to test3\$dirName"
    }
}
$srcScPath = Join-Path -Path $sourceProjectPath -ChildPath 'sc'
$tgtScPath = Join-Path -Path $targetProjectPath -ChildPath 'sc'
$scFileName = 'test2 - 7603117255995559190.mp4.lnk'
$srcScFile = Join-Path -Path $srcScPath -ChildPath $scFileName
if (Test-Path -LiteralPath $srcScFile) {
    if (-not (Test-Path -LiteralPath $tgtScPath)) { New-Item -ItemType Directory -Path $tgtScPath -Force | Out-Null }
    $tgtScFile = Join-Path -Path $tgtScPath -ChildPath $scFileName
    Move-Item -LiteralPath $srcScFile -Destination $tgtScFile -Force
    $shortcut = $wshell.CreateShortcut($tgtScFile)
    $shortcut.TargetPath = $targetVideoPath
    $shortcut.Save()
    Write-Host "Moved and updated subfolder shortcut for test2 - 7603117255995559190.mp4"
}
$rootScPath = Join-Path -Path $PSScriptRoot -ChildPath 'sc'
$rootScFile = Join-Path -Path $rootScPath -ChildPath $scFileName
if (Test-Path -LiteralPath $rootScFile) {
    $shortcut = $wshell.CreateShortcut($rootScFile)
    $shortcut.TargetPath = $targetVideoPath
    $shortcut.Save()
    Write-Host "Updated root shortcut target for test2 - 7603117255995559190.mp4"
}

$sourceVideoPath = Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'test3') -ChildPath 'test3 - 7081013928339213573.mp4'
$targetVideoPath = Join-Path -Path Join-Path -Path $PSScriptRoot -ChildPath 'test1' -ChildPath 'test3 - 7081013928339213573.mp4'
if (Test-Path -LiteralPath $sourceVideoPath) {
    Move-Item -LiteralPath $sourceVideoPath -Destination $targetVideoPath -Force
    Write-Host "Moved video to: test1\test3 - 7081013928339213573.mp4"
}
$sourceProjectPath = Join-Path -Path $PSScriptRoot -ChildPath 'test3'
$targetProjectPath = Join-Path -Path $PSScriptRoot -ChildPath 'test1'
$thumbDirs = @('Thumbnails', 'Edit Thumbnails')
foreach ($dirName in $thumbDirs) {
    $srcDirPath = Join-Path -Path $sourceProjectPath -ChildPath $dirName
    $tgtDirPath = Join-Path -Path $targetProjectPath -ChildPath $dirName
    if (-not (Test-Path -LiteralPath $tgtDirPath)) { New-Item -ItemType Directory -Path $tgtDirPath -Force | Out-Null }
    if (Test-Path -LiteralPath $srcDirPath) {
        Get-ChildItem -Path $srcDirPath -Filter 'test3 - 7081013928339213573_*' | ForEach-Object { Move-Item -LiteralPath $_.FullName -Destination $tgtDirPath -Force }
        Write-Host "Moved thumbnails for test3 - 7081013928339213573.mp4 to test1\$dirName"
    }
}
$srcScPath = Join-Path -Path $sourceProjectPath -ChildPath 'sc'
$tgtScPath = Join-Path -Path $targetProjectPath -ChildPath 'sc'
$scFileName = 'test3 - 7081013928339213573.mp4.lnk'
$srcScFile = Join-Path -Path $srcScPath -ChildPath $scFileName
if (Test-Path -LiteralPath $srcScFile) {
    if (-not (Test-Path -LiteralPath $tgtScPath)) { New-Item -ItemType Directory -Path $tgtScPath -Force | Out-Null }
    $tgtScFile = Join-Path -Path $tgtScPath -ChildPath $scFileName
    Move-Item -LiteralPath $srcScFile -Destination $tgtScFile -Force
    $shortcut = $wshell.CreateShortcut($tgtScFile)
    $shortcut.TargetPath = $targetVideoPath
    $shortcut.Save()
    Write-Host "Moved and updated subfolder shortcut for test3 - 7081013928339213573.mp4"
}
$rootScPath = Join-Path -Path $PSScriptRoot -ChildPath 'sc'
$rootScFile = Join-Path -Path $rootScPath -ChildPath $scFileName
if (Test-Path -LiteralPath $rootScFile) {
    $shortcut = $wshell.CreateShortcut($rootScFile)
    $shortcut.TargetPath = $targetVideoPath
    $shortcut.Save()
    Write-Host "Updated root shortcut target for test3 - 7081013928339213573.mp4"
}

$sourceVideoPath = Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'test3') -ChildPath 'test3 - 7046490541009276166.mp4'
$targetVideoPath = Join-Path -Path Join-Path -Path $PSScriptRoot -ChildPath 'test1' -ChildPath 'test3 - 7046490541009276166.mp4'
if (Test-Path -LiteralPath $sourceVideoPath) {
    Move-Item -LiteralPath $sourceVideoPath -Destination $targetVideoPath -Force
    Write-Host "Moved video to: test1\test3 - 7046490541009276166.mp4"
}
$sourceProjectPath = Join-Path -Path $PSScriptRoot -ChildPath 'test3'
$targetProjectPath = Join-Path -Path $PSScriptRoot -ChildPath 'test1'
$thumbDirs = @('Thumbnails', 'Edit Thumbnails')
foreach ($dirName in $thumbDirs) {
    $srcDirPath = Join-Path -Path $sourceProjectPath -ChildPath $dirName
    $tgtDirPath = Join-Path -Path $targetProjectPath -ChildPath $dirName
    if (-not (Test-Path -LiteralPath $tgtDirPath)) { New-Item -ItemType Directory -Path $tgtDirPath -Force | Out-Null }
    if (Test-Path -LiteralPath $srcDirPath) {
        Get-ChildItem -Path $srcDirPath -Filter 'test3 - 7046490541009276166_*' | ForEach-Object { Move-Item -LiteralPath $_.FullName -Destination $tgtDirPath -Force }
        Write-Host "Moved thumbnails for test3 - 7046490541009276166.mp4 to test1\$dirName"
    }
}
$srcScPath = Join-Path -Path $sourceProjectPath -ChildPath 'sc'
$tgtScPath = Join-Path -Path $targetProjectPath -ChildPath 'sc'
$scFileName = 'test3 - 7046490541009276166.mp4.lnk'
$srcScFile = Join-Path -Path $srcScPath -ChildPath $scFileName
if (Test-Path -LiteralPath $srcScFile) {
    if (-not (Test-Path -LiteralPath $tgtScPath)) { New-Item -ItemType Directory -Path $tgtScPath -Force | Out-Null }
    $tgtScFile = Join-Path -Path $tgtScPath -ChildPath $scFileName
    Move-Item -LiteralPath $srcScFile -Destination $tgtScFile -Force
    $shortcut = $wshell.CreateShortcut($tgtScFile)
    $shortcut.TargetPath = $targetVideoPath
    $shortcut.Save()
    Write-Host "Moved and updated subfolder shortcut for test3 - 7046490541009276166.mp4"
}
$rootScPath = Join-Path -Path $PSScriptRoot -ChildPath 'sc'
$rootScFile = Join-Path -Path $rootScPath -ChildPath $scFileName
if (Test-Path -LiteralPath $rootScFile) {
    $shortcut = $wshell.CreateShortcut($rootScFile)
    $shortcut.TargetPath = $targetVideoPath
    $shortcut.Save()
    Write-Host "Updated root shortcut target for test3 - 7046490541009276166.mp4"
}

# --- Creating Shortcuts ---
$rootScFolder = Join-Path -Path $PSScriptRoot -ChildPath 'sc'
if (-not (Test-Path -Path $rootScFolder)) { New-Item -ItemType Directory -Path $rootScFolder -Force | Out-Null }

$subfolderScPath = Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'test1') -ChildPath 'sc'
if (-not (Test-Path -Path $subfolderScPath)) { New-Item -ItemType Directory -Path $subfolderScPath -Force | Out-Null }
$subfolderScPath = Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'test2') -ChildPath 'sc'
if (-not (Test-Path -Path $subfolderScPath)) { New-Item -ItemType Directory -Path $subfolderScPath -Force | Out-Null }
$subfolderScPath = Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'test3') -ChildPath 'sc'
if (-not (Test-Path -Path $subfolderScPath)) { New-Item -ItemType Directory -Path $subfolderScPath -Force | Out-Null }

    # Shortcut for: test3\test3 - 7063353082922519814.mp4 (Type: root-sc)
    $targetPath = Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'test3') -ChildPath 'test3 - 7063353082922519814.mp4'
    $shortcutPath = Join-Path -Path $rootScFolder -ChildPath ('test3 - 7063353082922519814.mp4' + '.lnk')
    try {
        $shortcut = $wshell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $targetPath
        $shortcut.Save()
        Write-Host "Created root shortcut: test3 - 7063353082922519814.mp4"
    } catch {
        Write-Host "FAILED to create shortcut: $shortcutPath" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }

# --- Updating Rotation Data ---
$toAdd = @{
}
$toRemove = @(
    'test1 - 7285941165235522862.mp4'
)
Update-RotationData -filePath (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'test1') -ChildPath 'rotation_data.txt') -toAdd $toAdd -toRemove $toRemove
Write-Host "Surgically updated rotation_data.txt for test1"
$toAdd = @{
    'test1 - 7285941165235522862.mp4' = '270'
}
$toRemove = @(
    'test2 - 7603115258047237398.mp4'
    'test2 - 7603117255995559190.mp4'
)
Update-RotationData -filePath (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'test2') -ChildPath 'rotation_data.txt') -toAdd $toAdd -toRemove $toRemove
Write-Host "Surgically updated rotation_data.txt for test2"
$toAdd = @{
}
$toRemove = @(
    'test3 - 7081013928339213573.mp4'
    'test3 - 7046490541009276166.mp4'
)
Update-RotationData -filePath (Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'test3') -ChildPath 'rotation_data.txt') -toAdd $toAdd -toRemove $toRemove
Write-Host "Surgically updated rotation_data.txt for test3"

# --- Update Central LUA Rotation File ---
$luaPath = 'C:\Bridge\misc\tools\mpv-x86_64-v3-20260418-git-4377cce\portable_config\scripts\autorotate.lua'
$script:rotatedFiles = @{}
$script:rotatedFiles['test1 - 7285941165235522862.mp4'] = 270
$script:rotatedFiles['test3 - 7103232606384180486.mp4'] = 270
$script:rotatedFiles['test3 - 7118881399578316038.mp4'] = 270
$script:rotatedFiles['test3 - 7187110285096848646.mp4'] = 270
$script:rotatedFiles['test3 - 7461845602947566853.mp4'] = 90
$script:rotatedFiles['test3 - 7570424224985222408.mp4'] = 270
$script:rotatedFiles['test3 - 7570579895617703186.mp4'] = 270
$script:rotatedFiles['test3 - 7591604935335120146.mp4'] = 270
$allRotations = @{}
if (Test-Path -LiteralPath $luaPath) {
    $existingContent = Get-Content -LiteralPath $luaPath -Raw
    if ($existingContent -match '(?s)local rotations = \{(.*?)\}') {
        $inner = $matches[1]
        $entries = $inner -split ','
        foreach ($entry in $entries) {
            if ($entry -match "\[['\"" ]*(.*?)['\"" ]*\]\s*=\s*['\"" ]?(\d+)['\"" ]?") {
                $allRotations[$matches[1]] = $matches[2]
            }
        }
    }
}
foreach ($file in $script:rotatedFiles.Keys) { $allRotations[$file] = $script:rotatedFiles[$file] }
if ($allRotations.Count -gt 0) {
    $luaDir = Split-Path -Parent $luaPath
    if (-not (Test-Path -Path $luaDir)) { New-Item -ItemType Directory -Path $luaDir -Force | Out-Null }
    $luaContent = @()
    $luaContent += 'local rotations = {'
    foreach ($file in $allRotations.Keys) {
        $rot = $allRotations[$file]
        $escapedFile = $file.Replace("'", "\\'")
        $luaContent += "    ['$escapedFile'] = $rot,"
    }
    $luaContent += '}'
    $luaContent += ''
    $luaContent += 'mp.register_event("file-loaded", function()'
    $luaContent += '    local path = mp.get_property("path")'
    $luaContent += '    if not path then return end'
    $luaContent += ''
    $luaContent += '    mp.set_property("video-rotate", 0)'
    $luaContent += ''
    $luaContent += '    local filename = path:match("([^/\\\\]+)$") or path'
    $luaContent += ''
    $luaContent += '    if rotations[filename] then'
    $luaContent += '        mp.set_property("video-rotate", rotations[filename])'
    $luaContent += '    end'
    $luaContent += 'end)'
    $luaContent | Out-File -FilePath $luaPath -Encoding utf8
    Write-Host "Updated MPV autorotate script at: $luaPath"
}

# --- Placing Genuine Timestamps ---
Write-Host "Placing genuine timestamp for project: test1"
$projectFolder = Join-Path -Path $PSScriptRoot -ChildPath 'test1'
$scDateFile = Join-Path -Path $projectFolder -ChildPath 'scdate.txt'
Set-Content -Path $scDateFile -Value "2026-04-29T19:41:19.382Z"

Write-Host "Placing genuine timestamp for project: test2"
$projectFolder = Join-Path -Path $PSScriptRoot -ChildPath 'test2'
$scDateFile = Join-Path -Path $projectFolder -ChildPath 'scdate.txt'
Set-Content -Path $scDateFile -Value "2026-04-29T19:41:19.382Z"

Write-Host "Placing genuine timestamp for project: test3"
$projectFolder = Join-Path -Path $PSScriptRoot -ChildPath 'test3'
$scDateFile = Join-Path -Path $projectFolder -ChildPath 'scdate.txt'
Set-Content -Path $scDateFile -Value "2026-04-29T19:41:19.382Z"

Write-Host "Script execution complete."
Read-Host -Prompt "Press Enter to exit"