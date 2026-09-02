[CmdletBinding()]
param(
    [string]$Language,
    [switch]$SelfTest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$Repository = 'MatthewLukePublishing/typo-checker'
$GitHubApiBase = "https://api.github.com/repos/$Repository"
$WorkspacePrefix = 'matthew-luke-typo-checker-'
$RequiredFiles = @(
    'Typo_Checker.js',
    'Export All Content to Excel.jsx',
    'package.json'
)

function Assert-CommitSha {
    param([Parameter(Mandatory)] [string]$Value)
    if ($Value -notmatch '^[A-Fa-f0-9]{40}$') {
        throw "Expected an immutable 40-character Git commit SHA; found '$Value'."
    }
}

function Assert-ReleaseTag {
    param([Parameter(Mandatory)] [string]$Value)
    if ($Value -notmatch '^v\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?$') {
        throw "The latest GitHub release has an unsupported tag '$Value'."
    }
}

function Assert-Language {
    param([Parameter(Mandatory)] [string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value) -or $Value.Length -gt 80 -or
        $Value -match '[\x00-\x1F\x7F]' -or
        $Value -notmatch '^[\p{L}\p{M}][\p{L}\p{M}0-9 .()_-]*$') {
        throw "Language must be a plain language name of 1-80 characters, such as 'German'."
    }
}

function New-GitHubApiUri {
    param([Parameter(Mandatory)] [string]$RelativePath)
    if (-not $RelativePath.StartsWith('/') -or $RelativePath.Contains('..') -or
        $RelativePath.Contains('?') -or $RelativePath.Contains('#') -or
        $RelativePath -match '^[A-Za-z][A-Za-z0-9+.-]*:') {
        throw "Refusing an invalid GitHub API path: $RelativePath"
    }
    $uri = [Uri]("$GitHubApiBase$RelativePath")
    if ($uri.Scheme -ne 'https' -or $uri.Host -ne 'api.github.com' -or
        -not $uri.AbsolutePath.StartsWith("/repos/$Repository/", [StringComparison]::Ordinal)) {
        throw "Refusing a GitHub API URI outside the canonical repository: $uri"
    }
    return $uri
}

function Invoke-GitHubJson {
    param([Parameter(Mandatory)] [string]$RelativePath)
    $uri = New-GitHubApiUri -RelativePath $RelativePath
    $headers = @{
        Accept = 'application/vnd.github+json'
        'X-GitHub-Api-Version' = '2022-11-28'
        'User-Agent' = 'MatthewLukePublishing-TypoChecker'
    }
    return Invoke-RestMethod -Method Get -Uri $uri -Headers $headers
}

function Get-LatestReleaseResolution {
    $release = Invoke-GitHubJson -RelativePath '/releases/latest'
    if ($release.draft -or $release.prerelease) {
        throw 'GitHub returned a draft or prerelease instead of the latest stable release.'
    }
    $tag = [string]$release.tag_name
    Assert-ReleaseTag -Value $tag
    $encodedTag = [Uri]::EscapeDataString($tag)
    $reference = Invoke-GitHubJson -RelativePath "/git/ref/tags/$encodedTag"
    $objectType = [string]$reference.object.type
    $objectSha = [string]$reference.object.sha
    Assert-CommitSha -Value $objectSha
    for ($depth = 0; $depth -lt 5; $depth++) {
        if ($objectType -eq 'commit') {
            return [pscustomobject]@{
                Tag = $tag
                CommitSha = $objectSha.ToLowerInvariant()
                ReleaseUrl = [string]$release.html_url
            }
        }
        if ($objectType -ne 'tag') {
            throw "Release tag '$tag' resolves to unsupported Git object type '$objectType'."
        }
        $tagObject = Invoke-GitHubJson -RelativePath "/git/tags/$objectSha"
        $objectType = [string]$tagObject.object.type
        $objectSha = [string]$tagObject.object.sha
        Assert-CommitSha -Value $objectSha
    }
    throw "Release tag '$tag' did not resolve to a commit within five tag objects."
}

function Get-ArchiveUri {
    param([Parameter(Mandatory)] [string]$CommitSha)
    Assert-CommitSha -Value $CommitSha
    $uri = [Uri]("https://github.com/$Repository/archive/$CommitSha.zip")
    $expectedPath = "/$Repository/archive/$CommitSha.zip"
    if ($uri.Scheme -ne 'https' -or $uri.Host -ne 'github.com' -or
        -not [string]::Equals($uri.AbsolutePath, $expectedPath, [StringComparison]::Ordinal)) {
        throw "Refusing an archive URI outside the canonical repository: $uri"
    }
    return $uri
}

function Assert-SafeWorkspacePath {
    param([Parameter(Mandatory)] [string]$Path)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\', '/')
    $resolved = [IO.Path]::GetFullPath($Path).TrimEnd('\', '/')
    $parent = [IO.Path]::GetDirectoryName($resolved).TrimEnd('\', '/')
    $leaf = [IO.Path]::GetFileName($resolved)
    if (-not [string]::Equals($parent, $tempRoot, [StringComparison]::OrdinalIgnoreCase) -or
        $leaf -notmatch '^matthew-luke-typo-checker-[a-f0-9]{32}$') {
        throw "Refusing an unexpected temporary workspace path: $resolved"
    }
    return $resolved
}

function New-SafeWorkspace {
    $candidate = Join-Path ([IO.Path]::GetTempPath()) ($WorkspacePrefix + [Guid]::NewGuid().ToString('N'))
    $resolved = Assert-SafeWorkspacePath -Path $candidate
    [void](New-Item -ItemType Directory -Path $resolved)
    return $resolved
}

function Remove-SafeWorkspace {
    param([Parameter(Mandatory)] [string]$Path)
    $resolved = Assert-SafeWorkspacePath -Path $Path
    if ([IO.Directory]::Exists($resolved)) {
        [IO.Directory]::Delete($resolved, $true)
    }
}

function Get-NodeExecutable {
    $command = Get-Command node.exe -ErrorAction SilentlyContinue
    if ($null -eq $command) {
        throw 'Node.js 20 or newer is required and was not found on PATH.'
    }
    $versionText = (& $command.Source --version).Trim().TrimStart('v')
    $version = $null
    if (-not [Version]::TryParse($versionText, [ref]$version) -or $version.Major -lt 20) {
        throw "Node.js 20 or newer is required; found '$versionText'."
    }
    return $command.Source
}

function Save-ProcessEnvironment {
    param([Parameter(Mandatory)] [string[]]$Names)
    $saved = @{}
    foreach ($name in $Names) {
        $saved[$name] = [pscustomobject]@{
            Exists = Test-Path -LiteralPath "Env:$name"
            Value = [Environment]::GetEnvironmentVariable($name, 'Process')
        }
    }
    return $saved
}

function Restore-ProcessEnvironment {
    param([Parameter(Mandatory)] [hashtable]$Saved)
    foreach ($name in $Saved.Keys) {
        if ($Saved[$name].Exists) {
            [Environment]::SetEnvironmentVariable($name, [string]$Saved[$name].Value, 'Process')
        } else {
            [Environment]::SetEnvironmentVariable($name, $null, 'Process')
        }
    }
}

function Invoke-OfflineSelfTest {
    $testWorkspace = New-SafeWorkspace
    try {
        Assert-Language -Value 'German'
        Assert-CommitSha -Value ('a' * 40)
        Assert-ReleaseTag -Value 'v1.1.0'
        $archiveUri = Get-ArchiveUri -CommitSha ('b' * 40)
        if ($archiveUri.Host -ne 'github.com') { throw 'Archive host validation self-test failed.' }
        if ((New-GitHubApiUri -RelativePath '/releases/latest').Host -ne 'api.github.com') {
            throw 'GitHub API validation self-test failed.'
        }
        $rejected = $false
        try { [void](Assert-SafeWorkspacePath -Path (Join-Path ([IO.Path]::GetTempPath()) 'unexpected')) } catch { $rejected = $true }
        if (-not $rejected) { throw 'Temporary-path rejection self-test failed.' }
        $nodePath = Get-NodeExecutable
        if (-not (Test-Path -LiteralPath $nodePath -PathType Leaf)) { throw 'Node discovery self-test failed.' }
        [pscustomobject]@{
            status = 'ok'
            networkUsed = $false
            inDesignUsed = $false
            immutableCommitRequired = $true
            canonicalRepositoryLocked = $true
            safeCleanupBoundary = $true
            node20OrNewer = $true
        } | ConvertTo-Json
    } finally {
        Remove-SafeWorkspace -Path $testWorkspace
    }
}

if ($SelfTest) {
    Invoke-OfflineSelfTest
    exit 0
}

if ($env:OS -ne 'Windows_NT') {
    throw 'This launcher requires Windows and Adobe InDesign 2026.'
}
Assert-Language -Value $Language
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$workspace = $null
$savedEnvironment = $null
$environmentNames = @(
    'OPEN_INDESIGN_DOCUMENT',
    'USE_OPEN_INDESIGN_DOCUMENT',
    'CHECK_LANGUAGE',
    'TYPO_CHECKER_RELEASE_TAG',
    'TYPO_CHECKER_COMMIT_SHA',
    'TYPO_CHECKER_ARCHIVE_SHA256',
    'TYPO_CHECKER_SELF_TEST',
    'ACTIVE_JOB_CONFIG_PATH',
    'INDESIGN_EXPORT_SCRIPT_PATH',
    'INPUT_XLSX',
    'REUSE_EXPORT_JOB_PATH',
    'ALLOW_STANDALONE_INPUT',
    'OUTPUT_JSON',
    'OUTPUT_XLSX',
    'SHEET_NAME',
    'ROW_LIMIT'
)

try {
    $resolution = Get-LatestReleaseResolution
    $workspace = New-SafeWorkspace
    $archivePath = Join-Path $workspace 'release.zip'
    $extractPath = Join-Path $workspace 'release'
    [void](New-Item -ItemType Directory -Path $extractPath)
    $archiveUri = Get-ArchiveUri -CommitSha $resolution.CommitSha
    Write-Host "Latest stable release: $($resolution.Tag)"
    Write-Host "Immutable release commit: $($resolution.CommitSha)"
    Write-Host 'Downloading a fresh release archive from the canonical GitHub repository...'
    Invoke-WebRequest -Method Get -Uri $archiveUri -Headers @{ 'User-Agent' = 'MatthewLukePublishing-TypoChecker' } -UseBasicParsing -MaximumRedirection 5 -OutFile $archivePath
    if (-not (Test-Path -LiteralPath $archivePath -PathType Leaf) -or (Get-Item -LiteralPath $archivePath).Length -lt 1) {
        throw 'GitHub returned an empty or missing release archive.'
    }
    $archiveSha256 = (Get-FileHash -LiteralPath $archivePath -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($archiveSha256 -notmatch '^[A-F0-9]{64}$') { throw 'Could not calculate the release archive SHA-256.' }
    Expand-Archive -LiteralPath $archivePath -DestinationPath $extractPath
    $roots = @(Get-ChildItem -LiteralPath $extractPath -Directory)
    $topLevelFiles = @(Get-ChildItem -LiteralPath $extractPath -File)
    if ($roots.Count -ne 1 -or $topLevelFiles.Count -ne 0) {
        throw 'The GitHub release archive did not contain exactly one project root.'
    }
    $sourceRoot = $roots[0].FullName
    foreach ($relativePath in $RequiredFiles) {
        $requiredPath = Join-Path $sourceRoot $relativePath
        if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
            throw "The latest release is incomplete; missing '$relativePath'."
        }
    }
    $package = Get-Content -LiteralPath (Join-Path $sourceRoot 'package.json') -Raw | ConvertFrom-Json
    if ([string]$package.version -ne $resolution.Tag.Substring(1)) {
        throw "Release tag '$($resolution.Tag)' does not match package version '$($package.version)'."
    }
    $node = Get-NodeExecutable
    $savedEnvironment = Save-ProcessEnvironment -Names $environmentNames
    [Environment]::SetEnvironmentVariable('OPEN_INDESIGN_DOCUMENT', 'true', 'Process')
    [Environment]::SetEnvironmentVariable('USE_OPEN_INDESIGN_DOCUMENT', 'true', 'Process')
    [Environment]::SetEnvironmentVariable('CHECK_LANGUAGE', $Language.Trim(), 'Process')
    [Environment]::SetEnvironmentVariable('TYPO_CHECKER_RELEASE_TAG', $resolution.Tag, 'Process')
    [Environment]::SetEnvironmentVariable('TYPO_CHECKER_COMMIT_SHA', $resolution.CommitSha, 'Process')
    [Environment]::SetEnvironmentVariable('TYPO_CHECKER_ARCHIVE_SHA256', $archiveSha256, 'Process')
    foreach ($name in @('ACTIVE_JOB_CONFIG_PATH', 'INDESIGN_EXPORT_SCRIPT_PATH', 'INPUT_XLSX', 'REUSE_EXPORT_JOB_PATH', 'ALLOW_STANDALONE_INPUT', 'OUTPUT_JSON', 'OUTPUT_XLSX', 'SHEET_NAME')) {
        [Environment]::SetEnvironmentVariable($name, $null, 'Process')
    }
    [Environment]::SetEnvironmentVariable('ROW_LIMIT', '0', 'Process')

    Write-Host 'Running the release offline self-test (no InDesign or Codex call)...'
    [Environment]::SetEnvironmentVariable('TYPO_CHECKER_SELF_TEST', 'true', 'Process')
    & $node (Join-Path $sourceRoot 'Typo_Checker.js')
    $selfTestExit = $LASTEXITCODE
    [Environment]::SetEnvironmentVariable('TYPO_CHECKER_SELF_TEST', $null, 'Process')
    if ($selfTestExit -ne 0) { throw "The latest release failed its offline self-test with exit code $selfTestExit." }

    $currentResolution = Get-LatestReleaseResolution
    if ($currentResolution.Tag -ne $resolution.Tag -or $currentResolution.CommitSha -ne $resolution.CommitSha) {
        throw "The latest GitHub release changed from $($resolution.Tag) to $($currentResolution.Tag) during preparation. Run the launcher again so no stale release is executed."
    }

    Write-Host "Release archive SHA-256: $archiveSha256"
    Write-Host "Starting the exact latest release for language '$($Language.Trim())'."
    & $node (Join-Path $sourceRoot 'Typo_Checker.js')
    $checkerExit = $LASTEXITCODE
    if ($checkerExit -ne 0) { throw "Typo Checker failed with exit code $checkerExit." }
} finally {
    [Environment]::SetEnvironmentVariable('TYPO_CHECKER_SELF_TEST', $null, 'Process')
    if ($null -ne $savedEnvironment) { Restore-ProcessEnvironment -Saved $savedEnvironment }
    if ($null -ne $workspace) { Remove-SafeWorkspace -Path $workspace }
}
