<#
.SYNOPSIS
    Publishes a MD-Word release: exports the public snapshot of local master
    into a temporary worktree of the public repo, commits it there, pushes
    fast-forward, and creates the GitHub release with the installer attached.

.DESCRIPTION
    Automates docs/RELEASING.md steps 3-9. By construction this script CANNOT
    perform the FORBIDDEN operations: it never runs "git push" from the local
    repository (only "git -C <worktree> push origin HEAD:master", which is
    fast-forward-only because git rejects anything else and --force is never
    passed), and it never pushes tags (the public tag is created by
    "gh release create" server-side).

    Preconditions checked before touching anything:
      * current branch is master, working tree clean;
      * <Version> read from Directory.Build.props;
      * local tag v<Version> either missing (created at HEAD) or already at HEAD;
      * out\MD-Word-Setup-<Version>.exe exists (or -BuildInstaller);
      * -NotesFile exists (unless -SkipRelease or -DryRun);
      * gh CLI present (unless -SkipRelease or -DryRun).

    Safety guards:
      * refuses to run when the worktree path already exists (stale run);
      * after staging, aborts if any Added/Modified path is internal-only
        (docs/, START-HERE.md, .superpowers/) — staged DELETIONS of such
        paths are legitimate (that is how a past leak gets removed);
      * an explicit 'publish' confirmation is required before the push;
      * -DryRun stops after staging and reports what WOULD be committed.

.PARAMETER Summary
    One-line summary for the public squash commit:
    "MD-Word v<Version>: <Summary>". Required unless -DryRun.

.PARAMETER NotesFile
    Markdown file used as the GitHub release body. The installer's SHA-256
    is appended automatically. Required unless -DryRun or -SkipRelease.

.PARAMETER BuildInstaller
    Run installer\build-installer.ps1 first instead of requiring
    out\MD-Word-Setup-<Version>.exe to already exist.

.PARAMETER SkipRelease
    Stop after the push and print the exact "gh release create" command to
    run manually (for machines without an authenticated gh).

.PARAMETER DryRun
    Rehearsal: fetch + worktree + export + stage + guards + report, then
    clean up. No commit, no push, no release. Only needs a clean master.

.EXAMPLE
    powershell -NoProfile -File tools\publish-release.ps1 -DryRun

.EXAMPLE
    powershell -NoProfile -File tools\publish-release.ps1 -BuildInstaller `
        -Summary "subscript/superscript/highlight/underline, CI" `
        -NotesFile docs\release-notes\v1.0.2.md
#>

[CmdletBinding()]
param(
    [string]$Summary,
    [string]$NotesFile,
    [switch]$BuildInstaller,
    [switch]$SkipRelease,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$worktreePath = Join-Path (Split-Path -Parent $repoRoot) 'mdword-publish'
$publicRepo = 'PysmMax/MD-Word'
$noreplyEmail = '252583966+PysmMax@users.noreply.github.com'
$noreplyName = 'PysmMax'

# Paths exported to the public snapshot. docs/ and START-HERE.md are
# deliberately absent (internal); .github ships the CI workflow. Keep this
# list in sync with docs/RELEASING.md's "What does NOT go into the public
# snapshot" section.
$publicPaths = @(
    'README.md', 'LICENSE', 'THIRD-PARTY-NOTICES.md', '.gitignore',
    '.github', 'Directory.Build.props', 'MdWord.sln',
    'src', 'tests', 'tools', 'installer', 'samples'
)

function Invoke-Git {
    param([string[]]$GitArgs)
    & git @GitArgs
    if ($LASTEXITCODE -ne 0) {
        throw "git $($GitArgs -join ' ') exited with code $LASTEXITCODE."
    }
}

function Get-GitOutput {
    param([string[]]$GitArgs)
    $output = & git @GitArgs
    if ($LASTEXITCODE -ne 0) {
        throw "git $($GitArgs -join ' ') exited with code $LASTEXITCODE."
    }
    return $output
}

Push-Location $repoRoot
try {
    # --- Preconditions -------------------------------------------------------
    $branch = (Get-GitOutput @('rev-parse', '--abbrev-ref', 'HEAD') | Select-Object -First 1)
    if ($branch -ne 'master') {
        throw "Must run on master (current branch: $branch). Merge the release branch first."
    }

    $dirty = Get-GitOutput @('status', '--porcelain')
    if ($dirty) {
        throw "Working tree is not clean:`n$($dirty -join "`n")"
    }

    [xml]$props = Get-Content -LiteralPath (Join-Path $repoRoot 'Directory.Build.props')
    $version = ($props.Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1).ToString().Trim()
    if (-not $version) {
        throw 'Could not read <Version> from Directory.Build.props.'
    }
    $tagName = "v$version"
    Write-Host "Version: $version (tag $tagName)" -ForegroundColor Cyan

    $installerPath = Join-Path $repoRoot "out\MD-Word-Setup-$version.exe"

    if (-not $DryRun) {
        if (-not $Summary) {
            throw '-Summary is required (one-line public commit message).'
        }

        $headSha = (Get-GitOutput @('rev-parse', 'HEAD') | Select-Object -First 1)
        & git rev-parse -q --verify "refs/tags/$tagName" | Out-Null
        if ($LASTEXITCODE -eq 0) {
            $tagSha = (Get-GitOutput @('rev-parse', "$tagName^{commit}") | Select-Object -First 1)
            if ($tagSha -ne $headSha) {
                throw "Local tag $tagName exists but does not point at HEAD ($tagSha vs $headSha). Fix the tag or the checkout first."
            }
        } else {
            Write-Host "Creating local tag $tagName at HEAD (local only — never pushed)." -ForegroundColor Cyan
            Invoke-Git @('tag', $tagName)
        }

        if ($BuildInstaller) {
            & powershell -NoProfile -File (Join-Path $repoRoot 'installer\build-installer.ps1')
            if ($LASTEXITCODE -ne 0) {
                throw "build-installer.ps1 exited with code $LASTEXITCODE."
            }
        }
        if (-not (Test-Path -LiteralPath $installerPath)) {
            throw "Installer not found: $installerPath. Run installer\build-installer.ps1 or pass -BuildInstaller."
        }

        if (-not $SkipRelease) {
            if (-not $NotesFile) {
                throw '-NotesFile is required (release notes body), or pass -SkipRelease.'
            }
            if (-not (Test-Path -LiteralPath $NotesFile)) {
                throw "Notes file not found: $NotesFile"
            }
            $gh = Get-Command 'gh' -ErrorAction SilentlyContinue
            if (-not $gh) {
                throw 'gh CLI not found. Install GitHub CLI or pass -SkipRelease.'
            }
        }
    }

    if (Test-Path -LiteralPath $worktreePath) {
        throw "Worktree path already exists: $worktreePath. A previous run failed? Inspect it, then remove with: git worktree remove --force `"$worktreePath`""
    }

    # --- Export the snapshot into a temporary worktree -----------------------
    Invoke-Git @('fetch', 'origin', 'master')
    Invoke-Git @('worktree', 'add', $worktreePath, 'origin/master')

    try {
        # Replace the payload wholesale: staging deletions first means files
        # the new snapshot no longer ships arrive as clean 'git rm's.
        Invoke-Git @('-C', $worktreePath, 'rm', '-r', '-q', '.')

        $tarPath = Join-Path $env:TEMP ("mdword-publish-payload-" + [guid]::NewGuid().ToString('N') + ".tar")
        $archiveArgs = @('archive', "--output=$tarPath", 'master') + $publicPaths
        Invoke-Git $archiveArgs

        & tar -xf $tarPath -C $worktreePath
        if ($LASTEXITCODE -ne 0) {
            throw "tar -xf exited with code $LASTEXITCODE."
        }
        Remove-Item -LiteralPath $tarPath

        Invoke-Git @('-C', $worktreePath, 'add', '-A')

        # --- Leak guard ------------------------------------------------------
        $addedOrModified = Get-GitOutput @('-C', $worktreePath, 'diff', '--cached', '--name-only', '--diff-filter=AM')
        $suspicious = @($addedOrModified | Where-Object { $_ -match '^(docs/|START-HERE\.md$|\.superpowers/)' })
        if ($suspicious.Count -gt 0) {
            throw "LEAK GUARD: internal paths would be published: $($suspicious -join ', '). Aborting."
        }

        Write-Host "`nSnapshot diff (public repo <- local master):" -ForegroundColor Cyan
        & git -C $worktreePath diff --cached --stat
        if ($LASTEXITCODE -ne 0) {
            throw 'git diff --cached --stat failed.'
        }

        if ($DryRun) {
            Write-Host "`nDRY RUN — stopping before commit/push/release. Nothing was published." -ForegroundColor Yellow
            return
        }

        & git -C $worktreePath diff --cached --quiet
        if ($LASTEXITCODE -eq 0) {
            throw 'Nothing to publish: the snapshot is identical to the public repo.'
        }

        # --- Confirm, commit, push -------------------------------------------
        $answer = Read-Host "Type 'publish' to push this snapshot to the PUBLIC repo ($publicRepo)"
        if ($answer -ne 'publish') {
            throw 'Aborted by user (confirmation not given).'
        }

        $commitMessage = "MD-Word v${version}: $Summary"
        Invoke-Git @('-C', $worktreePath, '-c', "user.email=$noreplyEmail", '-c', "user.name=$noreplyName", 'commit', '-m', $commitMessage)
        Invoke-Git @('-C', $worktreePath, 'push', 'origin', 'HEAD:master')
        Write-Host "Pushed public snapshot: $commitMessage" -ForegroundColor Green

        # --- GitHub release ----------------------------------------------------
        $hash = (Get-FileHash -LiteralPath $installerPath -Algorithm SHA256).Hash.ToLowerInvariant()
        $installerName = [System.IO.Path]::GetFileName($installerPath)
        Write-Host "SHA-256 ($installerName): $hash" -ForegroundColor Cyan

        if ($SkipRelease) {
            Write-Host "`n-SkipRelease: create the release manually:" -ForegroundColor Yellow
            Write-Host "gh release create $tagName `"$installerPath`" -R $publicRepo --title `"MD-Word $tagName`" --notes-file `"<notes-with-sha256>`""
            return
        }

        $combinedNotes = Join-Path $env:TEMP ("mdword-release-notes-" + [guid]::NewGuid().ToString('N') + ".md")
        $notesBody = Get-Content -LiteralPath $NotesFile -Raw
        $checksumLine = "`n`n**SHA-256** ``$installerName``:`n``````text`n$hash`n``````"
        Set-Content -LiteralPath $combinedNotes -Value ($notesBody + $checksumLine) -Encoding UTF8

        & gh release create $tagName $installerPath -R $publicRepo --title "MD-Word $tagName" --notes-file $combinedNotes
        if ($LASTEXITCODE -ne 0) {
            Remove-Item -LiteralPath $combinedNotes -ErrorAction SilentlyContinue
            throw "gh release create exited with code $LASTEXITCODE. The snapshot IS already pushed; fix gh auth and create the release manually."
        }
        Remove-Item -LiteralPath $combinedNotes

        Write-Host "`nRelease $tagName published: https://github.com/$publicRepo/releases/tag/$tagName" -ForegroundColor Green
    }
    finally {
        if (Test-Path -LiteralPath $worktreePath) {
            & git worktree remove --force $worktreePath
            if ($LASTEXITCODE -ne 0) {
                Write-Warning "Could not remove the worktree — remove manually: git worktree remove --force `"$worktreePath`""
            }
        }
    }
}
finally {
    Pop-Location
}