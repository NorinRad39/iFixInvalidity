<#
Script PowerShell : télécharge et installe docfx (binaire) dans %USERPROFILE%\.dotnet\tools
- Télécharge https://github.com/dotnet/docfx/releases/latest/download/docfx.zip
- Extrait docfx.exe et le place dans %USERPROFILE%\.dotnet\tools
- Ajoute le dossier au PATH utilisateur (via setx) si nécessaire
- Met à jour la session courante PATH
- Vérifie l'installation et, si demandé, lance generate-docs.ps1

Usage :
  .\install-docfx-automatic.ps1            # installe docfx
  .\install-docfx-automatic.ps1 -RunGenerate # installe puis lance generate-docs.ps1
#>
param(
	[switch]$RunGenerate
)

function Write-Ok($m){ Write-Host "[OK] $m" -ForegroundColor Green }
function Write-Err($m){ Write-Host "[ERR] $m" -ForegroundColor Red }
function Write-Inf($m){ Write-Host "[..] $m" -ForegroundColor Cyan }

$installDir = Join-Path $env:USERPROFILE ".dotnet\tools"
if (-not (Test-Path $installDir)) {
	New-Item -ItemType Directory -Path $installDir -Force | Out-Null
}

# Check existing
$existing = Get-Command docfx -ErrorAction SilentlyContinue
if ($existing) {
	Write-Ok "docfx déjà disponible : $($existing.Source)"
	if ($RunGenerate) {
		Write-Inf "Lancement de generate-docs.ps1"
		pwsh -NoProfile -ExecutionPolicy Bypass -File .\generate-docs.ps1
	}
	return
}

# If docfx.exe already in installDir
$localExe = Join-Path $installDir 'docfx.exe'
if (Test-Path $localExe) {
	Write-Ok "docfx.exe trouvé dans $installDir"
	$env:PATH = "$env:PATH;$installDir"
	if ($RunGenerate) {
		Write-Inf "Lancement de generate-docs.ps1"
		pwsh -NoProfile -ExecutionPolicy Bypass -File .\generate-docs.ps1
	}
	return
}

# Try Chocolatey first if available
if (Get-Command choco -ErrorAction SilentlyContinue) {
	Write-Inf "Chocolatey détecté : tentative d'installation via choco install docfx -y"
	try {
		& choco install docfx -y
		if ($LASTEXITCODE -eq 0) {
			Write-Ok "docfx installé via Chocolatey"
			# attempt to locate docfx
			$existing = Get-Command docfx -ErrorAction SilentlyContinue
			if ($existing) { Write-Ok "docfx disponible : $($existing.Source)" }
			if ($RunGenerate) { pwsh -NoProfile -ExecutionPolicy Bypass -File .\generate-docs.ps1 }
			return
		}
	}
	catch {
		Write-Err "Installation via Chocolatey échouée : $($_.Exception.Message)"
	}
}

# Query GitHub Releases API to find the latest docfx zip asset
$api = 'https://api.github.com/repos/dotnet/docfx/releases/latest'
$tmp = Join-Path $env:TEMP "docfx_latest.zip"
if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
try {
	Write-Inf "Interrogation de l'API GitHub pour localiser l'artefact docfx..."
	$resp = Invoke-RestMethod -Uri $api -Headers @{ 'User-Agent' = 'install-docfx-script' } -ErrorAction Stop
	$asset = $resp.assets | Where-Object { $_.name -match 'docfx.*zip' -or $_.name -match 'docfx.zip' } | Select-Object -First 1
	if ($null -eq $asset) {
		Write-Err "Aucun artefact docfx.zip trouvé dans la release la plus récente."
		throw "AssetNotFound"
	}
	$downloadUrl = $asset.browser_download_url
	Write-Inf "Téléchargement depuis $downloadUrl"
	Invoke-WebRequest -Uri $downloadUrl -OutFile $tmp -UseBasicParsing -ErrorAction Stop
	Write-Ok "Téléchargement terminé : $tmp"
}
catch {
	Write-Err "Échec lors de la localisation/téléchargement de docfx via GitHub API : $($_.Exception.Message)"
	Write-Host "Vous pouvez installer docfx manuellement via Chocolatey (choco install docfx -y) ou télécharger l'archive depuis la page des releases GitHub." -ForegroundColor Yellow
	exit 1
}

# Extract to temp folder
$extractDir = Join-Path $env:TEMP "docfx_extracted_$(Get-Random)"
if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Path $extractDir | Out-Null

try {
	Expand-Archive -Path $tmp -DestinationPath $extractDir -Force
	Write-Ok "Archive extraite dans $extractDir"
}
catch {
	Write-Err "Impossible d'extraire l'archive : $($_.Exception.Message)"
	exit 1
}

# Find docfx.exe
$found = Get-ChildItem -Path $extractDir -Recurse -Filter 'docfx.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $found) {
	Write-Err "docfx.exe introuvable dans l'archive extraite"
	exit 1
}

try {
	Copy-Item -Path $found.FullName -Destination $localExe -Force
	Write-Ok "docfx.exe copié vers $localExe"
}
catch {
	Write-Err "Échec de la copie : $($_.Exception.Message)"
	exit 1
}

# Add installDir to user PATH if missing
$currentUserPath = [Environment]::GetEnvironmentVariable('Path', 'User')
if (-not ($currentUserPath -split ';' | Where-Object { $_ -eq $installDir })) {
	try {
		$newUserPath = if ([string]::IsNullOrEmpty($currentUserPath)) { $installDir } else { "$currentUserPath;$installDir" }
		[Environment]::SetEnvironmentVariable('Path', $newUserPath, 'User')
		Write-Ok "Ajout de $installDir au PATH utilisateur (persistant)."
	}
	catch {
		Write-Err "Impossible de modifier le PATH utilisateur : $($_.Exception.Message)"
	}
}
else {
	Write-Inf "$installDir déjà présent dans le PATH utilisateur"
}

# Update current session PATH
if (-not ($env:PATH -split ';' | Where-Object { $_ -eq $installDir })) {
	$env:PATH = "$env:PATH;$installDir"
}

# Verify
try {
	$version = & docfx --version 2>&1
	Write-Ok "docfx installé : $version"
}
catch {
	Write-Err "Impossible d'exécuter docfx : $($_.Exception.Message)"
	exit 1
}

# Optionally run generation
if ($RunGenerate) {
	Write-Inf "Lancement de generate-docs.ps1"
	pwsh -NoProfile -ExecutionPolicy Bypass -File .\generate-docs.ps1
}

Write-Host "Opération terminée." -ForegroundColor Green
