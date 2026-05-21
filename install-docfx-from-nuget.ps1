<#
Télécharge le package NuGet docfx.console (dernier) et extrait docfx.exe vers %USERPROFILE%\.dotnet\tools
Usage: .\install-docfx-from-nuget.ps1 [-RunGenerate]
#>
param([switch]$RunGenerate)

$installDir = Join-Path $env:USERPROFILE ".dotnet\tools"
if (-not (Test-Path $installDir)) { New-Item -ItemType Directory -Path $installDir -Force | Out-Null }

$tmpNupkg = Join-Path $env:TEMP "docfx_latest.nupkg"
if (Test-Path $tmpNupkg) { Remove-Item $tmpNupkg -Force -ErrorAction SilentlyContinue }

$nugetUrl = 'https://www.nuget.org/api/v2/package/docfx.console'
Write-Host "Téléchargement du package NuGet docfx.console depuis $nugetUrl"
try {
	Invoke-WebRequest -Uri $nugetUrl -OutFile $tmpNupkg -UseBasicParsing -ErrorAction Stop
}
catch {
	Write-Error "Échec du téléchargement du package NuGet : $($_.Exception.Message)"
	exit 1
}

$extractDir = Join-Path $env:TEMP "docfx_nupkg_$(Get-Random)"
if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Path $extractDir | Out-Null

try {
	Expand-Archive -Path $tmpNupkg -DestinationPath $extractDir -Force
}
catch {
	Write-Error "Impossible d'extraire le nupkg : $($_.Exception.Message)"
	exit 1
}

# Rechercher docfx.exe
$found = Get-ChildItem -Path $extractDir -Recurse -Filter 'docfx.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $found) {
	Write-Error "docfx.exe introuvable dans le package NuGet extrait. Contenu du dossier extrait :"
	Get-ChildItem -Path $extractDir -Recurse | Select-Object FullName | ForEach-Object { Write-Host $_.FullName }
	exit 1
}

$dest = Join-Path $installDir 'docfx.exe'
try {
	# Copy entire folder containing docfx.exe (includes required dlls) to the install dir
	$sourceDir = Split-Path $found.FullName -Parent
	Copy-Item -Path (Join-Path $sourceDir '*') -Destination $installDir -Recurse -Force
	Write-Host "Fichiers DocFX copiés depuis $sourceDir vers $installDir"
}
catch {
	Write-Error "Échec de la copie : $($_.Exception.Message)"
	exit 1
}

# Mettre à jour PATH utilisateur
$currentUserPath = [Environment]::GetEnvironmentVariable('Path','User')
if (-not ($currentUserPath -split ';' | Where-Object { $_ -eq $installDir })) {
	try {
		$newUserPath = if ([string]::IsNullOrEmpty($currentUserPath)) { $installDir } else { "$currentUserPath;$installDir" }
		[Environment]::SetEnvironmentVariable('Path',$newUserPath,'User')
		Write-Host "Ajout de $installDir au PATH utilisateur"
	}
	catch { Write-Warning "Impossible de modifier le PATH utilisateur: $($_.Exception.Message)" }
}
# Mettre à jour session courante
if (-not ($env:PATH -split ';' | Where-Object { $_ -eq $installDir })) { $env:PATH = "$env:PATH;$installDir" }

# Vérifier
try { $v = & docfx --version 2>&1; Write-Host "docfx installé : $v" } catch { Write-Warning "Impossible d'exécuter docfx : $($_.Exception.Message)" }

if ($RunGenerate) { pwsh -NoProfile -ExecutionPolicy Bypass -File .\generate-docs.ps1 }

Write-Host "Terminé."