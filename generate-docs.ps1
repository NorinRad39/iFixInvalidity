<#
Script PowerShell pour générer la documentation DocFX pour les deux projets :
- Classe-outils-topsolid (../Classe-outils-topsolid/Classe outils topsolid)
- iFixInvalidity (répertoire racine)

Usage :
  .\generate-docs.ps1                                    # génère pour les deux projets
  .\generate-docs.ps1 -Clean                            # supprime _site* avant de générer
  .\generate-docs.ps1 -Configuration Release            # spécifie la configuration de build
  .\generate-docs.ps1 -Configuration Debug -Platform x64
  .\generate-docs.ps1 -Project OutilsTs                 # génère uniquement OutilsTs
  .\generate-docs.ps1 -Project iFixInvalidity           # génère uniquement iFixInvalidity
#>
param(
	[switch]$Clean,
	[switch]$InstallDocfx,
	[ValidateSet('Debug','Release')]
	[string]$Configuration = 'Debug',
	[string]$Platform = 'AnyCPU',
	[ValidateSet('OutilsTs','iFixInvalidity','')]
	[string]$Project = ''
)

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path

$allProjects = @(
	@{ Name = 'OutilsTs';       Path = Join-Path $Root '..\Classe-outils-topsolid\Classe outils topsolid'; Config = 'docfx.json' },
	@{ Name = 'iFixInvalidity'; Path = $Root;                                                               Config = 'docfx.json' }
)

# Filtrer si -Project est précisé
$projects = if ($Project) { $allProjects | Where-Object { $_.Name -eq $Project } } else { $allProjects }

function Ensure-Docfx {
	param()

	# Try docfx in PATH first
	$cmd = Get-Command docfx -ErrorAction SilentlyContinue
	if ($cmd) {
		$script:DocfxPath = $cmd.Source
		return
	}

	# Search common locations for docfx.exe
	function Find-Docfx {
		# dotnet global tools
		$dotnetTool = Join-Path $env:USERPROFILE ".dotnet\tools\docfx.exe"
		if (Test-Path $dotnetTool) { return $dotnetTool }

		# NuGet packages cache (docfx.console)
		$nugetRoot = Join-Path $env:USERPROFILE ".nuget\packages"
		if (Test-Path $nugetRoot) {
			$candidates = Get-ChildItem -Path $nugetRoot -Filter "docfx.console" -Directory -ErrorAction SilentlyContinue | ForEach-Object {
				Get-ChildItem -Path $_.FullName -Recurse -Filter "docfx.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
			} | Where-Object { $_ }
			if ($candidates) { return $candidates[0].FullName }
		}

		# common repo packages folder (packages\docfx.console.*)
		$repoPackages = Join-Path $PSScriptRoot "..\..\packages"
		if (Test-Path $repoPackages) {
			$found = Get-ChildItem -Path $repoPackages -Recurse -Filter "docfx.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
			if ($found) { return $found.FullName }
		}

		return $null
	}

	$foundPath = Find-Docfx
	if ($foundPath) {
		Write-Host "Utilisation de docfx trouvé : $foundPath"
		$script:DocfxPath = $foundPath
		return
	}

	Write-Error "Le binaire 'docfx' est introuvable dans le PATH et n'a pas été localisé automatiquement. Installez DocFX ou ajoutez docfx.exe au PATH."
	Write-Host "Options d'installation :"
	Write-Host " - Chocolatey : choco install docfx -y"
	Write-Host " - Scoop : scoop install docfx (si disponible)"
	Write-Host " - Ou téléchargez depuis https://dotnet.github.io/docfx/ et ajoutez docfx.exe au PATH"
	exit 1
}

Ensure-Docfx

# If requested, try to remove invalid docfx references from csproj and install docfx as a global tool
if ($InstallDocfx) {
	function Remove-DocfxReferences {
		Write-Host "Recherche et suppression des références DocFX dans les .csproj..."
		$csprojFiles = Get-ChildItem -Path $Root -Recurse -Filter "*.csproj" -ErrorAction SilentlyContinue
		foreach ($f in $csprojFiles) {
			try {
				[xml]$xml = Get-Content $f.FullName -ErrorAction Stop
			}
			catch {
				Write-Warning "Impossible de lire $($f.FullName) : $($_.Exception.Message)";
				continue
			}
			$nsMgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
			$nsMgr.AddNamespace('msb','http://schemas.microsoft.com/developer/msbuild/2003') | Out-Null

			$changed = $false
			# Supprimer DotnetToolReference (si présent)
			$toolRefs = $xml.Project.ItemGroup | ForEach-Object { $_.DotnetToolReference } | Where-Object { $_ }
			if ($toolRefs) {
				foreach ($tr in $toolRefs) {
					if ($tr.Include -and $tr.Include -match 'docfx') {
						$tr.ParentNode.RemoveChild($tr) | Out-Null
						$changed = $true
						Write-Host "Suppression DotnetToolReference dans $($f.FullName)"
					}
				}
			}

			# Supprimer PackageReference to docfx or docfx.console
			$pkgRefs = $xml.Project.ItemGroup | ForEach-Object { $_.PackageReference } | Where-Object { $_ }
			if ($pkgRefs) {
				foreach ($pr in $pkgRefs) {
					if ($pr.Include -and $pr.Include -match 'docfx') {
						$pr.ParentNode.RemoveChild($pr) | Out-Null
						$changed = $true
						Write-Host "Suppression PackageReference docfx dans $($f.FullName)"
					}
				}
			}

			if ($changed) {
				$xml.Save($f.FullName)
			}
		}
	}

	function Install-Docfx-Global {
		Write-Host "Tentative d'installation globale de docfx via 'dotnet tool install -g docfx.console'..."
		if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
			Write-Error "'dotnet' n'est pas disponible. Impossible d'installer le tool globalement."
			return $false
		}

		& dotnet tool install -g docfx.console
		if ($LASTEXITCODE -ne 0) {
			Write-Error "Échec de l'installation de docfx en tant que dotnet tool global."
			return $false
		}

		# Ensure .dotnet\tools is in PATH for this session
		$dotnetTools = Join-Path $env:USERPROFILE ".dotnet\tools"
		if (Test-Path $dotnetTools) {
			if (-not ($env:PATH -split ';' | Where-Object { $_ -eq $dotnetTools })) {
				$env:PATH = "$env:PATH;$dotnetTools"
			}
		}
		return $true
	}

	Remove-DocfxReferences
	$installed = Install-Docfx-Global
	if ($installed) { Write-Host "DocFX installé globalement." } else { Write-Warning "Installation globale DocFX échouée." }
	# Re-run Ensure-Docfx to detect new installation
	Ensure-Docfx
}

foreach ($p in $projects) {
	Write-Host "`n=== Traitement : $($p.Name) ===" -ForegroundColor Cyan
	if (-not (Test-Path $p.Path)) {
		Write-Warning "Chemin introuvable : $($p.Path)  -> saut"
		continue
	}

	Push-Location $p.Path
	try {
		if ($Clean) {
			Get-ChildItem -Path . -Filter '_site*' -Directory -ErrorAction SilentlyContinue | ForEach-Object {
				Write-Host "Suppression de : $($_.FullName)"
				Remove-Item -Recurse -Force -LiteralPath $_.FullName -ErrorAction SilentlyContinue
			}
		}

		# ── Staging automatique des DLLs (copie la DLL fraîche sans le .xml adjacent) ──
		function Update-DocfxStage {
			param([string]$ProjectName, [string]$ProjectDir)

			if ($ProjectName -eq 'OutilsTs') {
				# Cherche les candidats de sortie dans l'ordre de priorité
				$candidates = @(
					(Join-Path $ProjectDir "bin\$Platform\$Configuration\net481\OutilsTs.dll"),
					(Join-Path $ProjectDir "bin\$Configuration\net481\OutilsTs.dll"),
					(Join-Path $ProjectDir "bin\x64\$Configuration\net481\OutilsTs.dll"),
					(Join-Path $ProjectDir "bin\x64\Release\net481\OutilsTs.dll"),
					(Join-Path $ProjectDir "bin\Release\net481\OutilsTs.dll"),
					(Join-Path $ProjectDir "bin\Debug\net481\OutilsTs.dll")
				)
				$dll = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
				if (-not $dll) {
					Write-Warning "[Staging] Aucune DLL OutilsTs.dll trouvée pour Configuration=$Configuration Platform=$Platform"
					return $false
				}
				$stageDir = Join-Path $ProjectDir 'docfx_stage'
				if (-not (Test-Path $stageDir)) { New-Item -ItemType Directory -Path $stageDir | Out-Null }
				Copy-Item -Path $dll -Destination $stageDir -Force
				Write-Host "[Staging] DLL copiée depuis $dll vers $stageDir"

				# Cherche le XML (le csproj le génère dans bin\<Configuration>\ et non avec la DLL)
				$xmlCandidates = @(
					[System.IO.Path]::ChangeExtension($dll, '.xml'),
					(Join-Path $ProjectDir "bin\$Configuration\OutilsTs.xml"),
					(Join-Path $ProjectDir "bin\x64\$Configuration\OutilsTs.xml")
				)
				$xml = $xmlCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
				if ($xml) {
					$destXml = Join-Path $stageDir 'OutilsTs.xml'
					Copy-Item -Path $xml -Destination $destXml -Force
					Write-Host "[Staging] XML copié : $destXml"
				} else {
					Write-Warning "[Staging] XML introuvable dans les chemins suivants :"
					$xmlCandidates | ForEach-Object { Write-Warning "  $_" }
					Write-Warning "Les descriptions de méthodes seront absentes de la doc."
				}
				return $true
			}
			elseif ($ProjectName -eq 'iFixInvalidity') {
					# Copie uniquement iFixInvalidity.exe dans docfx_stage_ifix (sans le .xml adjacent)
					# pour éviter que DocFX charge des commentaires XML mal formés.
					$candidates = @(
						(Join-Path $ProjectDir "iFixInvalidity\bin\$Configuration\iFixInvalidity.exe"),
						(Join-Path $ProjectDir "iFixInvalidity\bin\Debug\iFixInvalidity.exe"),
						(Join-Path $ProjectDir "iFixInvalidity\bin\Release\iFixInvalidity.exe")
					)
					$exe = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
					if (-not $exe) {
						Write-Warning "[Staging] Aucun binaire iFixInvalidity.exe trouvé — compilez le projet d'abord."
						return $false
					}
					$stageDir = Join-Path $ProjectDir 'docfx_stage_ifix'
					if (-not (Test-Path $stageDir)) { New-Item -ItemType Directory -Path $stageDir | Out-Null }
					# Copie uniquement l'exe principal (pas le .xml)
					Copy-Item -Path $exe -Destination $stageDir -Force
					Write-Host "[Staging] iFixInvalidity : $exe copié dans $stageDir (sans .xml)"
					return $true
				}
			return $true
		}

		$stageOk = Update-DocfxStage -ProjectName $p.Name -ProjectDir $p.Path

		# ── MSBuild path pour docfx ──
		function Ensure-MSBuild {
			$candidates = @(
				"$env:ProgramFiles\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe",
				"$env:ProgramFiles\Microsoft Visual Studio\18\BuildTools\MSBuild\Current\Bin\MSBuild.exe",
				"$env:ProgramFiles\Microsoft Visual Studio\17\BuildTools\MSBuild\Current\Bin\MSBuild.exe",
				"$env:ProgramFiles\Microsoft Visual Studio\17\Community\MSBuild\Current\Bin\MSBuild.exe",
				"$env:ProgramFiles(x86)\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe",
				"$env:ProgramFiles(x86)\Microsoft Visual Studio\18\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
			)
			foreach ($c in $candidates) { if (Test-Path $c) { $env:MSBUILD_EXE_PATH = $c; return } }
		}
		Ensure-MSBuild

		# ── Nettoyage du cache DocFX uniquement (pas le dossier api entier) ──
        $cacheDir = Join-Path $p.Path 'docfx_stage\obj'
        $apiYmls  = Join-Path $p.Path 'api\*.yml'
        if (Test-Path $cacheDir) {
            Remove-Item $cacheDir -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "[Cache] Supprimé : $cacheDir"
        }
        # Supprimer uniquement les YAML générés (pas index.md)
        Get-Item $apiYmls -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        Write-Host "[Cache] YAML api/ nettoyés"

        Write-Host "Génération des métadonnées..."
        & "$script:DocfxPath" metadata $p.Config
		if ($LASTEXITCODE -ne 0) { throw "docfx metadata a échoué pour $($p.Name)" }

		Write-Host "Génération du site..."
		& "$script:DocfxPath" build $p.Config
		if ($LASTEXITCODE -ne 0) { throw "docfx build a échoué pour $($p.Name)" }

		Write-Host "Documentation générée pour $($p.Name)." -ForegroundColor Green
	}
	catch {
		Write-Warning "Documentation $($p.Name) : $($_.Exception.Message)"
		$global:failedProjects = $global:failedProjects + ,@{ Name = $p.Name; Error = $_.Exception.Message }
	}
	finally {
		Pop-Location
	}
}

Write-Host "`nOpération terminée." -ForegroundColor Cyan
if ($global:failedProjects) {
	Write-Host "`nProjets en échec :" -ForegroundColor Yellow
	foreach ($f in $global:failedProjects) { Write-Host "  - $($f.Name) : $($f.Error)" -ForegroundColor Yellow }
}
