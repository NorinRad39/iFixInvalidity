$dest = (Resolve-Path '..\Classe-outils-topsolid\Classe outils topsolid').Path
$destFile = Join-Path $dest 'docfx.json'

# Copier OutilsTs.dll seul (sans XML) dans un sous-dossier dédié pour docfx
$dllSrc = Join-Path $dest 'bin\x64\Release\net481\OutilsTs.dll'
$dllStageDir = Join-Path $dest 'docfx_stage'
if (Test-Path $dllStageDir) { Remove-Item $dllStageDir -Recurse -Force }
New-Item -ItemType Directory -Path $dllStageDir | Out-Null
Copy-Item $dllSrc $dllStageDir
Write-Host "DLL copiée dans : $dllStageDir (sans XML)"

$json = @'
{
  "metadata": [
	{
	  "src": [
		{ "files": [ "docfx_stage/OutilsTs.dll" ] }
	  ],
	  "references": [
		{
		  "files": [ "bin/x64/Release/net481/*.dll" ],
		  "exclude": [ "bin/x64/Release/net481/OutilsTs.dll" ]
		}
	  ],
	  "dest": "api",
	  "disableGitFeatures": true
	}
  ],
  "build": {
	"content": [
	  { "files": [ "api/**.yml", "api/index.md" ] },
	  { "files": [ "**/*.md" ] }
	],
	"dest": "_site_outilsts",
	"globalMetadata": {
	  "_appTitle": "OutilsTs Documentation"
	},
	"disableGitFeatures": true
  }
}
'@

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($destFile, $json, $utf8NoBom)
Write-Host "Ecrit: $destFile"
Get-Content -LiteralPath $destFile
