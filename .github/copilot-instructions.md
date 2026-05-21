<Project>
  <PropertyGroup>
    <!-- Pour projets non-SDK, DocumentationFile est pris en compte ; pour SDK-style on utiliserait GenerateDocumentationFile -->
    <DocumentationFile>bin\$(Configuration)\$(AssemblyName).xml</DocumentationFile>
  </PropertyGroup>
</Project>

---

## Pipeline DocFX — documentation automatique

### Fichiers clés (dans chaque repo)

| Fichier | Repo | Rôle |
|---|---|---|
| `generate-docs.ps1` | iFixInvalidity (racine) | Script principal de génération DocFX |
| `start-docs.ps1` | iFixInvalidity (racine) | Lance les serveurs HTTP + ouvre le navigateur |
| `start-docs.ps1` | Classe-outils-topsolid (racine) | Idem pour l'autre repo |
| `docfx.json` | iFixInvalidity (racine) | Config DocFX pour iFixInvalidity |
| `toc.yml` + `index.md` | iFixInvalidity (racine) | Menu et page d'accueil du site iFixInvalidity |
| `Directory.Build.targets` | iFixInvalidity (racine) | Déclenche auto la doc après chaque build VS |
| `install-docfx-from-nuget.ps1` | iFixInvalidity (racine) | Installe DocFX localement (une seule fois) |
| `Classe outils topsolid/docfx.json` | Classe-outils-topsolid | Config DocFX pour OutilsTs |
| `Classe outils topsolid/toc.yml` + `index.md` | Classe-outils-topsolid | Menu et page d'accueil du site OutilsTs |
| `Directory.Build.targets` | Classe-outils-topsolid (racine) | Déclenche auto la doc après chaque build VS |

### Dossiers générés (exclus de Git — recréés automatiquement après build)

- `iFixInvalidity\_site_ifix\` → site iFixInvalidity
- `Classe-outils-topsolid\Classe outils topsolid\_site_outilsts\` → site OutilsTs
- `docfx_stage_ifix\`, `Classe outils topsolid\docfx_stage\` → staging binaires sans .xml
- `api\` → métadonnées YAML générées par DocFX

### Procédure sur une nouvelle machine

```powershell
# 1. Cloner les deux repos côte à côte (même dossier parent)
git clone https://github.com/NorinRad39/iFixInvalidity
git clone https://github.com/NorinRad39/Classe-outils-topsolid

# 2. Installer DocFX (une seule fois par machine)
cd iFixInvalidity
pwsh -ExecutionPolicy Bypass -File .\install-docfx-from-nuget.ps1

# 3. Ouvrir iFixInvalidity.sln dans Visual Studio et builder (Build Solution)
#    → la documentation se génère automatiquement en arrière-plan après chaque build

# 4. Ouvrir la documentation dans le navigateur
pwsh -File .\start-docs.ps1
```

### Ouvrir la doc manuellement (sans builder)

```powershell
cd C:\...\iFixInvalidity

# Générer pour les deux projets
.\generate-docs.ps1 -Configuration Debug -Platform x64

# Ou un seul projet
.\generate-docs.ps1 -Project iFixInvalidity -Configuration Debug -Platform AnyCPU
.\generate-docs.ps1 -Project OutilsTs -Configuration Release -Platform x64

# Puis ouvrir dans le navigateur
.\start-docs.ps1
```

### URLs des sites (via start-docs.ps1)

- **iFixInvalidity** → http://localhost:8080
- **OutilsTs** → http://localhost:8081

> ⚠️ Ne pas ouvrir les `index.html` directement depuis l'explorateur : le menu JS ne s'affiche pas en `file://`.
> Toujours passer par `start-docs.ps1` ou `docfx serve`.

### Désactiver temporairement la génération auto

```powershell
# Dans MSBuild (ligne de commande)
msbuild iFixInvalidity.sln /p:GenerateDocFXEnabled=false
```