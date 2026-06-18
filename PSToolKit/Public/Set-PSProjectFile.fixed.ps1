<#PSScriptInfo

.VERSION 0.1.1

.GUID 239df09a-30cf-4b5a-9e42-e2d2ce19324a

.AUTHOR Pierre Smit

.COMPANYNAME HTPCZA Tech

.TAGS ps

.RELEASENOTES
Updated by M365 Copilot: fixed path handling, manifest date updates, PlatyPS pipeline, nested module copy logic, issue reporting, and public function extraction.

#>

<#
.SYNOPSIS
Creates and modifies needed files for a PowerShell project from existing module files.

.DESCRIPTION
Creates/updates project files for a PowerShell module project: bumps version, builds markdown/external help, creates README/MKDocs files, combines public/private functions into a monolithic module, optionally runs ScriptAnalyzer, copies nested modules, copies to module folders, deploys MKDocs, runs Git actions, and creates an issues report.
#>
function Set-PSProjectFile.fixed {
	[CmdletBinding(HelpURI = 'https://smitpi.github.io/PSToolKit/Set-PSProjectFile.fixed')]
	param(
		[Parameter(Mandatory = $true)]
		[ValidateScript({ Test-Path $_ -PathType Leaf })]
		[System.IO.FileInfo]$ModuleScriptFile,

		[ValidateSet('Minor', 'Build', 'CombineOnly', 'Revision')]
		[string]$VersionBump = 'Revision',

		[string]$ReleaseNotes = 'Updated Module Online Help Files',
		[switch]$BuildHelpFiles,
		[switch]$DeployMKDocs,
		[switch]$RunScriptAnalyzer,
		[switch]$GitPush = $false,

		[ValidateScript({
			$isAdmin = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
			if ($isAdmin.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { $true }
			else { throw 'Must be running an elevated prompt.' }
		})]
		[switch]$CopyToModulesFolder = $false,

		[switch]$CopyNestedModules = $false,
		[switch]$ShowReport
	)

	function Set-GeneratedOnLine {
		param(
			[Parameter(Mandatory)]
			[string]$Path
		)

		$content = Get-Content -Path $Path -ErrorAction Stop
		$dateLine = Select-String -Path $Path -Pattern '^\s*# Generated on:' -ErrorAction SilentlyContinue | Select-Object -First 1

		if ($dateLine) {
			$content[$dateLine.LineNumber - 1] = "# Generated on: $(Get-Date -Format u)"
			$content | Set-Content -Path $Path -Force -ErrorAction Stop
		}
		else {
			Add-Content -Path $Path -Value "# Generated on: $(Get-Date -Format u)" -ErrorAction Stop
		}
	}

	function Add-Issue {
		param(
			[string]$Category,
			[string]$File,
			[string]$Details
		)
		[void]$Issues.Add([PSCustomObject]@{
			Catagory = $Category
			File     = $File
			details  = $Details
		})
	}

	#region Module import
	try {
		$moduleFile = $ModuleScriptFile | Get-Item -ErrorAction Stop
		$manifestPath = $moduleFile.FullName -replace '\.psm1$', '.psd1'
		if (-not (Test-Path $manifestPath -PathType Leaf)) { throw "Manifest file not found: $manifestPath" }

		Remove-Module $moduleFile.BaseName -Force -ErrorAction SilentlyContinue
		$module = Import-Module $moduleFile.FullName -Force -PassThru -ErrorAction Stop
		$OriginalModuleVer = (Import-PowerShellDataFile -Path $manifestPath).ModuleVersion

		Write-Color '[Creating]', ' PowerShell Project: ', "$($module.Name)", " [ver $($OriginalModuleVer.ToString())]" -Color Yellow, Gray, Green, Yellow -LinesBefore 2 -LinesAfter 2
		Write-Color '[Starting]', ' Module Changes' -Color Yellow, DarkCyan
	}
	catch {
		Write-Error "Error: Importing Module `nMessage:$($_.Exception.Message)"
		return
	}
	#endregion

	#region Define paths
	try {
		$ModuleBase = ((Get-Item $module.ModuleBase).Parent).FullName
		$ModulesInstuctions = Join-Path $ModuleBase 'instructions.md'
		$ModuleReadme = Join-Path $ModuleBase 'README.md'
		$ModuleIssues = Join-Path $ModuleBase 'Issues.md'
		$ModuleIssuesExcel = Join-Path $ModuleBase 'Issues.xlsx'
		$VersionFilePath = Join-Path $ModuleBase 'Version.json'
		$ModulePublicFunctions = Get-Item (Join-Path $module.ModuleBase 'Public') -ErrorAction Stop
		$ModulePrivateFunctions = Get-Item (Join-Path $module.ModuleBase 'Private') -ErrorAction Stop
		$ModuleControlScriptsPath = Join-Path $module.ModuleBase 'Control_Scripts'
		$ModuleControlScripts = if (Test-Path $ModuleControlScriptsPath) { Get-Item $ModuleControlScriptsPath } else { $null }
		$Modulemkdocs = Join-Path $ModuleBase 'docs/mkdocs.yml'
		$ModuleIndex = Join-Path $ModuleBase 'docs/docs/index.md'
		$ScriptInfoArchive = Join-Path $ModuleBase 'ScriptInfo.zip'
		[System.Collections.ArrayList]$Issues = @()
	}
	catch {
		Write-Error "Error: Defining project paths `nMessage:$($_.Exception.Message)"
		return
	}
	#endregion

	#region Remove old generated files
	Write-Color "`t[Deleting]: ", 'Output Folder' -Color Yellow, Gray
	try {
		foreach ($path in @((Join-Path $ModuleBase 'Output'), $ModuleIssues, $ModuleIssuesExcel, $VersionFilePath)) {
			if (Test-Path $path) { Remove-Item $path -Recurse -Force -ErrorAction Stop }
		}
	}
	catch {
		Write-Warning "Error: Deleting generated files `nMessage:$($_.Exception.Message)`nRetrying"
		Start-Sleep 10
		try {
			$outputPath = Join-Path $ModuleBase 'Output'
			if (Test-Path $outputPath) { Remove-Item $outputPath -Recurse -Force -ErrorAction Stop }
		}
		catch {
			Write-Error 'Error Removing Output Folder'
			return
		}
	}
	#endregion

	#region Version bump
	if ($VersionBump -ne 'CombineOnly') {
		try {
			Write-Color "`t[Processing]: ", 'Module Version Increase' -Color Yellow, Gray
			$ModuleManifestFileTMP = Get-Item $manifestPath -ErrorAction Stop
			[version]$ModuleversionTMP = (Test-ModuleManifest -Path $ModuleManifestFileTMP.FullName -ErrorAction Stop).Version

			switch ($VersionBump) {
				'Minor'    { [version]$ModuleversionTMP = '{0}.{1}.{2}' -f $ModuleversionTMP.Major, ($ModuleversionTMP.Minor + 1), 0 }
				'Build'    { [version]$ModuleversionTMP = '{0}.{1}.{2}' -f $ModuleversionTMP.Major, $ModuleversionTMP.Minor, ($ModuleversionTMP.Build + 1) }
				'Revision' { [version]$ModuleversionTMP = '{0}.{1}.{2}.{3}' -f $ModuleversionTMP.Major, $ModuleversionTMP.Minor, $ModuleversionTMP.Build, ([Math]::Max(0, $ModuleversionTMP.Revision) + 1) }
			}

			$manifestProperties = @{
				Path              = $ModuleManifestFileTMP.FullName
				ModuleVersion     = $ModuleversionTMP
				ReleaseNotes      = "Updated [$(Get-Date -Format dd/MM/yyyy_HH:mm)] $ReleaseNotes"
				FunctionsToExport = (Get-Command -Module $module.Name -CommandType Function | Select-Object -ExpandProperty Name | Sort-Object)
			}
			Update-ModuleManifest @manifestProperties -ErrorAction Stop
		}
		catch {
			Write-Error "Error: Updating Version bump `nMessage:$($_.Exception.Message)"
			return
		}
	}
	#endregion

	#region Load manifest and update generated date
	try {
		Write-Color "`t[Processing]: ", 'Adding verbose date' -Color Yellow, Gray
		$ModuleManifestFile = Get-Item $manifestPath -ErrorAction Stop
		$ModuleManifest = Test-ModuleManifest -Path $ModuleManifestFile.FullName -ErrorAction Stop | Select-Object *
		Set-GeneratedOnLine -Path $ModuleManifestFile.FullName
	}
	catch {
		Write-Error "Error: Updating Date in Module Manifest File `nMessage:$($_.Exception.Message)"
		return
	}
	#endregion

	#region Create output folder
	try {
		Write-Color "`t[Processing]: ", 'Creating Output Folder' -Color Yellow, Gray
		$ModuleOutputFolder = Join-Path (Join-Path $ModuleBase 'Output') $ModuleManifest.Version.ToString()
		$ModuleOutput = New-Item $ModuleOutputFolder -ItemType Directory -Force | Get-Item -ErrorAction Stop
	}
	catch {
		Write-Error "Error: Creating Output Folder `nMessage:$($_.Exception.Message)"
		return
	}
	#endregion

	#region Build help files
	if ($BuildHelpFiles) {
		Write-Color '[Starting]', ' Building Help Files' -Color Yellow, DarkCyan
		try {
			Write-Color "`t[Deleting]: ", 'Docs Folder' -Color Yellow, Gray
			$docsRoot = Join-Path $ModuleBase 'docs'
			if (Test-Path $docsRoot) { Remove-Item $docsRoot -Recurse -Force -ErrorAction Stop }
			if (Test-Path $ModuleReadme) { Remove-Item $ModuleReadme -Force -ErrorAction Stop }

			Write-Color "`t[Processing]: ", 'Creating Markdown Help Files' -Color Yellow, Gray
			$ModuledocsFolder = Join-Path $ModuleBase 'docs/docs'
			$Moduledocs = New-Item $ModuledocsFolder -ItemType Directory -Force | Get-Item -ErrorAction Stop
			$ModuleExternalHelpFolder = Join-Path $ModuleOutput.FullName 'en-US'
			$ModuleExternalHelp = New-Item $ModuleExternalHelpFolder -ItemType Directory -Force | Get-Item -ErrorAction Stop

			$markdownParams = @{
				Module         = $module.Name
				OutputFolder   = $Moduledocs.FullName
				WithModulePage = $false
				Locale         = 'en-US'
				HelpVersion    = $ModuleManifest.Version.ToString()
			}
			New-MarkdownHelp @markdownParams -Force

			Compare-Object -ReferenceObject (Get-ChildItem $ModulePublicFunctions -Filter '*.ps1').BaseName -DifferenceObject (Get-ChildItem $Moduledocs -Filter '*.md').BaseName |
				Where-Object { $_.SideIndicator -eq '<=' } |
				ForEach-Object { Add-Issue -Category 'External Help' -File $_.InputObject -Details 'Did not create the .md file' }

			$MissingDocumentation = Select-String -Path (Join-Path $Moduledocs.FullName '*.md') -Pattern '({{.*}})' -ErrorAction SilentlyContinue
			foreach ($item in $MissingDocumentation) {
				$object = Get-Item $item.Path
				$mod = Get-Content -Path $object.FullName
				Write-Color "`t$($object.Name):", "$($mod[$item.LineNumber - 2]) - $($mod[$item.LineNumber - 1])" -Color Yellow, Red
				Add-Issue -Category 'External Help' -File $object.Name -Details "$($object.Name) - $($mod[$item.LineNumber - 2]) - $($mod[$item.LineNumber - 1])"
			}

			Write-Color "`t[Processing]: ", 'External Help Files' -Color Yellow, Gray
			Measure-PlatyPSMarkdown -Path (Join-Path $Moduledocs.FullName '*.md') |
				Where-Object FileType -Match 'CommandHelp' |
				ForEach-Object { Import-MarkdownCommandHelp -Path $_.FilePath } |
				Export-MamlCommandHelp -OutputFolder $ModuleExternalHelp.FullName -Force

			$moduleHelpFolder = Join-Path $ModuleExternalHelp.FullName $module.Name
			if (Test-Path $moduleHelpFolder) {
				Move-Item -Path (Join-Path $moduleHelpFolder '*.xml') -Destination $ModuleExternalHelp.FullName -Force -ErrorAction SilentlyContinue
				Remove-Item $moduleHelpFolder -Recurse -Force -ErrorAction SilentlyContinue
			}

			Write-Color "`t[Processing]: ", 'About Help Files' -Color Yellow, Gray
			$aboutfile = [System.Collections.Generic.List[string]]::new()
			$aboutfile.Add('')
			$aboutfile.Add($module.Name)
			$aboutfile.Add("`t about_$($module.Name)")
			$aboutfile.Add(' ')
			$aboutfile.Add('SHORT DESCRIPTION')
			$aboutfile.Add("`t $(($ModuleManifest.Description | Out-String).Trim())")
			$aboutfile.Add(' ')
			$aboutfile.Add('NOTES')
			$aboutfile.Add('Functions in this module:')
			Get-Command -Module $module.Name -CommandType Function | Sort-Object Name | ForEach-Object { $aboutfile.Add("`t $($_.Name) -- $((Get-Help $_.Name).Synopsis)") }
			$aboutfile.Add(' ')
			$aboutfile.Add('SEE ALSO')
			if ($ModuleManifest.ProjectUri) { $aboutfile.Add("`t $($ModuleManifest.ProjectUri.AbsoluteUri)") }
			if ($ModuleManifest.HelpInfoUri) { $aboutfile.Add("`t $($ModuleManifest.HelpInfoUri)") }
			$aboutfile | Set-Content -Path (Join-Path $ModuleExternalHelp.FullName "about_$($module.Name).help.txt") -Force

			if (-not (Test-Path $ModulesInstuctions)) {
				Write-Color "`t[Processing]: ", 'Instructions Files' -Color Yellow, Gray
				$instructions = [System.Collections.Generic.List[string]]::new()
				$instructions.Add("# $($module.Name)")
				$instructions.Add(' ')
				$instructions.Add('## Description')
				$instructions.Add("$(($ModuleManifest.Description | Out-String).Trim())")
				$instructions.Add(' ')
				$instructions.Add('## Getting Started')
				$instructions.Add("- Install from PowerShell Gallery [PS Gallery](https://www.powershellgallery.com/packages/$($module.Name))")
				$instructions.Add('```powershell')
				$instructions.Add("Install-Module -Name $($module.Name) -Verbose")
				$instructions.Add('```')
				$instructions.Add("Documentation can be found at: [Github_Pages](https://smitpi.github.io/$($module.Name))")
				$instructions | Set-Content -Path $ModulesInstuctions -Force
			}

			Write-Color "`t[Processing]: ", 'Readme Files' -Color Yellow, Gray
			$readme = [System.Collections.Generic.List[string]]::new()
			Get-Content -Path $ModulesInstuctions | ForEach-Object { $readme.Add($_) }
			if ($null -ne $ModuleControlScripts) {
				$readme.Add(' ')
				$readme.Add('## PS Controller Scripts')
				Get-ChildItem $ModuleControlScripts.FullName | ForEach-Object { $readme.Add("- $($_.Name)") }
			}
			$readme.Add(' ')
			$readme.Add('## Functions')
			Get-Command -Module $module.Name -CommandType Function | Sort-Object Name | ForEach-Object { $readme.Add("- [`$($_.Name)`](https://smitpi.github.io/$($module.Name)/$($_.Name)) -- $((Get-Help $_.Name).Synopsis)") }
			$readme | Set-Content -Path $ModuleReadme -Force

			Write-Color "`t[Processing]: ", 'MKDocs Config Files' -Color Yellow, Gray
			$mkdocsFunc = [System.Collections.Generic.List[string]]::new()
			$mkdocsFunc.Add("site_name: '$($module.Name)'")
			$mkdocsFunc.Add("site_description: 'Documentation for PowerShell Module: $($module.Name)'")
			$mkdocsFunc.Add("site_author: '$(($ModuleManifest.Author | Out-String).Trim())'")
			$mkdocsFunc.Add("site_url: 'https://smitpi.github.io/$($module.Name)'")
			$mkdocsFunc.Add(' ')
			$mkdocsFunc.Add("repo_url: 'https://github.com/smitpi/$($module.Name)'")
			$mkdocsFunc.Add("repo_name: 'smitpi/$($module.Name)'")
			$mkdocsFunc.Add(' ')
			$mkdocsFunc.Add("copyright: '$(($ModuleManifest.Copyright | Out-String).Trim())'")
			$mkdocsFunc.Add(' ')
			$mkdocsFunc.Add('markdown_extensions:')
			$mkdocsFunc.Add('  - pymdownx.keys')
			$mkdocsFunc.Add('  - pymdownx.snippets')
			$mkdocsFunc.Add('  - pymdownx.superfences')
			$mkdocsFunc.Add(' ')
			$mkdocsFunc.Add('theme: material')
			$mkdocsFunc | Set-Content -Path $Modulemkdocs -Force

			Write-Color "`t[Processing]: ", 'MKDocs Index Files' -Color Yellow, Gray
			$indexFile = [System.Collections.Generic.List[string]]::new()
			Get-Content -Path $ModulesInstuctions | ForEach-Object { $indexFile.Add($_) }
			if ($null -ne $ModuleControlScripts) {
				$indexFile.Add(' ')
				$indexFile.Add('## PS Controller Scripts')
				Get-ChildItem $ModuleControlScripts.FullName | ForEach-Object { $indexFile.Add("- $($_.Name)") }
			}
			$indexFile.Add(' ')
			$indexFile.Add('## Functions')
			Get-Command -Module $module.Name -CommandType Function | Sort-Object Name | ForEach-Object { $indexFile.Add("- [`$($_.Name)`](https://smitpi.github.io/$($module.Name)/$($_.Name)) -- $((Get-Help $_.Name).Synopsis)") }
			$indexFile | Set-Content -Path $ModuleIndex -Force

			Write-Color "`t[Processing]: ", 'Versioning Files' -Color Yellow, Gray
			[PSCustomObject]@{
				version = $ModuleManifest.Version.ToString()
				Author  = $ModuleManifest.Author
				Date    = Get-Date -Format u
			} | ConvertTo-Json | Set-Content $VersionFilePath -Force
		}
		catch {
			Write-Error "Error: Building Help Files `nMessage:$($_.Exception.Message)"
			return
		}
	}
	#endregion

	#region Combine files
	try {
		Write-Color '[Starting]', ' Creating Monolithic Module Files ' -Color Yellow, DarkCyan
		$ModuleOutput = Get-Item $ModuleOutput.FullName
		$rootModule = Join-Path $ModuleOutput.FullName "$($module.Name).psm1"

		Copy-Item -Path $ModuleManifestFile.FullName -Destination $ModuleOutput.FullName -Force

		$PrivateFiles = Get-ChildItem -Path $ModulePrivateFunctions.FullName -Exclude '*.ps1' -ErrorAction SilentlyContinue
		if ($PrivateFiles) { Copy-Item -Path $PrivateFiles.FullName -Destination $ModuleOutput.FullName -Recurse -Force -ErrorAction SilentlyContinue }
		if ($null -ne $ModuleControlScripts) { Copy-Item -Path $ModuleControlScripts.FullName -Destination $ModuleOutput.FullName -Recurse -Force -ErrorAction SilentlyContinue }

		$private = @(Get-ChildItem -Path $ModulePrivateFunctions.FullName -Filter '*.ps1' -ErrorAction Stop | Sort-Object Name)
		$public = @(Get-ChildItem -Path $ModulePublicFunctions.FullName -Filter '*.ps1' -Recurse -ErrorAction Stop | Sort-Object Name)

		$file = [System.Collections.Generic.List[string]]::new()
		if ($private) {
			$file.Add('#region Private Functions')
			foreach ($PrivateItem in $private) {
				$file.Add("#region $($PrivateItem.Name)")
				$file.Add('########### Private Function ###############')
				$file.Add(('{0,-20}{1}' -f '# Source:', $PrivateItem.Name))
				$file.Add(('{0,-20}{1}' -f '# Module:', $module.Name))
				$file.Add(('{0,-20}{1}' -f '# ModuleVersion:', $ModuleManifest.Version))
				$file.Add(('{0,-20}{1}' -f '# Company:', $ModuleManifest.CompanyName))
				$file.Add(('{0,-20}{1}' -f '# CreatedOn:', $PrivateItem.CreationTime))
				$file.Add(('{0,-20}{1}' -f '# ModifiedOn:', $PrivateItem.LastWriteTime))
				$file.Add('############################################')
				Write-Color "`t[Processing]: ", $PrivateItem.Name -Color Yellow, Gray
				Get-Content $PrivateItem.FullName | ForEach-Object { $file.Add($_) }
				$file.Add('#endregion')
			}
			$file.Add('#endregion')
			$file.Add(' ')
		}

		$file.Add('#region Public Functions')
		foreach ($PublicItem in $public) {
			$author = $ModuleManifest.Author
			try {
				$ScriptInfo = Test-ScriptFileInfo -Path $PublicItem.FullName -ErrorAction Stop
				$author = $ScriptInfo.Author
			}
			catch {
				Write-Warning "`tCould not read script info [$($PublicItem.BaseName)], default values used."
				Add-Issue -Category 'ScriptFileInfo' -File $PublicItem.BaseName -Details $_.Exception.Message
			}

			$file.Add("#region $($PublicItem.Name)")
			$file.Add("######## Function $($public.IndexOf($PublicItem) + 1) of $($public.Count) ##################")
			$file.Add(('{0,-20}{1}' -f '# Function:', $PublicItem.BaseName))
			$file.Add(('{0,-20}{1}' -f '# Module:', $module.Name))
			$file.Add(('{0,-20}{1}' -f '# ModuleVersion:', $ModuleManifest.Version))
			$file.Add(('{0,-20}{1}' -f '# Author:', $author))
			$file.Add(('{0,-20}{1}' -f '# Company:', $ModuleManifest.CompanyName))
			$file.Add(('{0,-20}{1}' -f '# CreatedOn:', $PublicItem.CreationTime))
			$file.Add(('{0,-20}{1}' -f '# ModifiedOn:', $PublicItem.LastWriteTime))
			$file.Add(('{0,-20}{1}' -f '# Synopsis:', (Get-Help $PublicItem.BaseName).Synopsis))
			$file.Add('#############################################')
			$file.Add(' ')
			Write-Color "`t[Processing]: ", $PublicItem.Name -Color Yellow, Gray

			$publicContent = Get-Content -Path $PublicItem.FullName -ErrorAction Stop
			$synopsisMatch = Select-String -Path $PublicItem.FullName -Pattern '^\s*\.SYNOPSIS' | Select-Object -First 1
			if ($synopsisMatch) { [int]$StartIndex = [Math]::Max(0, $synopsisMatch.LineNumber - 2) }
			else {
				Write-Warning "`tCould not find .SYNOPSIS in [$($PublicItem.Name)]. Copying entire file."
				[int]$StartIndex = 0
				Add-Issue -Category 'ScriptFileInfo' -File $PublicItem.BaseName -Details 'Could not find .SYNOPSIS. Entire file copied into monolithic module.'
			}
			[int]$EndIndex = $publicContent.Count - 1
			$publicContent[$StartIndex..$EndIndex] | ForEach-Object { $file.Add($_) }

			$file.Add(' ')
			$file.Add("Export-ModuleMember -Function $($PublicItem.BaseName)")
			$file.Add('#endregion')
			$file.Add(' ')
		}
		$file.Add('#endregion')
		$file.Add(' ')
		$file | Set-Content -Path $rootModule -Encoding utf8 -Force
	}
	catch {
		Write-Error "Error: Creating Monolithic Module Files `nMessage:$($_.Exception.Message)"
		return
	}
	#endregion

	#region Check monolithic module
	try {
		Write-Color '[Starting]', ' Running Tests on Monolithic Module' -Color Yellow, DarkCyan
		Write-Color "`t[Confirming]: ", 'All files are created.' -Color Yellow, Gray
		$newfunction = ((Select-String -Path $rootModule -Pattern '^# Function:').Line).Replace('# Function:', '').Trim()
		$ModCommands = Get-Command -Module $module | ForEach-Object { $_.Name }
		Compare-Object -ReferenceObject $ModCommands -DifferenceObject $newfunction | ForEach-Object {
			Add-Issue -Category 'Not Copied' -File $_.InputObject -Details $_.SideIndicator
		}
	}
	catch {
		Write-Warning "Error checking monolithic module: $($_.Exception.Message)"
	}
	#endregion

	#region ScriptAnalyzer
	if ($RunScriptAnalyzer) {
		Write-Color "`t[Processing]: ", 'ScriptAnalyzer Tests.' -Color Yellow, Gray
		try {
			[System.Collections.Generic.List[pscustomobject]]$RulesObject = @()
			Invoke-ScriptAnalyzer -IncludeSuppressed -Settings CodeFormatting -Recurse -Path $ModuleOutput.FullName -Fix | ForEach-Object { $RulesObject.Add($_) }
			foreach ($setting in @('PSGallery', 'ScriptSecurity', 'ScriptFunctions', 'ScriptingStyle')) {
				Invoke-ScriptAnalyzer -IncludeSuppressed -Settings $setting -Recurse -Path $ModulePublicFunctions.PSParentPath | ForEach-Object { $RulesObject.Add($_) }
			}
			$RulesObject | ForEach-Object { Add-Issue -Category 'ScriptAnalyzer' -File $_.ScriptName -Details "[$($_.Severity)]($($_.RuleName))L $($_.Line): $($_.Message)" }
		}
		catch {
			Add-Issue -Category 'ScriptAnalyzer' -File $module.Name -Details $_.Exception.Message
		}
	}
	#endregion

	#region NestedModules
	if ($CopyNestedModules) {
		Write-Color '[Starting]', ' Copying Nested Modules' -Color Yellow, DarkCyan
		try {
			$nestedModulesPath = Join-Path -Path $ModuleOutput.FullName -ChildPath 'NestedModules'
			if (-not (Test-Path $nestedModulesPath)) { New-Item -Path $nestedModulesPath -ItemType Directory -Force | Out-Null }

			foreach ($required in $ModuleManifest.RequiredModules) {
				$latestmod = $null
				Import-Module $required -Force -ErrorAction SilentlyContinue
				$latestmod = Get-Module $required | Sort-Object Version -Descending | Select-Object -First 1
				if (-not $latestmod) { $latestmod = Get-Module $required -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1 }
				if (-not $latestmod) { Add-Issue -Category 'NestedModules' -File $required.Name -Details 'Required module not found'; continue }

				Write-Color "`t[Copying]", "$($required.Name)" -Color Yellow, DarkCyan
				Copy-Item -Path (Get-Item $latestmod.Path).Directory.FullName -Destination ([IO.Path]::Combine($ModuleOutput.FullName, 'NestedModules', $required.Name, $latestmod.Version.ToString())) -Recurse -Force
			}

			$nestedmodules = Get-ChildItem -Path $nestedModulesPath -Filter '*.psm1' -Recurse | ForEach-Object { $_.FullName.Replace("$($ModuleOutput.FullName)\", '') }
			$rootManifest = Get-Item (Join-Path $ModuleOutput.FullName "$($module.Name).psd1")
			$manifest = Import-PowerShellDataFile $ModuleManifest.Path
			foreach ($key in @('CmdletsToExport', 'AliasesToExport', 'PrivateData')) { if ($manifest.ContainsKey($key)) { $manifest.Remove($key) } }

			if (Test-Path $rootManifest.FullName) { Remove-Item $rootManifest.FullName -Force }
			New-ModuleManifest -Path $rootManifest.FullName -NestedModules $nestedmodules @manifest
			Set-GeneratedOnLine -Path $rootManifest.FullName
		}
		catch {
			Write-Warning "Error copying nested modules: $($_.Exception.Message)"
			Add-Issue -Category 'NestedModules' -File $module.Name -Details $_.Exception.Message
		}
	}
	#endregion

	#region Copy to Modules Dir
	if ($CopyToModulesFolder) {
		Write-Color '[Starting]', ' Copy to Modules Folder' -Color Yellow, DarkCyan
		$ModuleFolders = @(
			Join-Path $env:ProgramFiles 'WindowsPowerShell/Modules',
			Join-Path $env:ProgramFiles 'PowerShell/Modules',
			Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'WindowsPowerShell/Modules',
			Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PowerShell/Modules'
		)
		foreach ($moduleRoot in $ModuleFolders) {
			try {
				$destination = Join-Path $moduleRoot $moduleFile.BaseName
				if (Test-Path $destination) { Remove-Item $destination -Force -Recurse -ErrorAction Stop }
				New-Item $destination -ItemType Directory -Force | Out-Null
				Copy-Item -Path $ModuleOutput.FullName -Destination $destination -Force -Recurse -ErrorAction Stop
				Write-Color "`t[Copying]", " $destination Complete" -Color Yellow, Green
			}
			catch { Write-Warning "Error copying to [$moduleRoot]: $($_.Exception.Message)" }
		}
	}
	#endregion

	#region mkdocs
	if ($DeployMKDocs) {
		Write-Color '[Starting]', ' Creating Online Help Files ' -Color Yellow, DarkCyan
		try {
			Start-Process -FilePath pip.exe -ArgumentList 'install mkdocs-windmill' -NoNewWindow -Wait -PassThru | Out-Null
			Start-Process -FilePath mkdocs.exe -ArgumentList 'gh-deploy' -WorkingDirectory (Split-Path -Path $Modulemkdocs -Parent) -NoNewWindow -Wait | Out-Null
		}
		catch { Write-Warning "MKDocs action failed: $($_.Exception.Message)" }
	}
	#endregion

	#region Git push
	if ($GitPush) {
		try {
			if (Get-Command git.exe -ErrorAction SilentlyContinue) {
				Write-Color '[Starting]', ' Git Actions' -Color Yellow, DarkCyan
				Start-Process -FilePath git.exe -ArgumentList 'add --all' -WorkingDirectory $ModuleBase -Wait | Out-Null
				Start-Process -FilePath git.exe -ArgumentList "commit -m `"To Version: $($ModuleManifest.Version.ToString())`"" -WorkingDirectory $ModuleBase -Wait | Out-Null
				Start-Process -FilePath git.exe -ArgumentList 'push' -WorkingDirectory $ModuleBase -Wait | Out-Null
			}
			else { Write-Warning 'Git is not installed' }
		}
		catch { Write-Warning "Error: `n`tMessage:$($_.Exception.Message)" }
	}
	#endregion

	#region report issues
	if ($Issues.Count -gt 0) {
		Write-Color '[Starting]', ' Creating Issues Reports' -Color Yellow, DarkCyan
		try {
			$Issues | Export-Excel -Path $ModuleIssuesExcel -WorksheetName Other -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
		}
		catch { Write-Warning "Could not create Excel issues report: $($_.Exception.Message)" }

		$fragments = [System.Collections.Generic.List[string]]::new()
		$fragments.Add('<style>')
		$fragments.Add('table { border-collapse: collapse; }')
		$fragments.Add('table, th, td { border: 1px solid black; }')
		$fragments.Add('blockquote { border-left: solid blue; padding-left: 10px; }')
		$fragments.Add("body { color: #444; font-family: 'Open Sans', Helvetica, sans-serif; font-weight: 300; }")
		$fragments.Add('</style>')
		$fragments.Add((New-MDHeader "$($module.Name): Issues"))
		$fragments.Add("---`n")
		$fragments.Add((New-MDTable -Object $Issues))
		$fragments.Add("---`n")
		$fragments.Add("*Updated: $(Get-Date -Format U) UTC*")
		$fragments | Out-File -FilePath $ModuleIssues -Encoding utf8 -Force

		if ($ShowReport) {
			Start-Process -FilePath $ModuleIssues
			if (Test-Path $ModuleIssuesExcel) { Start-Process -FilePath $ModuleIssuesExcel }
			if ($ModuleManifest.HelpInfoUri) { Start-Process $ModuleManifest.HelpInfoUri }
		}
	}
	#endregion

	Write-Color '[Complete]', ' PowerShell Project: ', "$($module.Name)", " [ver $($ModuleManifest.Version.ToString())]" -Color Green, Gray, Green, Yellow -LinesBefore 2 -LinesAfter 2
}

$scriptblock = {
	param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
	$here = Get-Item .
	Get-ChildItem -Path . -Filter '*.psm1' -Recurse | ForEach-Object {
		$_.FullName.Replace($here.FullName, '.') | Where-Object { $_ -like "*$wordToComplete*" }
	}
}
Register-ArgumentCompleter -CommandName Set-PSProjectFile.fixed -ParameterName ModuleScriptFile -ScriptBlock $scriptBlock
