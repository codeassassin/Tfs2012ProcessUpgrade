$script:ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Get-WitAdmin {
    $RegPath = 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\VisualStudio\11.0\'
    if ([IntPtr]::Size -ne 8) { $RegPath = $RegPath -replace 'Wow6432Node', '' }
    $VSInstallDir = (Get-ItemProperty -Path $RegPath).InstallDir

    $WitadminExe = Join-Path -Path $VSInstallDir -ChildPath witadmin.exe
    if (-not (Test-Path -Path $WitadminExe -PathType Leaf)) {
        throw "Witadmin.exe not found at '$WitadminExe'"
    }

    if (([Version](Get-Item -Path $WitAdminExe).VersionInfo.ProductVersion).Major -lt 11) {
        throw "Witadmin.exe version 11 or newer required."
    }

    return $WitAdminExe
}

function Get-ProcessTemplateDefinition {
    param (
        [Parameter(Mandatory=$true)]
        $Collection,
        [Parameter(Mandatory=$true)]
        [string]
        $Name
    )

    $Collection = Get-Collection $Collection -EnsureAuthenticated

    $DefinitionPath = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ('{0}-{1}' -f $Collection.InstanceId, $Name)
    if (Test-Path -Path $DefinitionPath\ProcessTemplate.xml) {
        return $DefinitionPath
    }

    New-Item -Path $DefinitionPath -ItemType Container | Out-Null

    $TemplateService = $Collection.GetService([Microsoft.TeamFoundation.Server.IProcessTemplates])

    $Id = $TemplateService.GetTemplateIndex($Name)

    $DownloadedFile = $TemplateService.GetTemplateData($Id)

    $ZipLibPath = Join-Path -Path $PSScriptRoot -ChildPath ICSharpCode.SharpZipLib.dll
    Add-Type -Path $ZipLibPath
    $FastZip = New-Object -TypeName ICSharpCode.SharpZipLib.Zip.FastZip

    $FastZip.ExtractZip($DownloadedFile, $DefinitionPath, <# fileFilter= #> '')

    Remove-Item -Path $DownloadedFile 

    return $DefinitionPath

}


function Export-WorkItemTypeDefinition {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidatePattern('^https?://')]
        [string]
        $CollectionUri,

        [Parameter(Mandatory=$true)]
        [string]
        $ProjectName,

        [Parameter(Mandatory=$true)]
        [string]
        $WitName
    )

    Write-Verbose "Exporting work item type '$WitName' from project '$ProjectName' in collection '$CollectionUri'"
    $WitAdminExe = Get-WitAdmin
    $WorkingFile = [System.IO.Path]::GetTempFileName()
    try {
        & $WitadminExe exportwitd /collection:$CollectionUri /p:$ProjectName /n:$WitName /f:$WorkingFile | Out-Null
        if (-not $?) {
            throw "Failed to export work item type '$WitName' from project '$ProjectName' in collection '$CollectionUri'"
        }
        return [xml](Get-Content -Path $WorkingFile -Delimiter ([char]0))
    } finally {
        Remove-Item -Path $WorkingFile
    }
}

function Import-WorkItemTypeDefinition {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidatePattern('^https?://')]
        [string]
        $CollectionUri,

        [Parameter(Mandatory=$true)]
        [string]
        $ProjectName,

        [Parameter(Mandatory=$true)]
        [xml]
        $Definition
    )

    $WitName = $Definition.WITD.WORKITEMTYPE.name
    Write-Verbose "Importing work item type '$WitName' from project '$ProjectName' in collection '$CollectionUri'"
    $WitAdminExe = Get-WitAdmin
    $WorkingFile = [System.IO.Path]::GetTempFileName()
    try {
        $Definition.Save($WorkingFile)
        $Result = & $WitadminExe importwitd /collection:$CollectionUri /p:$ProjectName /f:$WorkingFile
        if (-not $?) {
            throw "Failed to import work item type '$WitName' to project '$ProjectName' in collection '$CollectionUri'`n$Result"
        }
    } finally {
        Remove-Item -Path $WorkingFile
    }
}

function Import-TeamFoundationTypes {
    [CmdletBinding()]
    param ()

    if ($script:MTF11.Count) { return }

    Add-Type -AssemblyName 'Microsoft.TeamFoundation.Client, Version=11.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a'
    Add-Type -AssemblyName 'Microsoft.TeamFoundation.WorkItemTracking.Client, Version=11.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a'
    Add-Type -AssemblyName 'Microsoft.TeamFoundation.ProjectManagement, Version=11.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a'
    Add-Type -AssemblyName 'Microsoft.TeamFoundation.TestManagement.Client, Version=11.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a'

    [AppDomain]::CurrentDomain.GetAssemblies() |
        Where-Object { $_.FullName -like 'Microsoft.TeamFoundation*, Version=11.*' } |
        ForEach-Object { 
            try {
                $_.GetTypes() |
                    Where-Object { $_.IsPublic } |
                    ForEach-Object {
                        $Key = $_.FullName -replace '^Microsoft\.TeamFoundation\.', ''
                        $script:MTF11.Add($Key, $_)
                    }
            } catch [System.Reflection.ReflectionTypeLoadException] {
                Write-Debug -Message ($_.Exception.LoaderExceptions | Format-List -Property * -Force | Out-String)
            }
        }

}

function Get-Tfs11TeamProject {
    [CmdletBinding()]
    param (
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [ValidatePattern('^https?://')]
        [string[]]
        $CollectionUri,

        [Parameter(Position=1, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [string]
        $ProjectName = '*'
    )

    begin {
        Import-TeamFoundationTypes
    }

    process {
        foreach ($Uri in $CollectionUri) {
            $Collection = Get-Collection $Uri -EnsureAuthenticated
            $Structure = $Collection.GetService($MTF11['Server.ICommonStructureService4'])
            $Projects = $Structure.ListProjects() |
                Where-Object { $_.Name -like $ProjectName }
            foreach ($Project in $Projects) {

                Get-Tfs11TeamProjectProcessTemplate -Collection $Collection -Project $Project

            }
        }
    }
    
}

function Get-Collection {
    param (
        $Collection,
        [switch]$EnsureAuthenticated
    )

    Import-TeamFoundationTypes

    if (-not $Collection) { throw "collection is null" }

    if ($Collection -is [string]) {
        $Collection = [uri]$Collection
    }

    if ($Collection -is [uri]) {
        $Collection = $MTF11['Client.TfsTeamProjectCollectionFactory']::GetTeamProjectCollection($Collection)
    }

    if ($Collection -isnot $MTF11['Client.TfsTeamProjectCollection']) {
        throw "Invalid Collection value"
    }

    if ($EnsureAuthenticated) {
        $Collection.EnsureAuthenticated()
    }

    return $Collection
}

function Get-ProjectInfo {
    param (
        [Microsoft.TeamFoundation.Client.TfsTeamProjectCollection, Microsoft.TeamFoundation.Client, Version=11.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a]
        $Collection,
        $Project
    )

    if (-not $Project) { return }

    $Structure = $Collection.GetService($MTF11['Server.ICommonStructureService4'])
    
    if ($Project -is [string]) {
        $Project = $Structure.GetProjectFromName($Project)
    }

    if ($Project -is [uri]) {
        $Project = $Structure.GetProject($Project)
    }

    if ($Project -isnot $MTF11['Server.ProjectInfo']) {
        throw "Invalid Project value"
    }

    return $Project

}

function Get-Tfs11TeamProjectProcessTemplate {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        $Collection,
        [Parameter(Position = 1, Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        $Project
    )

    begin {
        Import-TeamFoundationTypes
    }

    process {

        $Collection = Get-Collection $Collection -EnsureAuthenticated
        $Project = Get-ProjectInfo $Collection $Project

        $Structure = $Collection.GetService($MTF11['Server.ICommonStructureService4'])
        $Store = $Collection.GetService($MTF11['WorkItemTracking.Client.WorkItemStore'])
        
        $IsTemplateGuess = $false
        $TemplateName = $Structure.GetProjectProperty($Project.Uri, 'Process Template').Value
        if (-not $TemplateName) {
            $IsTemplateGuess = $true

            $WITs = $Store.Projects[$Project.Name].WorkItemTypes

            if ($WITs.Contains('Sprint Backlog Task')) {
                $TemplateName = 'Scrum for Team System v3.0.4190.00'

            } elseif ($WITs.Contains('Product Backlog Item')) {
                $TemplateName = 'Microsoft Visual Studio Scrum 1.0'

            } elseif ($WITs.Contains('User Story')) {
                $TemplateName = 'MSF for Agile Software Development v5.0'

            } elseif ($WITs.Contains('Change Request')) {
                $TemplateName = 'MSF for CMMI Process Improvement v5.0'
            }
        }

        New-Object -TypeName PSObject -Property @{
            CollectionUri = $Collection.Uri
            ProjectUri = $Project.Uri
            ProjectName = $Project.Name
            ProcessTemplateName = $TemplateName
            IsTemplateGuess = $IsTemplateGuess
        }

    }
}

function Set-Tfs11TeamProjectProcessTemplateName {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory=$true)]
        $Collection,

        [Parameter(Position = 1, Mandatory=$true)]
        $Project,

        [Parameter(Position = 2, Mandatory=$true)]
        [string]
        $Name,

        [string]
        $PreviousName
    )

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Status 'Setting project process template property'

    $Collection = Get-Collection $Collection -EnsureAuthenticated
    $Project = Get-ProjectInfo $Collection $Project

    $Structure = $Collection.GetService($MTF11['Server.ICommonStructureService4'])

    if ($PreviousName) {
        $Structure.SetProjectProperty($Project.Uri, 'Previous Process Template', $PreviousName)
    }
    $Structure.SetProjectProperty($Project.Uri, 'Process Template', $Name)
        
    Write-Progress -Activity $MyInvocation.MyCommand.Name -Completed

}

function Update-Tfs11CollectionWorkItemFields {
    [CmdletBinding()]
    param (
        [Parameter(Position=0, Mandatory=$true)]
        $Collection
    )

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Status 'Checking System.Description is Html'

    $Collection = Get-Collection $Collection -EnsureAuthenticated

    $Store = $Collection.GetService($MTF11['WorkItemTracking.Client.WorkItemStore'])

    $WitAdminPath = Get-WitAdmin

    if ($Store.FieldDefinitions['System.Description'].FieldType -ne 'Html') {
        throw "Collection's [System.Description] field hasn't been upgraded to Html yet."
        # TODO witadmin changefield 
    }

    $Fields = @(
        @{ RefName = 'System.IterationId'; NewName = 'Iteration ID' }
        @{ RefName = 'System.AreaId'; NewName = 'Area ID' }
        @{ RefName = 'Microsoft.VSTS.TCM.AutomatedTestId'; NewName = 'Automated Test Id' }
        @{ RefName = 'Microsoft.VSTS.TCM.AutomatedTestName'; NewName = 'Automated Test Name' }
        @{ RefName = 'Microsoft.VSTS.TCM.AutomatedTestStorage'; NewName = 'Automated Test Storage' }
        @{ RefName = 'Microsoft.VSTS.TCM.AutomatedTestType'; NewName = 'Automated Test Type' }
        @{ RefName = 'Microsoft.VSTS.TCM.LocalDataSource'; NewName = 'Local Data Source' }
        @{ RefName = 'Microsoft.VSTS.TCM.ReproSteps'; NewName = 'Repro Steps' }
    )

    $Count = 0
    foreach ($Field in $Fields) {
        Write-Progress -Activity $MyInvocation.MyCommand.Name -Status $Field.RefName -PercentComplete (100 * $Count / $Fields.Count)

        if ($Store.FieldDefinitions[$Field.RefName].Name -cne $Field.NewName) {
            Write-Verbose "Renaming '$($Field.RefName)' to '$($Field.NewName)'."
            & $WitAdminPath changefield /collection:$($Collection.Uri) /n:$($Field.RefName) /name:"$($Field.NewName)" /noprompt
            if (-not $?) {
                Write-Error "witadmin changefield failed"
            }
        }

        $Count++
    }
    
    Write-Progress -Activity $MyInvocation.MyCommand.Name -Completed

}

function Copy-Tfs11DescriptionHtmlToDescription {
    [CmdletBinding()]
    param (
        [Parameter(Position=0, Mandatory=$true)]
        $Collection,
        [Parameter(Position=1, Mandatory=$true)]
        $Project
    )

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Status 'Querying work items with the DescriptionHtml field'

    $Collection = Get-Collection $Collection -EnsureAuthenticated
    $Project = Get-ProjectInfo $Collection $Project 

    $Store = $Collection.GetService($MTF11['WorkItemTracking.Client.WorkItemStore'])
    $Store.RefreshCache()

    if ($Store.FieldDefinitions['System.Description'].FieldType -ne 'Html') {
        throw "Collection's [System.Description] field hasn't been upgraded to Html yet."
    }

    $CandidateWITs = @()
    foreach ($WIT in ($Store.Projects[$Project.Name].WorkItemTypes)) {
        if ($WIT.FieldDefinitions.Contains('Microsoft.VSTS.Common.DescriptionHtml')) {
            $CandidateWITs += $WIT.Name
        }
    }
    if (-not $CandidateWITs) { return }
    
    $WiqlWITs = "'" + ($CandidateWITs -join "','") + "'"

    $WorkItems = $Store.Query(@"
    SELECT [System.Id], [System.WorkItemType], [System.Description], [Microsoft.VSTS.Common.DescriptionHtml] 
    FROM WorkItems 
    WHERE [System.TeamProject] = @project  
    AND  [System.WorkItemType] IN ($WiqlWITs) 
    ORDER BY [System.Id]
"@, @{ project = $Project.Name })

    $Count = 0
    foreach ($WorkItem in $WorkItems) {
        Write-Progress -Activity $MyInvocation.MyCommand.Name -Status "Work item $($WorkItem.Id)" -PercentComplete (100 * $Count / $WorkItems.Count)

        $SourceFieldValue = $WorkItem['Microsoft.VSTS.Common.DescriptionHtml']
        if ($WorkItem.Description -cne $SourceFieldValue) {
            # element case and html entities can differ but TFS will ignore these superficial changes
            $WorkItem.Open()
            $WorkItem.Description = $SourceFieldValue 
            $WorkItem.Save()
        }
        $Count++
    }
    Write-Progress -Activity $MyInvocation.MyCommand.Name -Completed

}

function Copy-Tfs11SprintToIteration {
    [CmdletBinding()]
    param (
        [Parameter(Position=0, Mandatory=$true)]
        $Collection,
        [Parameter(Position=1, Mandatory=$true)]
        $Project
    )

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Status 'Querying Sprint work items'

    $Collection = Get-Collection $Collection -EnsureAuthenticated
    $Project = Get-ProjectInfo $Collection $Project 

    $Store = $Collection.GetService($MTF11['WorkItemTracking.Client.WorkItemStore'])
    $Structure = $Collection.GetService($MTF11['Server.ICommonStructureService4'])

    $StoreProject = $Store.Projects[$Project.Name]

    $WorkItems = $Store.Query(@"
    SELECT [System.Id], [System.IterationId], [System.IterationPath], [Microsoft.VSTS.Scheduling.StartDate], [Microsoft.VSTS.Scheduling.FinishDate]
    FROM WorkItems 
    WHERE [System.TeamProject] = @project  
    AND  [System.WorkItemType] = 'Sprint'
    ORDER BY [System.Id]
"@, @{ project = $Project.Name })


    $Count = 0
    foreach ($WorkItem in $WorkItems) {
        Write-Progress -Activity $MyInvocation.MyCommand.Name -Status $WorkItem.IterationPath -PercentComplete (100 * $Count / $WorkItems.Count)
        $Node = $null
        try {
            $Node = $StoreProject.FindNodeInSubTree($WorkItem.IterationId)
        } catch { <# swallow #> }
        if ($Node) {
            $NodeInfo = $Structure.GetNode($Node.Uri)
            $StartDate = $WorkItem['Microsoft.VSTS.Scheduling.StartDate']
            $FinishDate = $WorkItem['Microsoft.VSTS.Scheduling.FinishDate']
            if (($StartDate -and $StartDate.Date -ne $NodeInfo.StartDate) -or
                    ($FinishDate -and $FinishDate.Date -ne $NodeInfo.FinishDate)) {
                $Structure.SetIterationDates($NodeInfo.Uri, $StartDate, $FinishDate)
            }
        }
        $Count++
    }
    Write-Progress -Activity $MyInvocation.MyCommand.Name -Completed

}

function Set-Tfs11DefaultTeamSettings {
    [CmdletBinding()]
    param (
        [Parameter(Position=0, Mandatory=$true)]
        $Collection,
        [Parameter(Position=1, Mandatory=$true)]
        $Project
    )

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Status 'Setting default team area and iterations'

    $Collection = Get-Collection $Collection -EnsureAuthenticated
    $Project = Get-ProjectInfo $Collection $Project 

    $Structure = $Collection.GetService($MTF11['Server.ICommonStructureService4'])
    $TeamService = $Collection.GetService($MTF11['Client.TfsTeamService'])
    $TeamConfigService = $Collection.GetService($MTF11['ProcessConfiguration.Client.TeamSettingsConfigurationService'])

    $TeamId = $TeamService.GetDefaultTeamId($Project.Uri)
    $TeamConfig = $TeamConfigService.GetTeamConfigurations([guid[]]@($TeamId)) |
        Select-Object -First 1

    $TeamSettings = $TeamConfig.TeamSettings
    $Dirty = $false

    if (-not $TeamSettings.BacklogIterationPath) {
        $TeamSettings.BacklogIterationPath = $Project.Name
        $Dirty = $true
    }

    if ($TeamSettings.IterationPaths.Length -eq 0) {
                    
        $IterationRoot = $Structure.ListStructures($Project.Uri) |
            Where-Object { $_.StructureType -eq 'ProjectLifecycle' } |
            Select-Object -First 1
        $IterationsXml = $Structure.GetNodesXml(@($IterationRoot.Uri),  $true)
        $IterationNodes = Select-Xml -Xml $IterationsXml -XPath //Node |
            Select-Object -ExpandProperty Node
        $FullIterationPaths = $IterationNodes |
            Where-Object { -not $_.HasChildNodes } |
            Select-Object -ExpandProperty Path
                    
        $TeamSettings.IterationPaths = $FullIterationPaths -replace '^\\[^\\]+\\Iteration\\', ($Project.Name + '\')
        #$TeamSettings.CurrentIteration = $TeamSettings.IterationPaths[0]
        $Dirty = $true
    }

    if ($TeamSettings.TeamFieldValues.Length -eq 0) {
        $TFV = New-Object -TypeName $MTF11['ProcessConfiguration.Client.TeamFieldValue'].AssemblyQualifiedName
        $TFV.Value = $Project.Name
        $TFV.IncludeChildren = $true
        $TeamSettings.TeamFieldValues = @($TFV)
        $Dirty = $true
    }

    if ($Dirty) {
        $TeamConfigService.SetTeamSettings($TeamId, $TeamSettings)
    }

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Completed

}


function Import-Tfs11WorkItemTypeCategory {
    [CmdletBinding()]
    param (
        [Parameter(Position=0, Mandatory=$true)]
        $Collection,
        [Parameter(Position=1, Mandatory=$true)]
        $Project,
        [Parameter(Position=2, Mandatory=$true)]
        $ProcessTemplateName
    )

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Status 'Importing categories'

    $WitAdminPath = Get-WitAdmin

    $Collection = Get-Collection $Collection -EnsureAuthenticated
    $Project = Get-ProjectInfo $Collection $Project 

    $ProcessTemplatePath = Get-ProcessTemplateDefinition -Collection $Collection -Name $ProcessTemplateName
    $ProcessTemplateXmlPath = Join-Path -Path $ProcessTemplatePath -ChildPath ProcessTemplate.xml
    
    $WorkItemsXmlPath = Join-Path -Path $ProcessTemplatePath -ChildPath (
        (Select-Xml -Path $ProcessTemplateXmlPath -XPath '//group[@id="WorkItemTracking"]/taskList').Node.filename
    )
    
    $CategoriesXmlPath = Join-Path -Path $ProcessTemplatePath -ChildPath (
        (Select-Xml -Path $WorkItemsXmlPath -XPath '//task[@id="Categories"]/taskXml/CATEGORIES').Node.fileName
    )
    
    $Result = & $WitAdminPath importcategories /collection:$($Collection.Uri) /p:"$($Project.Name)" /f:"$CategoriesXmlPath"
    if (-not $?) {
        Write-Error "Category import failed"
    }

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Completed

}

function Import-Tfs11CommonProcessConfiguration {
    [CmdletBinding()]
    param (
        [Parameter(Position=0, Mandatory=$true)]
        $Collection,
        [Parameter(Position=1, Mandatory=$true)]
        $Project,
        [Parameter(Position=2, Mandatory=$true)]
        $ProcessTemplateName
    )

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Status 'Importing common process configuration'

    $WitAdminPath = Get-WitAdmin

    $Collection = Get-Collection $Collection -EnsureAuthenticated
    $Project = Get-ProjectInfo $Collection $Project 

    $ProcessTemplatePath = Get-ProcessTemplateDefinition -Collection $Collection -Name $ProcessTemplateName
    $ProcessTemplateXmlPath = Join-Path -Path $ProcessTemplatePath -ChildPath ProcessTemplate.xml
    
    $WorkItemsXmlPath = Join-Path -Path $ProcessTemplatePath -ChildPath (
        (Select-Xml -Path $ProcessTemplateXmlPath -XPath '//group[@id="WorkItemTracking"]/taskList').Node.filename
    )
    
    $CommonConfigXmlPath = Join-Path -Path $ProcessTemplatePath -ChildPath (
        (Select-Xml -Path $WorkItemsXmlPath -XPath '//task[@id="ProcessConfiguration"]/taskXml/PROCESSCONFIGURATION/CommonConfiguration').Node.fileName
    )
    
    $Result = & $WitAdminPath importcommonprocessconfig /collection:$($Collection.Uri) /p:"$($Project.Name)" /f:"$CommonConfigXmlPath"
    if (-not $?) {
        Write-Error "Common process config import failed"
    }

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Completed

}

function Import-Tfs11AgileProcessConfiguration {
    [CmdletBinding()]
    param (
        [Parameter(Position=0, Mandatory=$true)]
        $Collection,
        [Parameter(Position=1, Mandatory=$true)]
        $Project,
        [Parameter(Position=2, Mandatory=$true)]
        $ProcessTemplateName
    )

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Status 'Importing agile process configuration'

    $WitAdminPath = Get-WitAdmin

    $Collection = Get-Collection $Collection -EnsureAuthenticated
    $Project = Get-ProjectInfo $Collection $Project 

    $ProcessTemplatePath = Get-ProcessTemplateDefinition -Collection $Collection -Name $ProcessTemplateName
    $ProcessTemplateXmlPath = Join-Path -Path $ProcessTemplatePath -ChildPath ProcessTemplate.xml
    
    $WorkItemsXmlPath = Join-Path -Path $ProcessTemplatePath -ChildPath (
        (Select-Xml -Path $ProcessTemplateXmlPath -XPath '//group[@id="WorkItemTracking"]/taskList').Node.filename
    )
    
    $ConfigXmlPath = Join-Path -Path $ProcessTemplatePath -ChildPath (
        (Select-Xml -Path $WorkItemsXmlPath -XPath '//task[@id="ProcessConfiguration"]/taskXml/PROCESSCONFIGURATION/AgileConfiguration').Node.fileName
    )
    
    $Result = & $WitAdminPath importagileprocessconfig /collection:$($Collection.Uri) /p:"$($Project.Name)" /f:"$ConfigXmlPath"
    if (-not $?) {
        Write-Error "Common process config import failed"
    }

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Completed

}

function Import-Tfs11WorkItemTypeDefinition {
    [CmdletBinding()]
    param (
        [Parameter(Position=0, Mandatory=$true)]
        $Collection,

        [Parameter(Position=1, Mandatory=$true)]
        $Project,

        [Parameter(Position=2, Mandatory=$true)]
        $ProcessTemplateName
    )

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Status 'Retrieving work item type definitions'

    $WitAdminPath = Get-WitAdmin

    $Collection = Get-Collection $Collection -EnsureAuthenticated
    $Project = Get-ProjectInfo $Collection $Project 

    $ProcessTemplatePath = Get-ProcessTemplateDefinition -Collection $Collection -Name $ProcessTemplateName
    $ProcessTemplateXmlPath = Join-Path -Path $ProcessTemplatePath -ChildPath ProcessTemplate.xml
    
    $WorkItemsXmlPath = Join-Path -Path $ProcessTemplatePath -ChildPath (
        (Select-Xml -Path $ProcessTemplateXmlPath -XPath '//group[@id="WorkItemTracking"]/taskList').Node.filename
    )

    $WitdFiles = Select-Xml -Path $WorkItemsXmlPath -XPath '//task[@id="WITs"]/taskXml/WORKITEMTYPES/WORKITEMTYPE' |
        Select-Object -Property @{N='file'; E={ $_.Node.fileName } } |
        Select-Object -ExpandProperty file

    $Count = 0
    foreach ($WitdFile in $WitdFiles) {
        Write-Progress -Activity $MyInvocation.MyCommand.Name -Status $WitdFile -PercentComplete (100 * $Count / $WitdFiles.Count)

        $WitdPath = Join-Path -Path $ProcessTemplatePath -ChildPath $WitdFile
        $Result = & $WitAdminPath importwitd /collection:$($Collection.Uri) /p:"$($Project.Name)" /f:"$WitdPath"
        if (-not $?) {
            Write-Error "FAIL: $WitdPath"
        }

        $Count++
    }

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Completed

}


function Add-Tfs11TestVariable {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $Collection,
        [Parameter(Mandatory=$true)]
        $Project,
        [Parameter(Mandatory=$true)]
        [string]
        $VariableName,
        [Parameter(Mandatory=$true)]
        [string[]]
        $AllowedValue
    )

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Status 'Adding test variable values'

    $Collection = Get-Collection $Collection -EnsureAuthenticated
    $Project = Get-ProjectInfo $Collection $Project 

    $TestManagementService = $Collection.GetService($MTF11['TestManagement.Client.ITestManagementService'])

    $TestProject = $TestManagementService.GetTeamProject($Project.Name)

    $Variable = $TestProject.TestVariables.Query() |
        Where-Object { $_.Name -eq $VariableName }

    $Dirty = $false
    foreach ($SingleValue in $AllowedValue) {
        $ExistingValue = $Variable.AllowedValues |
            Where-Object { $_.Value -ceq $SingleValue }
        if (-not $ExistingValue) {
            Write-Verbose "Adding allowed value '$SingleValue' to test variable '$VariableName'."
            $NewValue = $TestProject.TestVariables.CreateVariableValue($SingleValue)
            $Variable.AllowedValues.Add($NewValue)
            $Dirty = $true

        }
    }

    if ($Dirty) {
        $Variable.Save()
    }

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Completed

}

function Remove-Tfs11WorkItemQuery {
    [CmdletBinding()]
    param (
        [Parameter(Position=0, Mandatory=$true)]
        $Collection,
        [Parameter(Position=1, Mandatory=$true)]
        $Project,
        [Parameter(Position=2, Mandatory=$true)]
        $Name
    )

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Status 'Removing work item query'

    $Collection = Get-Collection $Collection -EnsureAuthenticated
    $Project = Get-ProjectInfo $Collection $Project 

    $Store = $Collection.GetService($MTF11['WorkItemTracking.Client.WorkItemStore'])

    $Hierarchy = $Store.Projects[$Project.Name].QueryHierarchy
    $SharedQueriesFolder = $Hierarchy | Where-Object { -not $_.IsPersonal }

    if ($SharedQueriesFolder.Contains($Name)) {
        $SharedQueriesFolder[$Name].Delete()
        $Hierarchy.Save()
    }

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Completed

}

function Import-ProcessTemplateQuery {
    [CmdletBinding()]
    param (
        $Hierarchy,
        $ProcessTemplatePath,
        $Folder,
        $QueriesXml
    )
    
    $SubfoldersXml = Select-Xml -Xml $QueriesXml -XPath QueryFolder | Select-Object -ExpandProperty Node
    foreach ($SubfolderXml in $SubfoldersXml) {
        if ($Folder.Contains($SubfolderXml.name)) {
            $Subfolder = $Folder[$SubfolderXml.name]
            if ($Subfolder -isnot [Microsoft.TeamFoundation.WorkItemTracking.Client.QueryFolder]) {
                throw ("folder name used by query: " + $SubfolderXml.name)
            }
        } else {
            $Subfolder = New-Object -TypeName Microsoft.TeamFoundation.WorkItemTracking.Client.QueryFolder -ArgumentList $SubfolderXml.name, $Folder
            $Hierarchy.Save()
        }
        Import-ProcessTemplateQuery $Hierarchy $ProcessTemplatePath $Subfolder $SubfolderXml
    }

    $DefinitionsXml = Select-Xml -Xml $QueriesXml -XPath Query | Select-Object -ExpandProperty Node
    foreach ($DefinitionXml in $DefinitionsXml) {

        $WiqPath = Join-Path -Path $ProcessTemplatePath -ChildPath $DefinitionXml.fileName
        $QueryText = (Select-Xml -Path $WiqPath -XPath WorkItemQuery/Wiql).Node.InnerText
        $QueryText = $QueryText -replace '\$\$PROJECTNAME\$\$', $Hierarchy.Project.Name

        if ($Folder.Contains($DefinitionXml.name)) {
            $QDef = $Folder[$DefinitionXml.name]
            if ($QDef.QueryText -cne $QueryText) {
                Write-Verbose ('Updating query ' + $DefinitionXml.name)
                $QDef.QueryText = $QueryText
                $Hierarchy.Save()
            }
        } else {
            Write-Verbose ('Adding query ' + $DefinitionXml.name)
            $QDef = New-Object -TypeName Microsoft.TeamFoundation.WorkItemTracking.Client.QueryDefinition -ArgumentList $DefinitionXml.name, $QueryText, $Folder
            $Hierarchy.Save()
        }

    }

}

function Import-Tfs11WorkItemQuery {
    [CmdletBinding()]
    param (
        [Parameter(Position=0, Mandatory=$true)]
        $Collection,

        [Parameter(Position=1, Mandatory=$true)]
        $Project,

        [Parameter(Position=2, Mandatory=$true)]
        $ProcessTemplateName
    )

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Status 'Importing work item queries'

    $Collection = Get-Collection $Collection -EnsureAuthenticated
    $Project = Get-ProjectInfo $Collection $Project 

    $Store = $Collection.GetService($MTF11['WorkItemTracking.Client.WorkItemStore'])

    $ProcessTemplatePath = Get-ProcessTemplateDefinition -Collection $Collection -Name $ProcessTemplateName
    $ProcessTemplateXmlPath = Join-Path -Path $ProcessTemplatePath -ChildPath ProcessTemplate.xml
    
    $WorkItemsXmlPath = Join-Path -Path $ProcessTemplatePath -ChildPath (
        (Select-Xml -Path $ProcessTemplateXmlPath -XPath '//group[@id="WorkItemTracking"]/taskList').Node.filename
    )

    $QueriesXml = (Select-Xml -Path $WorkItemsXmlPath -XPath 'tasks/task[@id="Queries"]/taskXml/QUERIES').Node

    $Hierarchy = $Store.Projects[$Project.Name].QueryHierarchy
    $SharedQueriesFolder = $Hierarchy | Where-Object { -not $_.IsPersonal }
    Import-ProcessTemplateQuery -Hierarchy $Hierarchy -ProcessTemplatePath $ProcessTemplatePath -Folder $SharedQueriesFolder -QueriesXml $QueriesXml

    Write-Progress -Activity $MyInvocation.MyCommand.Name -Completed

}

$script:MTF11 = @{}

Export-ModuleMember -Function *-Tfs11*