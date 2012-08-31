[CmdletBinding()]
param(
    [parameter(Mandatory=$true)]
    [ValidatePattern('^https?://')]
    [string]
    $CollectionUri,

    [string]
    $ProjectName = '*'
)

function Update-Scrum1ToScrum2 {
    param (
        $Collection,
        $Project
    )

    $NewTemplateName = 'Microsoft Visual Studio Scrum 2.0'

    Update-Tfs11CollectionWorkItemFields -Collection $Collection

    Copy-Tfs11DescriptionHtmlToDescription -Collection $Collection -Project $Project

    Import-Tfs11WorkItemTypeDefinition -Collection $Collection -Project $Project -ProcessTemplateName $NewTemplateName
    
    Import-Tfs11WorkItemTypeCategory -Collection $Collection -Project $Project -ProcessTemplateName $NewTemplateName
                
    Import-Tfs11WorkItemQuery -Collection $Collection -Project $Project -ProcessTemplateName $NewTemplateName

    Remove-Tfs11WorkItemQuery -Collection $Collection -Project $Project -Name 'All Sprints'

    Import-Tfs11CommonProcessConfiguration -Collection $Collection -Project $Project -ProcessTemplateName $NewTemplateName
                
    Import-Tfs11AgileProcessConfiguration -Collection $Collection -Project $Project -ProcessTemplateName $NewTemplateName

    Copy-Tfs11SprintToIteration -Collection $Collection -Project $Project

    Set-Tfs11DefaultTeamSettings -Collection $Collection -Project $Project

    Add-Tfs11TestVariable -Collection $Collection -Project $Project -VariableName Browser -AllowedValue 'Internet Explorer 9.0', 'Internet Explorer 10.0'
    Add-Tfs11TestVariable -Collection $Collection -Project $Project -VariableName 'Operating System' -AllowedValue 'Windows 8'

    # TODO rename 'Builders' security group to 'Build Administrators'
    
    # TODO check-in new build process template xaml files

    # TODO upload new process guidance files to Project Portal / replace Project Portal with new SP Site template

    # TODO upload new reports to SSRS

    Set-Tfs11TeamProjectProcessTemplateName -Collection $Collection -Project $Project -Name $NewTemplateName -PreviousName 'Microsoft Visual Studio Scrum 1.0'

}

function Update-Scrum2Preview3ToRTM {
    param (
        $Collection,
        $Project
    )

    $NewTemplateName = 'Microsoft Visual Studio Scrum 2.0'

    Update-Tfs11CollectionWorkItemFields -Collection $Collection
                
    Import-Tfs11WorkItemTypeDefinition -Collection $Collection -Project $Project -ProcessTemplateName $NewTemplateName

    Import-Tfs11WorkItemQuery -Collection $Collection -Project $Project -ProcessTemplateName $NewTemplateName

    Remove-Tfs11WorkItemQuery -Collection $Collection -Project $Project -Name 'Feedback Requests'

    Add-Tfs11TestVariable -Collection $Collection -Project $Project -VariableName Browser -AllowedValue 'Internet Explorer 10.0'

    Set-Tfs11TeamProjectProcessTemplateName -Collection $Collection -Project $Project -Name $NewTemplateName -PreviousName 'Microsoft Visual Studio Scrum 2.0 - Preview 3'

}

function Update-Scrum2Preview4ToRTM {
    param (
        $Collection,
        $Project
    )

    $NewTemplateName = 'Microsoft Visual Studio Scrum 2.0'

    Import-Tfs11WorkItemTypeDefinition -Collection $Collection -Project $Project -ProcessTemplateName $NewTemplateName

    Set-Tfs11TeamProjectProcessTemplateName -Collection $Collection -Project $Project -Name $NewTemplateName -PreviousName 'Microsoft Visual Studio Scrum 2.0 - Preview 4'

}

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$PSScriptRoot = $MyInvocation.MyCommand.Path | Split-Path

Import-Module -Name $PSScriptRoot\Tfs11Upgrade.psm1 -Force

$Projects = Get-Tfs11TeamProject -CollectionUri $CollectionUri -ProjectName $ProjectName |
    Sort-Object -Property ProjectName

foreach ($Project in $Projects) {
    $Project | Select-Object *Name, Is*
    switch ($Project.ProcessTemplateName) {
        'Microsoft Visual Studio Scrum 2.0 - Preview 3' {

            Update-Scrum2Preview3ToRTM -Collection $Project.CollectionUri -Project $Project.ProjectName

        }

        'Microsoft Visual Studio Scrum 2.0 - Preview 4' {

            Update-Scrum2Preview4ToRTM -Collection $Project.CollectionUri -Project $Project.ProjectName

        }

        'Microsoft Visual Studio Scrum 1.0' {
            
            Update-Scrum1ToScrum2 -Collection $Project.CollectionUri -Project $Project.ProjectName

        }
    }
}

