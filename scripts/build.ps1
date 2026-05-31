#Requires -Version 5.1
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '..\..\Contensive5\scripts\contensive-build.psm1') -Force

$projectRoot = (Resolve-Path "$PSScriptRoot\..").Path

Invoke-ContensiveBuild `
    -CollectionName    'Resource Library' `
    -CollectionPath    "$projectRoot\collection\Resource Library" `
    -SolutionPath      "$projectRoot\source\ResourceLibrary.sln" `
    -BinPath           "$projectRoot\source\ResourceLibrary\bin\Release\netstandard2.0" `
    -DeploymentRoot    'C:\Deployments\aoLibrary' `
    -CleanFolders      @(
                           "$projectRoot\source\ResourceLibrary\bin"
                           "$projectRoot\source\ResourceLibrary\obj"
                       )
