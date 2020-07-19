+++
author = "Chendrayan Venkatesan"
categories = ["AzDevOps"]
tags = ["KQL" , "PowerShell" , "Azure DevOps", "Reporting"]
date = "2020-07-19"
description = ""
featured = "STG01.JPG"
featuredalt = "STG01"
featuredpath = "date"
linktitle = ""
title = "Azure DevOps Pipeline To Send Files Through Email With No MarketPlace Extension"
type = "post"
+++

# Introduction

I was developing a PowerShell script to send Azure inventory through email, which generates a Word output with charts and tables. So, I thought of using Open XML, I used for SharePoint document library reporting (A few years ago), I searched GitHub and found a great PowerShell module [PSWriteWord](https://github.com/EvotecIT/PSWriteWord), and docs are super cool. Yes, this blog post is to show you the simple steps to generate a word document through Azure DevOps pipeline. 

> What do we expect? 

A word document with a TOC, charts, tables to illustrate the Azure infrastructure information. For example

![Storage Accounts](img\2020\07\STG01.JPG)

![Resources](img\2020\07\Resources01.JPG)

> What do we need?

1. Azure Account
2. Azure DevOps Account 
3. VS Code (or any of your favorite IDE)
4. PowerShell
5. SMTP information (Yes, send DOCX through email)

To make this blog post short, let me show the piece of code I used in my assignment at work. 

> Azure Resources By Location

```PowerShell
#region
$resources = Search-AzGraph -Query "Resources | summarize count() by type | top 5 by type | project type, count_"
$WordDocument = New-WordDocument -FilePath "C:\Temp\Azure-Infrastructure-Report.docx"
Add-WordText -WordDocument $WordDocument -Text "Azure Inventory" -FontSize 72 -Alignment center -Color Black
Add-WordPageBreak -WordDocument $WordDocument
Add-WordTOC -WordDocument $WordDocument
Add-WordPageBreak -WordDocument $WordDocument
#Add-WordText -WordDocument $WordDocument -Text "Top 5 Azure Resource type by Count" -HeadingType Heading3 -Color Black -Alignment center
Add-WordBarChart -WordDocument $WordDocument -ChartName "Azure Resource by Type" -Names $($resources).type -Values $($resources).count_ -ChartLegendPosition Left -ChartLegendOverlay $false -BarDirection Column
#endregion
```

> Virtual Machines (Count by Location)

```PowerShell
#region
$virtualMachines = Search-AzGraph -Query "Resources | where type =~ 'microsoft.compute/virtualMachines' | summarize count() by location"
Add-WordText -WordDocument $WordDocument -Text "Virtual Machines" -HeadingType Heading3 -Color Black -Alignment center
Add-WordBarChart -WordDocument $WordDocument -ChartName "Virtual Machines by Location" -Names $($virtualMachines.location) -Values $($virtualMachines.count_) -NoLegend
#endregion
```

> Storage Accounts (Count by Location)

```PowerShell
#region
$storageAccounts = Search-AzGraph -Query "Resources | where type =~ 'microsoft.storage/storageAccounts' | summarize count() by location"
Add-WordText -WordDocument $WordDocument -Text "Storage Accounts" -HeadingType Heading3 -Color Black -Alignment center
Add-WordBarChart -WordDocument $WordDocument -ChartName "Storage Accounts by Location" -Names $($storageAccounts.location) -Values $($storageAccounts.count_) -ChartLegendPosition Left -ChartLegendOverlay $false -BarDirection Column
#endregion
```

> Save the document 

```PowerShell
#region
Save-WordDocument -WordDocument $WordDocument -FilePath ".\Azure-Infrastructure-Report.docx"
#endregion 
```

> Az Pipeline

```YAML
trigger:
  branches:
    include:
      - 'dev'

# Windows Image
pool:
  vmImage: 'windows-latest'

variables:
  - group: 'Reporting-Credentials'
  - name: 'Reporting-Credentials'
  
stages:
  - stage: Build
    jobs:
      - job: 
        steps:
          - task: AzurePowerShell@5
            inputs:
              azurePowerShellVersion: LatestVersion
              azureSubscription: 'Reporting'
              ScriptType: FilePath
              ScriptPath: 'TruGreen\azureInventory.ps1'
          - task: CopyFiles@2
            inputs:
              Contents: 'Azure Infrastructure Report.docx'
              TargetFolder: $(Build.ArtifactStagingDirectory)
          - task: PublishBuildArtifacts@1
            inputs:
              ArtifactName: 'azureInventory'
              PathtoPublish: $(Build.ArtifactStagingDirectory)
              publishLocation: Container

  - stage: Release
    jobs:
      - job: 
        steps:
          - task: DownloadBuildArtifacts@0
            inputs:
              buildType: current
              downloadType: single
              artifactName: 'azureInventory'
              downloadPath: $(System.ArtifactsDirectory)
          - task: PowerShell@2
            inputs:
              targetType: filePath
              filePath: 'TruGreen\sendEmail.ps1'
              arguments: '-Attachments "$(System.ArtifactsDirectory)\azureInventory\Azure Infrastructure Report.docx" -MailID "$(MailID)" -MailPassword "$(MailPassword)"'
```