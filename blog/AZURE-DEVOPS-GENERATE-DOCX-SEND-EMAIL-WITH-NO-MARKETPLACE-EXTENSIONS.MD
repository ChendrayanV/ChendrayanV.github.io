+++
author = "Chendrayan Venkatesan"
categories = ["Azure"]
tags = ["Send DOCX" , "Azure DevOps", "PowerShell"]
date = "2020-07-18"
description = ""
featured = "MSWord.jpg"
featuredalt = "Sand-Clock"
featuredpath = "date"
linktitle = ""
title = "AZURE DEVOPS GENERATE DOCX SEND EMAIL WITH NO MARKETPLACE EXTENSIONS"
type = "post"
+++

## Introduction

I was developing an as-built script for Azure infrastructure, which should produce a docx file output with charts and tables (especially for inventory). So, I thought of using Open XML, which I used, a few years ago for SharePoint document library reporting. I searched GitHub and found a great PowerShell module [PSWriteWord](https://github.com/EvotecIT/PSWriteWord), and docs are super cool. Yes, this blog post is to show you the simple steps to generate a word document through the Azure DevOps pipeline.

**Credits To: Author [@PrzemyslawKlys](https://twitter.com/PrzemyslawKlys) of the PowerShell module PSWriteWord.**

> What do we expect? 

A word document with a TOC, charts, tables to illustrate the Azure infrastructure information. For example, the chart in the docx is below 

![image](img/2020/07/Resource-Top-05.jpg)

> What do we need?

1. Azure Account
2. Azure DevOps Account 
3. VS Code (or any of your favorite IDE)
4. PowerShell
5. SMTP information (Yes, send DOCX through email)

To make this blog post short, let me show the piece of code I used in my assignment at work. 

### Resources Summary By Count

```PowerShell
Search-AzGraph -Query "Resources | summarize count() by type | top 5 by type | project type, count_"
```

### Unhealthy Resource Summary 

```PowerShell
Search-AzGraph -Query "securityResources | where type == 'microsoft.security/assessments' and properties.status.code =~ 'Unhealthy' | summarize count() by resourceGroup | top 5 by resourceGroup | project resourceGroup , count_"
```

Keep adding the queries for your need and make an excellent infrastructure document. Here is the snippet to generate a document and save it

```PowerShell
#region - Resources By Count
$resources = Search-AzGraph -Query "Resources | summarize count() by type | top 5 by type"
$WordDocument = New-WordDocument -FilePath ".\Azure-Infrastructure-Report.docx"
Add-WordTOC -WordDocument $WordDocument
Add-WordPageBreak -WordDocument $WordDocument
Add-WordText -WordDocument $WordDocument -Text "Top 5 Azure Resource type by Count" -HeadingType Heading3 -Color Black -Alignment center
Add-WordText -WordDocument $WordDocument -NewLine $true
DocText -LineBreak 
Add-WordBarChart -WordDocument $WordDocument -ChartName "Azure Resources Count by Type" -Names $($resources.type) -Values $($resources.count_) -ChartLegendPosition Left -ChartLegendOverlay $false -BarDirection Column
DocText -LineBreak 
#endregion

#region - Resources By Count
$unHealthyResources = Search-AzGraph -Query "securityResources | where type == 'microsoft.security/assessments' and properties.status.code =~ 'Unhealthy' | summarize count() by resourceGroup | top 5 by resourceGroup | project resourceGroup , count_"
$WordDocument = New-WordDocument -FilePath ".\Azure-Infrastructure-Report.docx"
Add-WordTOC -WordDocument $WordDocument
Add-WordPageBreak -WordDocument $WordDocument
Add-WordText -WordDocument $WordDocument -Text "Top 5 Azure Resource type by Count" -HeadingType Heading3 -Color Black -Alignment center
Add-WordText -WordDocument $WordDocument -NewLine $true
DocText -LineBreak 
Add-WordBarChart -WordDocument $WordDocument -ChartName "Azure Resources Count by Type" -Names $($unHealthyResources.resourceGroup) -Values $($unHealthyResources.count_) -ChartLegendPosition Left -ChartLegendOverlay $false -BarDirection Column
DocText -LineBreak 
#endregion
```

> Note: In your script file add the below three lines of code - This is required for generating the document and for Az Graph Query

```PowerShell
#region
Install-PackageProvider -Name NuGet -Force -Scope CurrentUser
Install-Module -Name PSWriteWord -Force -Verbose -Scope CurrentUser
Install-Module -Name Az.ResourceGraph -Force -Verbose -Scope CurrentUser
#endregion
```

Isn’t it easy work?  Yes, it is! The complete YAML pipeline is below 

```YML
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
              ScriptPath: 'src\azureInventory.ps1'
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
              filePath: 'src\sendEmail.ps1'
              arguments: '-Attachments "$(System.ArtifactsDirectory)\azureInventory\Azure Infrastructure Report.docx" -MailID "$(MailID)" -MailPassword "$(MailPassword)"'
```