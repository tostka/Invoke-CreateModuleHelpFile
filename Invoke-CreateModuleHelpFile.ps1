#*------v Function Invoke-CreateModuleHelpFile v------
function Invoke-CreateModuleHelpFile {
    <#
    .SYNOPSIS
    Invoke-CreateModuleHelpFile.ps1 - Create a HTML help file for a PowerShell module.
    .NOTES
    Version     : 1.2.1
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2023-
    FileName    : Invoke-CreateModuleHelpFile.ps1
    License     : MIT License
    Copyright   : (c) 2023 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : Øyvind Kallstad @okallstad
    AddedWebsite: https://communary.net/
    AddedTwitter: @okallstad / https://twitter.com/okallstad
    REVISIONS
    * 9:17 AM 9/29/2023 rewrote to support conversion for scripts as well; added 
    -script & -nopreview params (as it now also auto-previews in default browser);  
    ould be to move the html building code into a function, and leave the module /v script logic external to that common process.
    expanded CBH; put into OTB & advanced function format; split trycatch into beg & proc blocks
    10/18/2014 OK's posted rev 1.1
    .DESCRIPTION
    Invoke-CreateModuleHelpFile.ps1 - Create a HTML help file for a PowerShell module.
    This function will generate a full HTML help file for all commands in a PowerShell module.

    This function is dependent on jquery, the bootstrap framework and the jasny bootstrap add-on.


    .PARAMETER ModuleName
    Name of module. Note! The module must be imported before running this function[-ModuleName myMod]
    .PARAMETER Destination
    Full path and filename to the generated html helpfile[-Destination c:\pathto\MyModuleHelp.html]
    .PARAMETER SkipDependencyCheck
    Skip dependency check[-SkipDependencyCheck] 
    .PARAMETER Script
    Switch for processing target Script files (vs Modules)[-Script]
    .PARAMETER NoPreview
    Switch to suppress trailing preview of html in default browser[-NoPreview]
    .PARAMETER whatIf
    Whatif Flag  [-whatIf]
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    None. Returns no objects or output (.NET types)
    System.Boolean
    [| get-member the output to see what .NET obj TypeName is returned, to use here]
    .EXAMPLE
    PS> Invoke-CreateModuleHelpFile -ModuleName 'verb-text' -Dest 'c:\temp\verb-text_HLP.html' -verbose ; 
    Generate Html Help file for 'verb-text' module and save it as 'c:\temp\verb-text_HLP.html' with verbose output.
    .EXAMPLE
    PS> Invoke-CreateModuleHelpFile -ModuleName 'c:\usr\work\ps\scripts\move-ConvertedVidFiles.ps1' -Script -Dest 'c:\temp\'  -verbose ; 
    Generate Html Help file for the 'move-ConvertedVidFiles.ps1' script and save it as with a generated default name (move-ConvertedVidFiles_HELP.html) to the 'c:\temp\' directory with verbose output.
    EXDESCRIPTION
    .LINK
    https://github.com/tostka/Invoke-CreateModuleHelpFile
    .LINK
    https://github.com/gravejester/Invoke-CreateModuleHelpFile
    .LINK
    [ name related topic(one keyword per topic), or http://|https:// to help, or add the name of 'paired' funcs in the same niche (enable/disable-xxx)]
    #>
    [CmdletBinding()]
    PARAM(
        # Name of module. Note! The module must be imported before running this function.
        [Parameter(Mandatory = $true,HelpMessage="Name of module. Note! The module must be imported before running this function[-ModuleName myMod]")]
            [ValidateNotNullOrEmpty()]
            [string] $ModuleName,
        # Full path and filename to the generated html helpfile.
        [Parameter(Mandatory = $true,HelpMessage="Full path and filename to the generated html helpfile[-Path c:\pathto\MyModuleHelp.html]")]
            [ValidateScript({Test-Path $_ })]
            [string] $Destination,
        [Parameter(HelpMessage="Skip dependency check[-SkipDependencyCheck]")]
            [switch] $SkipDependencyCheck,
        [Parameter(HelpMessage="Switch for processing target Script files (vs Modules)[-Script]")]
            [switch] $Script,
        [Parameter(HelpMessage="Switch to suppress trailing preview of html in default browser[-NoPreview]")]
            [switch] $NoPreview
    ) ; 
    BEGIN{
        #region CONSTANTS-AND-ENVIRO #*======v CONSTANTS-AND-ENVIRO v======
        # function self-name (equiv to script's: $MyInvocation.MyCommand.Path) ;
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        $PSParameters = New-Object -TypeName PSObject -Property $PSBoundParameters ;
        write-verbose "`$PSBoundParameters:`n$(($PSBoundParameters|out-string).trim())" ;
        $Verbose = ($VerbosePreference -eq 'Continue') ; 
        #region ENVIRO_DISCOVER ; #*------v ENVIRO_DISCOVER v------
        #if ($PSScriptRoot -eq "") {
        # 8/29/2023 fix logic break on psv2 ISE (doesn't test PSScriptRoot -eq '' properly, needs $null test).
        #if( -not (get-variable -name PSScriptRoot -ea 0) -OR ($PSScriptRoot -eq '')){
        if( -not (get-variable -name PSScriptRoot -ea 0) -OR ($PSScriptRoot -eq '') -OR ($PSScriptRoot -eq $null)){
            if ($psISE) { $ScriptName = $psISE.CurrentFile.FullPath } 
            elseif($psEditor){
                if ($context = $psEditor.GetEditorContext()) {$ScriptName = $context.CurrentFile.Path } 
            } elseif ($host.version.major -lt 3) {
                $ScriptName = $MyInvocation.MyCommand.Path ;
                $PSScriptRoot = Split-Path $ScriptName -Parent ;
                $PSCommandPath = $ScriptName ;
            } else {
                if ($MyInvocation.MyCommand.Path) {
                    $ScriptName = $MyInvocation.MyCommand.Path ;
                    $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent ;
                } else {throw "UNABLE TO POPULATE SCRIPT PATH, EVEN `$MyInvocation IS BLANK!" } ;
            };
            if($ScriptName){
                $ScriptDir = Split-Path -Parent $ScriptName ;
                $ScriptBaseName = split-path -leaf $ScriptName ;
                $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($ScriptName) ;
            } ; 
        } else {
            if($PSScriptRoot){$ScriptDir = $PSScriptRoot ;}
            else{
                write-warning "Unpopulated `$PSScriptRoot!" ; 
                $ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
            }
            if ($PSCommandPath) {$ScriptName = $PSCommandPath } 
            else {
                $ScriptName = $myInvocation.ScriptName
                $PSCommandPath = $ScriptName ;
            } ;
            $ScriptBaseName = (Split-Path -Leaf ((& { $myInvocation }).ScriptName))  ;
            $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
        } ;
        if(-not $ScriptDir){
            write-host "Failed `$ScriptDir resolution on PSv$($host.version.major): Falling back to $MyInvocation parsing..." ; 
            $ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
            $ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ; 
            $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;     
        } else {
            if(-not $PSCommandPath ){
                $PSCommandPath  = $ScriptName ; 
                if($PSCommandPath){ write-host "(Derived missing `$PSCommandPath from `$ScriptName)" ; } ;
            } ; 
            if(-not $PSScriptRoot  ){
                $PSScriptRoot   = $ScriptDir ; 
                if($PSScriptRoot){ write-host "(Derived missing `$PSScriptRoot from `$ScriptDir)" ; } ;
            } ; 
        } ; 
        if(-not ($ScriptDir -AND $ScriptBaseName -AND $ScriptNameNoExt)){ 
            throw "Invalid Invocation. Blank `$ScriptDir/`$ScriptBaseName/`ScriptNameNoExt" ; 
            BREAK ; 
        } ; 

        $smsg = "`$ScriptDir:$($ScriptDir)" ;
        $smsg += "`n`$ScriptBaseName:$($ScriptBaseName)" ;
        $smsg += "`n`$ScriptNameNoExt:$($ScriptNameNoExt)" ;
        $smsg += "`n`$PSScriptRoot:$($PSScriptRoot)" ;
        $smsg += "`n`$PSCommandPath:$($PSCommandPath)" ;  ;
        write-verbose $smsg ; 
        #endregion ENVIRO_DISCOVER ; #*------^ END ENVIRO_DISCOVER ^------

        # jquery filename - remember to update if you update jquery to a newer version
        $jqueryFileName = 'jquery-1.11.1.min.js'

        # define dependencies
        $dependencies = @('bootstrap.min.css','jasny-bootstrap.min.css','navmenu.css',$jqueryFileName,'bootstrap.min.js','jasny-bootstrap.min.js')

        TRY {
            # check dependencies - revise pathing to $ScriptDir (don't have to run pwd the mod dir)
            if (-not($SkipDependencyCheck)) {
                $missingDependency = $false
                foreach($dependency in $dependencies) {
                    #if(-not(Test-Path -Path ".\$($dependency)")) {
                    if(-not(Test-Path -Path (join-path -path $scriptdir -ChildPath $dependency))) {
                        Write-Warning "Missing: $($dependency)"
                        $missingDependency = $true
                    }
                }
                if($missingDependency) { break }
                Write-Verbose 'Dependency check OK'
            } ; 

            # add System.Web - used for html encoding
            Add-Type -AssemblyName System.Web ; 
        } CATCH {
            Write-Warning $_.Exception.Message ; 
        } ; 
    }  ;  # BEG-E
    PROCESS {
        $Error.Clear() ; 
    
        foreach($ModName in $ModuleName) {
            $smsg = $sBnrS="`n#*------v PROCESSING : $($ModName) v------" ; 
            if($Script){
                $smsg = $smsg.replace(" v------", " (PS1 scriptfile) v------")
            } ; 
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;

            if($Script){
                TRY{
                    $gcmInfo = get-command -Name $ModName -ea STOP ;
                    $moduleData = [pscustomobject]@{
                        #Name = $gcminfo.name ; 
                        Name = $gcminfo.Source ; 
                    } ;
                } CATCH {
                    Write-Warning $_.Exception.Message ;
                    Continue ; 
                } ;  
            } else { 
                # 9:51 AM 9/29/2023 silly, try to find & forceload the mod:
                 # try to get module info from imported modules first
                 <# there's a risk of ipmo/import-module'ing scripts, as they execute when run through it. 
                    but gmo checks for a path on the -Name spec and throws:
                    Get-Module : Running the Get-Module cmdlet without ListAvailable parameter is not supported for module names that include a path. Name parameter has this 
                    which prevents running scripts - if most scripts would include a path to their target. 
                    if loading a scripot in the path, how do we detect it's not a functional module?
                    can detect path by running split-path
                #>
                if( (split-path $modName -ea 0) -OR ([uri]$modName).isfile){
                    # pathed module, this will throw an error in gmo and exit
                    # or string that evalutes as a uri IsFile
                    $smsg = "specified -ModuleName $($modname) is a pathed specification,"
                    $smsg += "`nand -Script parameter *has not been specified*!" ;
                    $smsg += "`nget-module will *refuse* to execute against a pathed Module -Name specification, and will abort this script!"
                    $smsg += "`nif the intent is to process a _script_, rather than a module, please include the -script parameter!" ; 
                    write-warning $smsg ; 
                    throw $smsg ; 
                    Continue ; 
                } else {
                    # unpathed spec, doesn't eval as [uri].isfile
                    # check for function def's in the target file?                     
                    <#$rgxFuncDef = 'function\s+\w+\s+\{'
                    if(get-childitem $modName -ea 0 | select-string | -pattern $rgxFuncDef){
                    
                    } ; 
                    #>
                    # of course a lot of scripts have internal functions, and still execute on iflv...
                    # *better! does it have an extension!
                    # insufficient, periods are permitted in module names (MS powershell modules frequently are dot-delimtied fq names).
                    # just in case do a gcm and check for result.source value
                    if((get-command -Name $modName -ea 0).source){
                        # it's possible to have scripts with same name as modules
                        # and in most cases modules should have .psm1 extension
                        # tho my old back-load module copies were named .ps1
                        # check for path-hosted file with gcm on the name
                        if($xgcm = get-command $modName -ea 0 ){
                            # then check if the file's extension is .ps1, and hard block any work with it
                            if( (get-childitem -path $xgcm.source).extension -eq '.ps1'){
                                $smsg = "specified -ModuleName $($modname) resolves to a pathed .ps1 file, and -Script parameter *has not been specified*!" ;
                                $smsg += "`nto avoid risk of ipmo'ing scripts (which will execute them, rather than load a target module), this item is being skipped" ; 
                                $smsg += "`nif the intent is to process a _script_, rather than a module, please include the -script parameter when using this specification!" ; 
                                write-warning $smsg ; 
                                throw $smsg ; 
                                Continue ; 
                            } else { 

                            } ; 
                        } else { 
                            
                        } ; 
                    } ; 
                }
                if($moduleData = Get-Module -Name $ModName -ErrorAction Stop){} else { 
                    write-verbose "unable to gmo $ModName : Attempting ipmo..." ; 
                    if($tmod = Get-Module $modname -ListAvailable){
                        TRY{import-module -force -Name $ModName -ErrorAction Stop
                        } CATCH {
                            Write-Warning $_.Exception.Message ;
                            Continue ; 
                        } ; 
                        if($moduleData = Get-Module -Name $ModName -ErrorAction Stop){} else { 
                            throw "Unable to gmo or ipmo $ModName!" ; 
                        } ; 
                    } else { 
                        throw "Unable to gmo -list $ModName!" ; 
                    } ; 
                } ; 
            }
            TRY{
                # abort if no module data returned
                if(-not ($moduleData)) {
                    Write-Warning "The module '$($ModName)' was not found. Make sure that the module is imported before running this function." ; 
                    break ; 
                } ; 

                # abort if return type is wrong
                #if(($moduleData.GetType()).Name -ne 'PSModuleInfo') {
                if($Script){
                    if($gcminfo.CommandType -eq 'ExternalScript'){
                        $moduleCommands = $gcminfo.source ; 
                    }else {
                        Write-Warning "The 'Script' specified - '$($ModName)' - did not return an gcm CommandType of 'ExternalScript'." ; 
                        continue ; 
                    } ; 
                } else { 
                    if(($moduleData.GetType()).Name -ne 'PSModuleInfo') {
                        Write-Warning "The module '$($ModName)' did not return an object of type PSModuleInfo." ; 
                        continue ; 
                    } ; 
                    # get module commands
                    $moduleCommands = $moduleData.ExportedCommands | Select-Object -ExpandProperty 'Keys'
                    Write-Verbose 'Got Module Commands OK' ; 
                } ; 

                # start building html
                $html = @"
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="">
    <meta name="author" content="">
    <title>$($ModName)</title>
    <link href="bootstrap.min.css" rel="stylesheet">
    <link href="jasny-bootstrap.min.css" rel="stylesheet">
    <link href="navmenu.css" rel="stylesheet">
    <!-- HTML5 shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
      <script src="https://oss.maxcdn.com/libs/respond.js/1.4.2/respond.min.js"></script>
    <![endif]-->
  </head>
  <body>
    <div class="navmenu navmenu-default navmenu-fixed-left offcanvas-sm hidden-print">
      <nav class="sidebar-nav" role="complementary">
      <a class="navmenu-brand visible-md visible-lg" href="#" data-toggle="tooltip" title="$($ModName)">$($ModName)</a>
      <ul class="nav navmenu-nav">
        <li><a href="#About">About</a></li>

"@ ; 

                # loop through the commands to build the menu structure
                $count = 0 ; 
                foreach($command in $moduleCommands) {
                    $count++ ; 
                    Write-Progress -Activity "Creating HTML for $($command)" -PercentComplete ($count/$moduleCommands.count*100) ; 
                    $html += @"
          <!-- $($command) Menu -->
          <li class="dropdown">
          <a href="#" class="dropdown-toggle" data-toggle="dropdown">$($command) <b class="caret"></b></a>
          <ul class="dropdown-menu navmenu-nav">
            <li><a href="#$($command)-Synopsis">Synopsis</a></li>
            <li><a href="#$($command)-Syntax">Syntax</a></li>
            <li><a href="#$($command)-Description">Description</a></li>
            <li><a href="#$($command)-Parameters">Parameters</a></li>
            <li><a href="#$($command)-Inputs">Inputs</a></li>
            <li><a href="#$($command)-Outputs">Outputs</a></li>
            <li><a href="#$($command)-Examples">Examples</a></li>
            <li><a href="#$($command)-RelatedLinks">RelatedLinks</a></li>
            <li><a href="#$($command)-Notes">Notes</a></li>
          </ul>
        </li>
        <!-- End $($command) Menu -->

"@ ; 
                } ; 

                # finishing up the menu and starting on the main content
                $html += @"
        <li><a class="back-to-top" href="#top"><small>Back to top</small></a></li>
      </ul>
    </nav>
    </div>
    <div class="navbar navbar-default navbar-fixed-top hidden-md hidden-lg hidden-print">
      <button type="button" class="navbar-toggle" data-toggle="offcanvas" data-target=".navmenu">
        <span class="icon-bar"></span>
        <span class="icon-bar"></span>
        <span class="icon-bar"></span>
      </button>
      <a class="navbar-brand" href="#">$($ModName)</a>
    </div>
    <div class="container">
      <div class="page-content">
        <!-- About $($ModName) -->
        <h1 id="About" class="page-header">About $($ModName)</h1>
        <div class="row">
          <div class="col-md-4 col-xs-4">
            Description<br>
            ModuleBase<br>
            Version<br>
            Author<br>
            CompanyName<br>
            Copyright
          </div>
          <div class="col-md-6 col-xs-6">
            $([System.Web.HttpUtility]::HtmlEncode($moduleData.Description))<br>
            $([System.Web.HttpUtility]::HtmlEncode($moduleData.ModuleBase))<br>
            $([System.Web.HttpUtility]::HtmlEncode($moduleData.Version))<br>
            $([System.Web.HttpUtility]::HtmlEncode($moduleData.Author))<br>
            $([System.Web.HttpUtility]::HtmlEncode($moduleData.CompanyName))<br>
            $([System.Web.HttpUtility]::HtmlEncode($moduleData.Copyright))
          </div>
        </div>
        <br>
        <!-- End About -->

"@ ; 

                # loop through the commands again to build the main content
                foreach($command in $moduleCommands) {
                    $commandHelp = Get-Help $command ; 
                    $html += @"
        <!-- $($command) -->
        <div class="panel panel-default">
          <div class="panel-heading">
            <h2 id="$($command)-Header">$($command)</h1>
          </div>
          <div class="panel-body">
            <h3 id="$($command)-Synopsis">Synopsis</h3>
            <p>$([System.Web.HttpUtility]::HtmlEncode($commandHelp.Synopsis))</p>
            <h3 id="$($command)-Syntax">Syntax</h3>

"@ ; 
                    # get and format the command syntax
                    $syntaxString = '' ; 
                    foreach($syntax in ($commandHelp.syntax.syntaxItem)) {
                        $syntaxString += "$($syntax.name)" ; 
                        foreach ($syntaxParameter in ($syntax.parameter)) {
                            $syntaxString += ' ' ; 
                            # parameter is required
                            if(($syntaxParameter.required) -eq 'true') {
                                $syntaxString += "-$($syntaxParameter.name)" ; 
                                if($syntaxParameter.parameterValue) { $syntaxString += " <$($syntaxParameter.parameterValue)>" } ; 
                            } else {
                                # parameter is not required
                                $syntaxString += "[-$($syntaxParameter.name)" ; 
                                if($syntaxParameter.parameterValue) { $syntaxString += " <$($syntaxParameter.parameterValue)>]" }
                                elseif($syntaxParameter.parameterValueGroup) { $syntaxString += " {$($syntaxParameter.parameterValueGroup.parameterValue -join ' | ')}]" } 
                                else { $syntaxString += ']' } ; 
                            } ; 
                        } ; 
                        $html += @"
            <pre>$([System.Web.HttpUtility]::HtmlEncode($syntaxString))</pre>

"@ ; 
                        Remove-Variable -Name 'syntaxString' ; 
                    } ; 

                    $html += @"
            <h3 id="$($command)-Description">Description</h3>
            <p>$([System.Web.HttpUtility]::HtmlEncode($commandHelp.Description.Text -join [System.Environment]::NewLine) -replace([System.Environment]::NewLine, '<br>'))</p>
            <h3 id="$($command)-Parameters">Parameters</h3>
            <dl class="dl-horizontal">

"@ ; 
                    # get all parameter data
                    foreach($parameter in ($commandHelp.parameters.parameter)) {
                        $parameterValueText = "<$($parameter.parameterValue)>" ; 
                        $html += @" 
              <dt data-toggle="tooltip" title="$($parameter.name)">-$($parameter.name)</dt>
              <dd>$([System.Web.HttpUtility]::HtmlEncode($parameterValueText))<br>
                $($parameter.description.Text)<br><br>
                <div class="row">
                  <div class="col-md-4 col-xs-4">
                    Required?<br>
                    Position?<br>
                    Default value<br>
                    Accept pipeline input?<br>
                    Accept wildchard characters?
                  </div>
                  <div class="col-md-6 col-xs-6">
                    $([System.Web.HttpUtility]::HtmlEncode($parameter.required))<br>
                    $([System.Web.HttpUtility]::HtmlEncode($parameter.position))<br>
                    $([System.Web.HttpUtility]::HtmlEncode($parameter.defaultValue))<br>
                    $([System.Web.HttpUtility]::HtmlEncode($parameter.pipelineInput))<br>
                    $([System.Web.HttpUtility]::HtmlEncode($parameter.globbing))
                  </div>
                </div>
                <br>
              </dd>

"@ ; 
                    } ; 

                    $html += @"
            </dl>
            <h3 id="$($command)-Inputs">Inputs</h3>
            <p>$([System.Web.HttpUtility]::HtmlEncode($commandHelp.inputTypes.inputType.type.name))</p>
            <h3 id="$($command)-Outputs">Outputs</h3>
            <p>$([System.Web.HttpUtility]::HtmlEncode($commandHelp.returnTypes.returnType.type.name))</p>
            <h3 id="$($command)-Examples">Examples</h3>

"@ ; 
                    # get all examples
                    $exampleCount = 0 ; 
                    foreach($commandExample in ($commandHelp.examples.example)) {
                        $exampleCount++ ; 
                        $html += @"
            <b>Example $($exampleCount.ToString())</b>
            <pre>$([System.Web.HttpUtility]::HtmlEncode($commandExample.code))</pre>
            <p>$([System.Web.HttpUtility]::HtmlEncode($commandExample.remarks.text -join [System.Environment]::NewLine) -replace([System.Environment]::NewLine, '<br>'))</p>
            <br>

"@ ; 
                    } ; 

                    $html += @"
            <h3 id="$($command)-RelatedLinks">RelatedLinks</h3>
            <p><a href="$([System.Web.HttpUtility]::HtmlEncode($commandHelp.relatedLinks.navigationLink.uri -join ''))">$([System.Web.HttpUtility]::HtmlEncode($commandHelp.relatedLinks.navigationLink.uri -join ''))</a></p>
            <h3 id="$($command)-Notes">Notes</h3>
            <p>$([System.Web.HttpUtility]::HtmlEncode($commandHelp.alertSet.alert.text -join [System.Environment]::NewLine) -replace([System.Environment]::NewLine, '<br>'))</p>
            <br>
          </div>
        </div>
        <!-- End ConvertFrom-HexIP -->

"@ ; 
                } ; 

                # finishing up the html
                $html += @"
        </div>
    </div><!-- /.container -->
    <script src="$($jqueryFileName)"></script>
"@ ; 
            $html += @'
    <script src="bootstrap.min.js"></script>
    <script src="jasny-bootstrap.min.js"></script>
    <script>$('body').scrollspy({ target: '.sidebar-nav' })</script>
    <script>
      $('[data-spy="scroll"]').on("load", function () {
        var $spy = $(this).scrollspy('refresh')
    })
    </script>
  </body>
</html>
'@ ; 

                Write-Verbose 'Generated HTML OK' ; 

                # if $modName is pathed, split it to the leaf
                if( (split-path $modName -ea 0) -OR ([uri]$modName).isfile){
                    write-verbose "converting pathed $($modname) to leaf..." ; 
                    #$leafFilename = split-path $modname -leaf ; 
                    $leaffilename = (get-childitem -path $modname).basename ; 
                } else {
                    $leafFilename = $modname ; 
                }; 

                #if(test-path -path $Destination -PathType container -ErrorAction SilentlyContinue){
                # test-path has a registry differentiating issue, safer to use gi!
                if( (get-item -path $Destination -ea 0).PSIsContainer){
                    $smsg = "-Destination specified - $($Destination) - is a container" ; 
                    $smsg += "`nconstructing output file as 'Name' $($leafFilename )_HELP.html..." ; 
                    write-host -ForegroundColor Yellow $smsg ; 
                    [System.IO.DirectoryInfo[]]$Destination = $Destination ; 
                    [system.io.fileinfo]$ofile = join-path -path $Destination.fullname -ChildPath "$($leafFilename)_HELP.html" ; 
                #}elseif(test-path -path $Destination -PathType Leaf -ErrorAction SilentlyContinue){
                }elseif( -not (get-item -path $Destination -ea 0).PSIsContainer){
                    [system.io.fileinfo]$Destination = $Destination ; 
                    if($Destination.extension -eq '.html'){
                        [system.io.fileinfo]$ofile = $Destination ; 
                    } else { 
                        throw "$($Destination) does *not* appear to have a suitable extension (.html):$($Destination.extension)" ; 
                    } ; 
                } else{
                    # not an existing dir (target) & not an existing file, so treat it as a full path
                    if($Destination.extension -eq 'html'){
                        [system.io.fileinfo]$ofile = $Destination ; 
                    } else { 
                        throw "$($Destination) does *not* appear to have a suitable extension (.html):$($Destination.extension)" ; 
                    } ; 
                } ; 
                write-host -ForegroundColor Yellow "Out-File -FilePath $($Ofile) -Encoding 'UTF8'" ; 

                # write html file
                $html | Out-File -FilePath $ofile.fullname -Force -Encoding 'UTF8' ; 
                Write-Verbose "$($ofile.fullname) written OK" ; 
                write-verbose "returning output path to pipeline" ; 
                $ofile.fullname | write-output ;
                if(-not $NoPreview){
                    write-host "Previewing $($ofile.fullname) in default browser..." ; 
                    Invoke-Item -Path $ofile.fullname ; 
                } ; 
            } CATCH {
                Write-Warning $_.Exception.Message ; 
            } ; 

            $smsg = "$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
        } ;  # loop-E
    } ;  # PROC-E
} ; 
#*------^ END Function Invoke-CreateModuleHelpFile ^------