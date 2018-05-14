function New-WrappedSQLQueryRoute
    {
    [cmdletbinding()]
    Param (
        [string]$Path,
        [string]$SQLConnectionString,
        [string]$SQLQuery,
        [string]$SQLQueryFile,
        [ValidateSet( 'JSON', 'CSV', 'TXT', 'HTML' )]
        [string]$OutputType = 'JSON',
        [string]$Method = 'Get',
        [switch]$Force,
        $Polaris )

    #  If SQLQueryFile was specified
    #    Load SQLQueryFile contents into $SQLQuery
    If ( $SQLQueryFile )
        {
        If ( Test-Path -Path $SQLQueryFile )
            {
            $SQLQuery = Get-Content -Path $SQLQueryFile -Raw
            }
        }

    #  Define the base route script
    #  (We'll tweak it below to hardcode the specified parameter values.)
    $BaseScriptblock =
        {
        #  "Parameters"

        <#Parameter#>$SQLConnectionString = ''
        <#Parameter#>$SQLQuery = ''
        <#Parameter#>$OutputType = ''

        #region  Functions

        function Invoke-SqlQuery
            {
            [cmdletbinding()]
            Param ( [string]                             $Query, 
                    [System.Collections.Hashtable]       $Parameters, 
                    [string]                             $ConnectionString,
                    [System.Data.SqlClient.SqlConnection]$SQLConnection,
                    [int]                                $TimeOut = 60,
                    [System.Data.CommandType]            $CommandType = 'Text'
                    )
            try
                {
                If ( $SQLConnection )
                    {
                    $SQLConnection = $SQLConnection.Clone()
                    }
                Else
                    {
                    $SQLConnection = New-Object System.Data.SQLClient.SQLConnection $ConnectionString
                    }

                #  Create the SQL command object as specified
                $SQLCommand = $SQLConnection.CreateCommand()

                $SQLCommand.CommandText    = $Query
                $SQLCommand.CommandType    = $CommandType
		        $SQLCommand.CommandTimeout = $TimeOut
                If ( $Parameters )
                    {
                    ForEach ( $Key in $Parameters.Keys )
                        {
                        [void]$SQLCommand.Parameters.AddWithValue( $Key, $Parameters[$Key] )
                        }
                    }

                #  Execute the Command
                $SQLConnection.Open()

                $SQLReader = $SQLCommand.ExecuteReader()

                $Datatable = New-Object System.Data.DataTable
                While ( $SQLReader.VisibleFieldCount )
                    {
                    try { $DataTable.Load( $SQLReader ) } catch {}
                    }

                return $DataTable | Select $DataTable.Columns.ColumnName
                }
            finally
                {
                If ( $SQLConnection -and $SQLConnection.State -ne [System.Data.ConnectionState]::Closed )
                    {
                    $SQLConnection.Close()
                    }
                }
            }

        #endregion

        try
            {
            $RequestQueryKeys = [System.Collections.ArrayList]$Request.Query.Keys

            #  If request query specifies output type
            #    Override default
            If ( $RequestQueryKeys -contains 'OutputType' )
                {
                #  If requested output type is valid
                #    Set output type as requested 
                If ( $Request.Query['OutputType'] -in 'JSON', 'CSV', 'TXT', 'HTML' )
                    {
                    $OutputType = $Request.Query['OutputType']

                    #  Remove OutputType from list to pass on to command
                    $RequestQueryKeys.Remove( 'OutputType' )
                    }
                }

            #  Initialize collection
            $SQLParameters = @{}

            #  If parameters were supplied in query...
            If ( $RequestQueryKeys.Count )
                {
                #  For each parameter in query...
                ForEach ( $Key in $RequestQueryKeys )
                    {
                    [void]$SQLParameters.Add( $Key, $Request.Query[$Key] )
                    }
                }

            #  Invoke SQL query
            $Results = Invoke-SQLQuery -Query $SQLQuery -ConnectionString $SQLConnectionString -Parameters $SQLParameters

            #  If results were received from SQL
            #    Return results as requested output type
            If ( $Results )
                {
                switch ( $OutputType )
                    {
                    'JSON'
                        {
                        #  Convert results to JSON
                        $ResultsJson = $Results | ConvertTo-Json -Depth 1 -Compress
                
                        #  Return results
                        $Response.Json( $ResultsJson )
                        }
                    
                    'CSV'
                        {
                        #  Convert results to CSV
                        $ResultsCSV = $Results | ConvertTo-CSV -NoTypeInformation
                
                        #  Set response type
                        $Response.ContentType = 'text/csv'
                        
                        #  Return results
                        $Response.Send( $ResultsCSV )
                        }

                    'TXT'
                        {
                        #  Convert results to CSV
                        $ResultsCSV = $Results | ConvertTo-CSV -NoTypeInformation
                
                        #  Set response type
                        $Response.ContentType = 'text/txt'
                        
                        #  Return results
                        $Response.Send( $ResultsCSV )
                        }

                    'HTML'
                        {
                        #  Convert results to CSV
                        $ResultsHTML = $Results | ConvertTo-Html -Fragment
                
                        #  A little bit of pretty for the directory
                        $Style = '<style>
                            table { border-collapse: collapse; }
                            table, th, td { border: 1px solid black; }
                            th, td { padding: 3px; } </style>'

                        #  Convert the collection to an HTML table
                        #    Tweak it so it display properly
                        #    Right align the file length cells
                        $ResultsHTML = $Style + ( $Results | ConvertTo-Html -Fragment )

                        #  Set the content type
                        $Response.ContentType = 'text/html'

                        #  Return the directoy contents
                        $Response.Send( $ResultsHTML )
                        }
                    }
                }
            }
 
        #  Error
        #    Return it
        catch
            {
            $Response.StatusCode = 502
            }
        }

    #  Convert script to string for tweaking
    #  (It's a scriptblock above for ease of reading and editing.)
    $BaseScript = [string]$BaseScriptblock

    #  Inject parameters
    $BaseScript = $BaseScript.ToString().Replace( "<#Parameter#>`$SQLConnectionString = ''", "<#Parameter#>`$SQLConnectionString = '$SQLConnectionString'" )
    $BaseScript = $BaseScript.ToString().Replace( "<#Parameter#>`$SQLQuery = ''", ( "<#Parameter#>`$SQLQuery = @'" + [environment]::NewLine + $SQLQuery + [environment]::NewLine + "'@" ) )
    $BaseScript = $BaseScript.ToString().Replace( "<#Parameter#>`$OutputType = ''", "<#Parameter#>`$OutputType = '$OutputType'" )

    #  Convert script back to scriptblock
    $Scriptblock = [scriptblock]::Create( $BaseScript )

    #  Default parameters for route
    $Route = @{
        Path        = $Path
        Method      = $Method
        Scriptblock = $Scriptblock }
        
    #  Optional parameters for route
    If ( $Polaris ) { $Route += @{ Polaris = $Polaris } }
    If ( $Force   ) { $Route += @{ Force   = $True    } }

    #  Create route
    New-PolarisRoute @Route
    }


function New-WrappedScriptRoute
    {
    [cmdletbinding()]
    Param (
        [string]$Path,
        [string]$ScriptPath,
        [int]$Depth = 1,
        [string]$Method = 'Get',
        [switch]$Force,
        $Polaris )

    #  Define the base route script
    #  (We'll tweak it below to hardcode the specified parameter values.)
    $BaseScriptblock =
        {
        #  "Parameters"

        <#Parameter#>$ScriptPath = ''
        <#Parameter#>$Depth = 1

        try
            {
            #  If parameters were supplied in query...
            If ( $Request.Query.Keys.Count )
                {

                #  For each parameter in query...
                $Parameters = @{}
                ForEach ( $Key in $Request.Query.Keys )
                    {

                    #  Add parameters to parameter collection
                    #  Convert text true/false to boolean true/false
                    #  (Rely on PowerShell to make any other required conversions)
                    switch ( $Request.Query[$Key] )
                        {
                        'True'  { [void]$Parameters.Add( $Key, $True                ); break }
                        'False' { [void]$Parameters.Add( $Key, $False               ); break }
                        default { [void]$Parameters.Add( $Key, $Request.Query[$Key] )        }
                        }
                    }

                #  Invoke script with parameters
                $Results = . $ScriptPath @Parameters
                }

            #  Else (no parameters supplied in query)...
            Else
                {
                #  Invoke script without parameters
                $Results = . $ScriptPath
                }

            #  If results were received from invoked script...
            If ( $Results )
                {
                #  Convert results to JSON
                $ResultsJson = $Results | ConvertTo-Json -Depth $Depth -Compress
                
                #  Return results
                $Response.Json( $ResultsJson )
                }
            }

        #  Error
        #    Return it
        catch
            {
            $Response.StatusCode = 502
            }
        }

    #  Convert script to string for tweaking
    #  (It's a scriptblock above for ease of reading and editing.)
    $BaseScript = [string]$BaseScriptblock

    #  Inject ScriptPath parameter
    $BaseScript = $BaseScript.ToString().Replace( "<#Parameter#>`$ScriptPath = ''", "<#Parameter#>`$ScriptPath = '$ScriptPath'" )
    $BaseScript = $BaseScript.ToString().Replace( "<#Parameter#>`$Depth = 1", "<#Parameter#>`$Depth = '$Depth'" )

    #  Convert script back to scriptblock
    $Scriptblock = [scriptblock]::Create( $BaseScript )

    #  Default parameters for route
    $Route = @{
        Path        = $Path
        Method      = $Method
        Scriptblock = $Scriptblock }
        
    #  Optional parameters for route
    If ( $Polaris ) { $Route += @{ Polaris = $Polaris } }
    If ( $Force   ) { $Route += @{ Force   = $True    } }

    #  Create route
    New-PolarisRoute @Route
    }


function New-WrappedCommandRoute
    {
    [cmdletbinding()]
    Param (
        [string]$Path,
        [string]$WrappedCommand,
        [int]$Depth = 1,
        [string[]]$SelectProperties = '*',
        [string]$Method = 'Get',
        [switch]$Force,
        $Polaris )

    #  Define the base route script
    #  (We'll tweak it below to hardcode the specified parameter values.)
    $BaseScriptblock =
        {
        #  "Parameters"

        <#Parameter#>$WrappedCommand = ''
        <#Parameter#>$Depth = 1
        <#Parameter#>$SelectProperties = ''.Split( ',' )

        try
            {
            $RequestQueryKeys = [System.Collections.ArrayList]$Request.Query.Keys

            #  If request query specifies properties to select
            #    Override default
            If ( $RequestQueryKeys -contains 'SelectProperties' )
                {
                $SelectProperties = $Request.Query['SelectProperties'].Split( ',' )

                #  Remove SelectProperties from list to pass on to command
                $RequestQueryKeys.Remove( 'SelectProperties' )
                }

            #  If parameters were supplied in query...
            If ( $RequestQueryKeys.Count )
                {

                $ParamDef = (Get-Command $WrappedCommand).Parameters

                #  For each parameter in query...
                $Parameters = @{}
                ForEach ( $Key in $RequestQueryKeys )
                    {

                    #  If query parameter matches a command parameter...
                    If ( $Key -in $ParamDef.Keys )
                        {

                        #  If command parameter type is an array
                        #    Assume comma-delimited string
                        #    Split into array on commas and add to parameters
                        If ( $ParamDef[$Key].ParameterType.IsArray )
                            {
                            [void]$Parameters.Add( $Key, $Request.Query[$Key].Split( ',' ) )
                            }

                        #  If command parameter type is a switch...
                        ElseIf ( $ParamDef[$Key].ParameterType.Name -eq 'SwitchParameter' )
                            {
                            #  If query parameter is set to string true or is present without a value
                            #    Add to parameters as $True
                            If ( $Request.Query[$Key] -eq 'True' -or $Request.Query[$Key] -eq $Null )
                                {
                                [void]$Parameters.Add( $Key, $True )
                                }

                            #  Else (Switch parameter is not set to true or present without a value)
                            #    Add to parameters as $False
                            Else
                                {
                                [void]$Parameters.Add( $Key, $False )
                                }
                            }

                        #  If command parameter type is a boolean...
                        ElseIf ( $ParamDef[$Key].ParameterType.Name -eq 'Boolean' )
                            {
                            #  Convert to boolean and add to parameters
                            [void]$Parameters.Add( $Key, $Request.Query[$Key] -eq 'True' )
                            }

                        #  Else (other parameter type)
                        #    Add to parameters as is
                        #    (Rely on PowerShell to do any required conversion)
                        Else
                            {
                            [void]$Parameters.Add( $Key, $Request.Query[$Key] )
                            }
                        }
                    }

                #  Invoke command with parameters
                $Results = Invoke-Expression "$WrappedCommand @Parameters"
                }

            #  Else (no parameters supplied in query)...
            Else
                {
                #  Invoke command without parameters
                $Results = Invoke-Expression $WrappedCommand
                }

            #  If results were received from invoked command...
            If ( $Results )
                {
                #  Convert results to JSON
                $ResultsJson = $Results |
                    Select-Object -Property $SelectProperties |
                    ConvertTo-Json -Depth $Depth -Compress
                
                #  Return results
                $Response.Json( $ResultsJson )
                }
            }
 
        #  Error
        #    Return it
        catch
            {
            $Response.StatusCode = 502
            }
        }

    #  Convert script to string for tweaking
    #  (It's a scriptblock above for ease of reading and editing.)
    $BaseScript = [string]$BaseScriptblock

    #  Inject Command and Depth parameters
    $BaseScript = $BaseScript.ToString().Replace( "<#Parameter#>`$WrappedCommand = ''", "<#Parameter#>`$WrappedCommand = '$WrappedCommand'" )
    $BaseScript = $BaseScript.ToString().Replace( "<#Parameter#>`$SelectProperties = ''", "<#Parameter#>`$SelectProperties = '$($SelectProperties -join ',')'" )
    $BaseScript = $BaseScript.ToString().Replace( "<#Parameter#>`$Depth = 1", "<#Parameter#>`$Depth = $Depth" )

    #  Convert script back to scriptblock
    $Scriptblock = [scriptblock]::Create( $BaseScript )

    #  Default parameters for route
    $Route = @{
        Path        = $Path
        Method      = $Method
        Scriptblock = $Scriptblock }
        
    #  Optional parameters for route
    If ( $Polaris ) { $Route += @{ Polaris = $Polaris } }
    If ( $Force   ) { $Route += @{ Force   = $True    } }

    #  Create route
    New-PolarisRoute @Route
    }


function New-SQLReportSite
    {
    <#
        .SYNOPSIS
            Create directory browsing route for folder and SQL Query routes for SQL queries within

        .DESCRIPTION
            If local SQLConnectionString.xml
                Import local SQLConnectionString.xml
                Set SQLConnectionString.xml attribute to hidden
            If root folder
                Create dummy Rebuild file, as needed
                Create route for Rebuild file
                Create directory browsing route for root folder
            For each child .sql file
                Create SQL query route for .sql file
            For each child subfolder
                Recurse

        .PARAMETER  Path
            String - Relative path of the URL to map

        .PARAMETER  Folder
            String - Full path and name of the NTFS folder to map

        .COMPONENT
            Module - Polaris

        .NOTES
            Designed to enhance modified version of Polaris 0.7

        .NOTES
            v1.0  5/8/2018  Tim Curwick  Created
    #>
    [cmdletbinding()]
    Param (
        [string]$Path,
        [string]$Folder )

    New-SQLReportSiteHelper -Path $Path -Folder $Folder
    }


function New-SQLReportSiteHelper
    {
    [cmdletbinding()]
    Param (
        [string]$Path,
        [string]$Folder,
        [switch]$IsSubfolder )

    #  Clean paths
    $Path = $Path.TrimEnd( '/' )
    $Folder = $Folder.TrimEnd( '\' )

    #  If SQL connection string is present in folder
    #    Process it
    If ( Test-Path -Path "$Folder\SQLConnectionString.xml" )
        {
        #  Import connection string
        $ConStringSecure = Import-Clixml -Path "$Folder\SQLConnectionString.xml"
        $SQLConnectionString = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR( $ConStringSecure ) )
        
        #  If SQL connection string file is not hidden
        #    Hide it
        $SQSFile = Get-Item -Path "$Folder\SQLConnectionString.xml" -Force
        If ( -not ( $SQSFile.Attributes -band 'Hidden' ) )
            {
            $SQSFile.Attributes = $SQSFile.Attributes -bor 'Hidden'
            }
        }

    #  Else (SQL connection string is not present in folder)
    Else
        {
        #  $SQLConnectionString from parent scope will be used
        }

    #  If this is the root reports folder...
    If ( -not $IsSubfolder )
        {
        #  Create dummy Rebuild file as needed
        If ( -not ( Test-Path -Path "$Folder\Rebuild this report site" ) )
            {
            New-Item -Path $Folder -Name 'Rebuild this report site' -ItemType File
            }

        #  Create Polaris route pointing dummy rebuild file to this function
        New-PolarisRoute `
            -Path        "$Path/Rebuild this report site" `
            -Method      Get `
            -Scriptblock ( [scriptblock]::Create( "New-SQLReportSite -Path '$Path' -Folder '$Folder'; `$Response.Send( 'Rebuild complete' )" ) ) `
            -Force

        #  Create Polaris directory browsing route at reporting root
        New-PolarisStaticRoute -RoutePath $Path -FolderPath $Folder -Force
        }

    #  Get all reports (.sql files) in folder
    $Reports = Get-ChildItem -Path "$Folder\*.sql"

    #  For each report (.SQL file)...
    ForEach ( $Report in $Reports )
        {
        #  Create Polaris SQL Query route for report
        New-WrappedSQLQueryRoute `
            -Path                "$Path/$($Report.Name)" `
            -SQLConnectionString $SQLConnectionString `
            -SQLQueryFile        $Report.FullName `
            -OutputType          HTML `
            -Force
        }

    #  Recurse subfolders
    $Subfolders = Get-ChildItem -Path $Folder -Directory
    ForEach ( $Subfolder in $Subfolders )
        {
        New-SQLReportSiteHelper -Path "$Path/$($Subfolder.Name)" -Folder $Subfolder.FullName -IsSubfolder
        }
    }


function Build-SQLLogSite
    {
    [cmdletbinding()]
    Param (
        [string]$Path,
        [string]$Folder,
        [string[]]$SQLServer )

    #  Delete existing dummy folder tree in root folder
    #  (Will error if folder tree contains files)
    Get-ChildItem -Path $Folder -Directory -Recurse | Remove-Item -Force

    #  Create Polaris directory browsing route for root folder
    New-PolarisStaticRoute -RoutePath $Path -FolderPath $Folder -Force

    #  Create dummy Rebuild file as needed
    If ( -not ( Test-Path -Path "$Folder\Rebuild this report site" ) )
        {
        New-Item -Path $Folder -Name 'Rebuild this report site' -ItemType File
        }

    #  Create Polaris route pointing dummy rebuild file to this function
    New-PolarisRoute `
        -Path        "$Path/Rebuild this report site" `
        -Method      Get `
        -Scriptblock ( [scriptblock]::Create( "Build-SQLLogSite -Path '$Path' -Folder '$Folder' -SQLServer '$($SQLServer -join ',')'.Split( ',' ); `$Response.Send( 'Rebuild complete' )" ) ) `
        -Force

    #  For each SQL server to process...
    ForEach ( $SQLServerName in $SQLServer )
        {
        #  Create folder in dummy browsing tree for SQL server
        New-Item -Path $Folder -Name $SQLServerName -ItemType Directory

        #  Get SQL server instance folders
        $SQLFolders = Get-ChildItem -Path "\\$SQLServerName\D$\ProgramFiles\Microsoft SQL Server" -Directory

        #  For each SQL server instance folder
        ForEach ( $SQLFolder in $SQLFolders )
            {
            #  Derive instance name and path to log folder
            $Instance = $SQLFolder.Name.Split( '.' )[-1]
            $LogPath = "$($SQLFolder.FullName)\MSSQL\Log"

            #  If expected instance log folder exists...
            If ( Test-Path $LogPath )
                {
                #  Create folder in dummy browsing tree for SQL server instance
                New-Item -Path "$Path/$SQLServerName" -Name $Instance -ItemType Directory

                #  Create Polaris directory browsing route pointing dummy instance folder to instance log folder
                New-PolarisStaticRoute -RoutePath "$Path/$SQLServerName/$Instance" -FolderPath $LogPath -Force
                }
            }
        }
    }
