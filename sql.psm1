$noConnections      = 'Please fill the Connection string or SQLConnection!'
$noConnectionString = 'Please fill the Connection string!'
$noCommand          = 'Please fill the Command!'
$noTableName        = 'Please fill the neme of table!'
$noCommandType      = 'Please fill the CommandType parameter!'
$noCommandText      = 'Please fill the CommandText parameter!'
$noProcedureName    = 'Please fill the neme of procedure!'
$noQuery            = 'Please fill the query!'
$noDatabaseName     = 'Please fill the neme of database!'
$noFullFilePath     = 'Please fill the location of the backup file!'
$noSourceTableName  = 'Please fill the neme of source table!'
$noTargetTableName  = 'Please fill the neme of target table!'


filter Add-Property($Name, $Value) {
    Add-Member -type NoteProperty -name $Name -value $Value -InputObject $_ -PassThru
}

function Add-SQLCommandParameter {
    Param (
        [System.Data.SqlClient.SqlCommand] $Command = $(throw $noCommand),
        [Hashtable]                        $Params  = @{}
    )

    foreach($p in $Params.Keys) {
        $Command.Parameters.AddWithValue('@' + $p, $Params[$p]) | Out-Null 
        $paramString += $paramDelim + '@' + $p + '=' + $Params[$p]
        $paramDelim = ', '
    }
    
    Write-Verbose "SQL Command: Adding params $paramString"

    return $Command
}

function Inovke-Error {
    Param (
        [string] $message 
    )

    Write-Output "ERROR: $message"

    throw "ERROR: $message"
}

<#
.SYNOPSIS
    Open SQL connection by Connection string
#>
function Invoke-SQLStartConnection {
     Param (
        [String] $ConnectionString = $(throw $noConnectionString)
    )
    
    Write-Verbose 'SQL Connection: Creating connection'

    try {
        $sql = New-Object Data.SqlClient.SqlConnection
        $sql.ConnectionString = $ConnectionString
        $sql.Open() | Out-Null

        return $sql
    }
    catch [System.Exception] {
        Inovke-Error -message $_
    }
}


<#
.SYNOPSIS
    Close SQL connection
#>
function Invoke-SQLEndConnection {
     Param (
        [Data.SqlClient.SqlConnection] $SQLConnection = $(throw $noConnectionString)
    )
    
    Write-Verbose 'SQL Connection: Disposing connection'

    if ($SQLConnection) {
        $SQLConnection.Dispose() | Out-Null
    }
}


<#
.SYNOPSIS
    Make new SQL Command
    - Create SQL connection by connection string 
    - Create SQL command with parameters
#>
function Invoke-SQLStartCommand {
     Param (
        [Data.SqlClient.SqlConnection] $SQLConnection,

        [String]                       $ConnectionString,
        [string]                       $CommandType = $(throw $noCommandType),
        [int]                          $CommandTimeout = 30,
        [string]                       $CommandText = $(throw $noCommandText)
    )
    
    Write-Verbose "SQL Command: Creating new command"

    if ($ConnectionString) { 
        $SQLConnection = Invoke-SQLStartConnection -ConnectionString $ConnectionString
    }
    elseif (-not $SQLConnection) {
        throw $noConnections
    }

    try {
        $cmd = New-Object System.Data.SqlClient.SqlCommand
        $cmd.Connection = $SQLConnection
        $cmd = $SQLConnection.CreateCommand()
    
        $cmd.Parameters.Clear()
    
        $cmd.CommandType = $CommandType 
        $cmd.CommandTimeout = $CommandTimeout
        $cmd.CommandText = $CommandText
        
        return $cmd
    }
    catch [System.Exception] {
        Inovke-Error -message $_
    }
}


<#
.SYNOPSIS
    Close SQL Command and by parameter $CloseSQLConnection close or not active connection
#>
function Invoke-SQLEndCommand {
     Param (
        [System.Data.SqlClient.SqlCommand] $Command            = $(throw $noCommand),
        [bool]                             $CloseSQLConnection = $false
    )

    Write-Verbose "SQL Command: Disposing"
    
    if ($Command) {
        $Command.Dispose() | Out-Null
    }

    if ($CloseSQLConnection) {
        Invoke-SQLEndConnection -SQLConnection $Command.Connection
    }
}


<#
.SYNOPSIS
    Invoke stored procedure on DB
    - Create SQL connection with a SQL command and invoke procedure, after that close or not command and connection (by parameters)
#>
function Invoke-SQLStoredProcedure {
     Param (
        [Data.SqlClient.SqlConnection] $SQLConnection,

        [String]    $ConnectionString,
        [String]    $ProcedureName      = $(throw $noProcedureName),
        [Hashtable] $Params             = @{},
        [bool]      $CloseSQLConnection = $true
    )
    
    try 
    {
        if ($ConnectionString) {
            $SQLConnection = Invoke-SQLStartConnection -ConnectionString $ConnectionString
        }
        elseif (-not $SQLConnection) {
            throw $noConnections
        }

        $cmd = Invoke-SQLStartCommand -SQLConnection $SQLConnection -CommandType 'StoredProcedure' -CommandText $ProcedureName
            
        Add-SQLCommandParameter -Command $cmd -Params $Params
        
        Write-Verbose "SQL Stored procedure: Executing procedure $ProcedureName"
        
        $result = $cmd.ExecuteReader()

        $table = New-Object System.Data.DataTable
        $table.Load($result)

        return $table
    }
    catch [System.Exception] {
        Inovke-Error -message $_
    }
    finally {
        Invoke-SQLEndCommand -Command $cmd -CloseSQLConnection $CloseSQLConnection

        $result.Dispose()
    }
}


<#
.SYNOPSIS
    Invoke query into DB
    - Create SQL connection with SQL command and after query close command and connection by parameters
#>
function Invoke-SQLQuery {
    Param ( 
        [Data.SqlClient.SqlConnection] $SQLConnection,

	[string]    $ConnectionString,
        [string]    $Query              = $(throw $noQuery),
        [Hashtable] $Params             = @{},
        [bool]      $CloseSQLConnection = $true
	) 
    
    try {
        if ($ConnectionString) {
            $SQLConnection = Invoke-SQLStartConnection -ConnectionString $ConnectionString
        }
        elseif (-not $SQLConnection) {
            throw $noConnections
        }

        $cmd = Invoke-SQLStartCommand -SQLConnection $SQLConnection -CommandType 'Text' -CommandText $Query

	# if exist any parameter 
        if ($Params) {
            Add-SQLCommandParameter -Command $cmd -Params $Params
        }
    
        Write-Verbose "SQL Query: $Query"

        $result = $cmd.ExecuteReader()
        
        $table = New-Object System.Data.DataTable
        $table.Load($result)

        return $table
    }
    catch [System.Exception] {
        Inovke-Error -message $_
    }
    finally {
        Invoke-SQLEndCommand -Command $cmd -CloseSQLConnection $CloseSQLConnection

        $result.Dispose()
    }
}


<#
.SYNOPSIS
    Function for dropping table from DB
#>
function Invoke-SQLDropTable {
    Param (
        [string] $ConnectionString   = $(throw $noConnectionString),
        [string] $TableName          = $(throw $noTableName),
        [bool]   $CloseSQLConnection = $true
    )
    
    Write-Verbose "SQL Drop table: $TableName"
    
    $cmd = Invoke-SQLStartCommand -ConnectionString $ConnectionString -CommandType "Text" -CommandText "DROP TABLE $TableName"
    
    $cmd.ExecuteNonQuery() | Out-Null
    
    Invoke-SQLEndCommand -Command $cmd -CloseSQLConnection $CloseSQLConnection
}


<#
.SYNOPSIS
    Function for completly clean table (truncate)
#>
function Invoke-SQLTruncateTable {
    Param (
        [string] $ConnectionString   = $(throw $noConnectionString),
        [string] $TableName          = $(throw $noTableName),
        [bool]   $CloseSQLConnection = $true
    )
    
    Write-Verbose "SQL Truncate table: $TableName"
    
    $cmd = Invoke-SQLStartCommand -ConnectionString $ConnectionString -CommandType "Text" -CommandText "TRUNCATE TABLE $TableName"
    
    $cmd.ExecuteNonQuery() | Out-Null
    
    Invoke-SQLEndCommand -Command $cmd -CloseSQLConnection $CloseSQLConnection
}


<#
.SYNOPSIS
    Function for full backup DB into selected path
#>
function Invoke-SQLBackupDatabaseFull {
    Param (
        [string] $ConnectionString   = $(throw $noConnectionString),
        [string] $DatabaseName       = $(throw $noDatabaseName),
        [string] $FullFilePath       = $(throw $noFullFilePath),
        [bool]   $CloseSQLConnection = $true
    )
    
    Write-Verbose "SQL Backup Database: $DatabaseName to file: $FullFilePath"
    
    $cmd = Invoke-SQLStartCommand -ConnectionString $ConnectionString -CommandType "Text" -CommandText "BACKUP DATABASE $DatabaseName TO DISK = '$FullFilePath' WITH FORMAT, MEDIANAME = '$DatabaseName', NAME = 'Full Backup of $DatabaseName';"
    
    $cmd.ExecuteNonQuery() | Out-Null
    
    Invoke-SQLEndCommand -Command $cmd -CloseSQLConnection $CloseSQLConnection
}


<#
.SYNOPSIS
    Function which by name of table create copy of that table
#>
function Invoke-SQLCopyTable {
    Param (
        [string] $ConnectionString          = $(throw $noConnectionString),
        [string] $SourceTableName           = $(throw $noSourceTableName),
        [string] $TargetTableName           = $(throw $noTargetTableName),
        [string] $DeleteIfExistTargetTable  = $true,
        [bool]   $CloseSQLConnection        = $true
    )
    
    Write-Verbose "SQL Copy table: $SourceTableName to table: $TargetTableName"
    
    if ($DeleteIfExistTargetTable) {
        $deleteTable = "IF OBJECT_ID('$TargetTableName') IS NOT NULL DROP TABLE $TargetTableName;"
    } else { $deleteTable = "" }

    $cmd = Invoke-SQLStartCommand -ConnectionString $ConnectionString -CommandType "Text" -CommandText "$deleteTable SELECT * INTO $TargetTableName FROM $SourceTableName;"
    
    $cmd.ExecuteNonQuery() | Out-Null

    Invoke-SQLEndCommand -Command $cmd -CloseSQLConnection $CloseSQLConnection
}

#-----------------------------------------------------------------------------------------------------------------

Export-ModuleMember -Function @(
    'Invoke-SQLStartConnection'
    'Invoke-SQLEndConnection'
    'Invoke-SQLStartCommand'
    'Invoke-SQLEndCommand'
    'Invoke-SQLStoredProcedure'
    'Invoke-SQLQuery'
    'Invoke-SQLDropTable'
    'Invoke-SQLTruncateTable'
    'Invoke-SQLBackupDatabaseFull'
    'Invoke-SQLCopyTable'
)
