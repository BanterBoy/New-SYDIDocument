<#
.SYNOPSIS
    Short description
.DESCRIPTION
    Long description
.EXAMPLE
    Example of how to use this cmdlet
.EXAMPLE
    Another example of how to use this cmdlet
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    General notes
.COMPONENT
    The component this cmdlet belongs to
.ROLE
    The role this cmdlet belongs to
.FUNCTIONALITY
    The functionality that best describes this cmdlet
#>

function New-SYDIDocument {
    [CmdletBinding(DefaultParameterSetName='Default',
                   SupportsShouldProcess=$true,
                   PositionalBinding=$false,
                   HelpUri = 'https://github.com/BanterBoy/New-SYDIDocument',
                   ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([String])]
    Param (
        # help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='Default')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("cn")] 
        [string[]]
        $ComputerName,

        # help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='Default')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        [Alias("cred")]
        $Credential,

        # help description
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='Default')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("fs")] 
        [int]
        $FontSize = 12,

        # help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='Default')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("op")] 
        [string]
        $OutputPath
        
    )
    
    begin {
        $SYDIPath = "$PSScriptRoot\Sydi-Server.vbs"
        $Filename = $OutputPath + $ComputerName + ".docx"
        $UserName = $Credential.UserName
        $Password = $Credential.GetNetworkCredential().Password
    }
    
    process {
        if ($pscmdlet.ShouldProcess("Target", "Operation")) {
            cscript.exe $SYDIPath -wabefghipPqrsSu -racdklp -ew -f"$FontSize" -d -o"$FileName" -t"$ComputerName" -u"$Username" -p"$Password"
        }
    }
    
    end {
    }
}
