function New-ServerDocumentation {
    param (
        $ComputerName,
        $Username,
        $Password,
        $OutputPath
        )
        
        $SYDIPath = "$PSScriptRoot\sydi-server-2.4\Sydi-Server.vbs"
        $Filename = $OutputPath + $ComputerName + ".docx"

    cscript.exe $SYDIPath -wabefghipPqrsSu -racdklp -ew -f10 -d -o"$FileName" -t"$ComputerName" -u"$Username" -p"$Password"
    
}
