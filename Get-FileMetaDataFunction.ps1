function Get-FileMetaData{
    [cmdletbinding()]
    param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('Fullname')]
        [string[]]$path
    )
    Begin{
        $oshell = New-Object -ComObject shell.Application
        #$objArray = @{}
    }
    Process{
        $path | ForEach-Object{
            # "Processing <$_>"
            if (Test-Path -Path $_ -PathType Leaf){
                $FileItem = Get-Item -Path $_
                $ofolder = $oshell.namespace($FileItem.DirectoryName)
                $ofile = $ofolder.parsename($FileItem.Name)
                
                $props = @{}

                0..287 | ForEach-Object{
                    $ExtPropName = $ofolder.GetDetailsof($ofolder.items, $_)
                    $ExtValName = $ofolder.GetDetailsof($ofile, $_)
                    if(-not $props.ContainsKey($ExtPropName) -and ($ExtValName -ne '')){
                        $props.Add($ExtPropName,$ExtValName)
                    }
                }
                New-Object psobject -Property $props | Tee-Object -Variable obj
            }
        }
    }
    End{
        $oshell = $null
    }
}