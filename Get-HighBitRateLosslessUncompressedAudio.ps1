function Get-HighBitRate([string]$folder=".",[int]$minBitrate=1144, [switch]$details=$true) {
$shellObject=New-Object -ComObject Shell.Application
$bitrateAttrib=0
$collection = @()
 
$audioTypes=Get-ChildItem $folder -include ('*.wav','*.flac','*.aiff','*.pcm','*.wma','*.m4a','*.alac','*.mqa') -recurse #hard-coded lossless and uncompressed audio formats

ForEach($audioType in $audioTypes) {
    # Get a shell object to retrieve file metadata.
    $directoryObject=$shellObject.NameSpace($audioType.Directory.FullName)
    $fileObject=$directoryObject.ParseName($audioType.Name)
 

    # Find the index of the bit rate attribute.
    For($index=5; -Not $bitrateAttrib;++$index) {
        $name=$directoryObject.GetDetailsOf($directoryObject.Items,$index)
        $name
        if($name -eq 'Bit rate') { $bitrateAttrib=$index }
    }
 
    # Get the bit rate of the file.
    $bitrateString=$directoryObject.GetDetailsOf($fileObject,$bitrateAttrib)
    if($bitrateString -match '\d+'){ 
        [int]$bitrate=$matches[0]
    }
    else { $bitrate=-1 }
 
    # Include the file in the results if it has the desired bit rate.
    if($bitrate -ge $minBitrate){ 
        #create PSObject to add properties to
        $custom_obj = new-object psobject        
        
        if ($details) { 
            #audioType
            $custom_obj | Add-Member -MemberType NoteProperty -Name "FileName" -Value $audioType.name
            $custom_obj | Add-Member -MemberType NoteProperty -Name "Path" -Value $audioType.Directory.FullName
            $custom_obj | Add-Member -MemberType NoteProperty -Name "Bitrate" -Value $bitratestring
        }else { 
            #$directoryObject
            $custom_obj | Add-Member -MemberType NoteProperty -Name "FileName" -Value $audioType.name
            $custom_obj | Add-Member -MemberType NoteProperty -Name "Path" -Value $audioType.Directory.FullName
            $custom_obj | Add-Member -MemberType NoteProperty -Name "Bitrate" -Value $bitratestring
            $collection += $custom_obj
        }
        $collection += $custom_obj 
    }
    }
    #this will display the collection to the screen. If you want to export to csv, pipe to export-csv 
    $collection 
}
