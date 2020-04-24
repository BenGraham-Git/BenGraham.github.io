####
# This is a custom script written for Candice Graham to open all images in a folder sequentially
# inside of ms paint and save each one in a new specified file format.
# Author : Benjamin Graham - www.BenGraham.co.za
####

#### Edit the below parameters as necessary ####
        
    ### Source Folder ###
    $srcfolder = "E:\WebDev\Powershell\CandiceImages"
   
    ### Destination Folder ###
    $destfolder = "E:\WebDev\Powershell\CandiceImages\Convert"

    ### Remove preceeding hash comment on the extension you require below ###
        ### Source file extension
        $src_filter = "*.jpg"
        #$src_filter = "*.jpeg"
        #$src_filter = "*.bmp"
        #$src_filter = "*.gif"

        ### Destination file extension
        $dest_ext = "jpg"
        #$dest_ext = "bmp"
        #$dest_ext = "png"
        #$dest_ext = "gif"

#### ----- End of user parameters ---- ####


#create conversion folder NB must not exist as script does not test for exhisting Convert folder
#New-Item -Path $destfolder -ItemType directory

#start ms Paint
Start-Process 'C:\Windows\system32\mspaint.exe'

$wshell = New-Object -ComObject wscript.shell
$wshell.AppActivate('Paint')
Sleep 1

foreach ($srcitem in $(Get-ChildItem $srcfolder -include $src_filter -recurse)) {

    $srcname = $srcitem.fullname

    #create new name
    $partial = $srcitem.FullName.Substring( $srcfolder.Length )
    $destname = $destfolder + $partial
    $destname= [System.IO.Path]::ChangeExtension( $destname , $dest_ext )
    $destpath = [System.IO.Path]::GetDirectoryName( $destname )

    # Create the destination path if it does not exist
    if (-not (test-path $destpath))
    {
        New-Item $destpath -type directory | Out-Null
    }

    #open file in paint and press enter    
    $wshell.SendKeys('^(o)')
    $wshell.SendKeys($srcname)
    $wshell.SendKeys('~')

    #send f12 (save as) to paint window    
    $wshell.SendKeys('{F12}');

    #send file name to paint window and press enter
    $wshell.SendKeys($destname)
    $wshell.SendKeys('~')   
     
}

#close ms paint
$wshell.SendKeys('%{F4}')
