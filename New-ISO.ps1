function New-ISO {
    <#
    .SYNOPSIS
        Create an ISO file from files or folders.

    .DESCRIPTION
        Create an ISO file from selected files or folders.
        Uses a GUI interface for file/folder selection if no source is specified.

    .PARAMETER source
        The source files/folder to add to the ISO. If not specified, a file selection dialog will open.

    .PARAMETER destination
        The ISO file to create. If not specified, a file save dialog will open.

    .PARAMETER title
        Optional. Title of the ISO file. Defaults to the filename provided in -destination.

    .PARAMETER force
        Optional. Force overwrite of an existing ISO file.

    .EXAMPLE
        New-ISO

        Opens file/folder selection dialog and save dialog to create an ISO interactively.

    .EXAMPLE
        New-ISO -source C:\MyFiles -destination C:\Output\archive.iso

        Creates archive.iso containing the contents of C:\MyFiles.
    #>

    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="Low")]
    Param
    (
        [parameter(Mandatory=$false,ValueFromPipeline=$false)]
        [string]$source,
        [parameter(Mandatory=$false,ValueFromPipeline=$false)]
        [string]$destination,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false)]
        [string]$title,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false)]
        [switch]$force
    )

    begin {
        Add-Type -AssemblyName System.Windows.Forms
        $script:proceedWithISO = $true  # Flag to control if we should proceed with ISO creation

        # If -source is not provided, open the selection dialog
        if (-not $PSBoundParameters.ContainsKey('source')) {
            $script:selectedPaths = Select-Files
            if ($null -eq $selectedPaths) {
                Write-Host "No source selected, exiting." -ForegroundColor Green
                $script:proceedWithISO = $false
                return
            }
            $source = "GUI_SELECTION"  # Special flag to indicate GUI selection
        }

        # If -destination is not provided, open the file save dialog
        if ($proceedWithISO -and -not $PSBoundParameters.ContainsKey('destination')) {
            $destinationFile = Select-Destination
            if ($null -eq $destinationFile) {
                Write-Host "No destination selected, exiting." -ForegroundColor Green
                $script:proceedWithISO = $false
                return
            }
            $destination = $destinationFile
        }

        # If -title is not provided, generate it from the destination filename
        if ($proceedWithISO -and -not $PSBoundParameters.ContainsKey('title')) {
            $title = [System.IO.Path]::GetFileNameWithoutExtension($destination)
        }

        if ($proceedWithISO) {
            Write-Verbose ("Function start.")
        }
    }

    process {
        if (-not $proceedWithISO) {
            return
        }
        
        ## Set type definition for ISO creation
        Write-Verbose ("Adding ISOFile type.")

        $typeDefinition = @'
        public class ISOFile  {
            public unsafe static void Create(string Path, object Stream, int BlockSize, int TotalBlocks) {
                int bytes = 0;
                byte[] buf = new byte[BlockSize];
                var ptr = (System.IntPtr)(&bytes);
                var o = System.IO.File.OpenWrite(Path);
                var i = Stream as System.Runtime.InteropServices.ComTypes.IStream;

                if (o != null) {
                    while (TotalBlocks-- > 0) {
                        i.Read(buf, BlockSize, ptr); o.Write(buf, 0, bytes);
                    }
                    o.Flush(); o.Close();
                }
            }
        }
'@

        ## Create type ISOFile if not already created
        if (!('ISOFile' -as [type])) {
            switch ($PSVersionTable.PSVersion.Major) {
                ## PowerShell 7 and later
                {$_ -ge 7} {
                    Write-Verbose ("Adding type for PowerShell 7 or later.")
                    Add-Type -TypeDefinition $typeDefinition -CompilerOptions "/unsafe"
                }
                ## PowerShell 5
                5 {
                    Write-Verbose ("Adding type for PowerShell 5.")
                    $compOpts = New-Object System.CodeDom.Compiler.CompilerParameters
                    $compOpts.CompilerOptions = "/unsafe"
                    Add-Type -TypeDefinition $typeDefinition -CompilerParameters $compOpts
                }
                default {
                    throw ("Unsupported PowerShell version.")
                }
            }
        }

        ## Initialize image object
        Write-Verbose ("Initializing image object.")
        try {
            $image = New-Object -ComObject IMAPI2FS.MsftFileSystemImage -Property @{VolumeName=$title} -ErrorAction Stop
            Write-Verbose ("Initialized.")
        }
        catch {
            throw ("Failed to initialize image. " + $_.exception.Message)
        }

        ## Create target ISO file
        if ($PSCmdlet.ShouldProcess($destination)) {
            if (!($targetFile = New-Item -Path $destination -ItemType File -Force:$Force -ErrorAction SilentlyContinue)) {
                throw ("Cannot create file " + $destination + ". Use -Force parameter to overwrite if the target file already exists.")
            }
        }

        ## Get source content
        Write-Verbose ("Fetching source items.")
        try {
            $sourceItems = @()
            
            if ($source -eq "GUI_SELECTION") {
                # Handle GUI-selected paths
                foreach ($path in $selectedPaths) {
                    if (Test-Path -Path $path -PathType Container) {
                        # If it's a directory, add the folder itself
                        $sourceItems += Get-Item -LiteralPath $path -ErrorAction Stop
                    } else {
                        # If it's a file, get the item directly
                        $sourceItems += Get-Item -LiteralPath $path -ErrorAction Stop
                    }
                }
            } else {
                # Handle command-line specified path
                if (Test-Path -Path $source -PathType Container) {
                    # Add the folder itself
                    $sourceItems += Get-Item -LiteralPath $source -ErrorAction Stop
                } else {
                    $sourceItems += Get-Item -LiteralPath $source -ErrorAction Stop
                }
            }

            if ($sourceItems.Count -eq 0) {
                throw "No source items found to add to the ISO."
            }

            Write-Verbose ("Got source items.")
        }
        catch {
            throw ("Failed to get source items. " + $_.exception.message)
        }

        ## Add items to image
        Write-Verbose ("Adding items to image.")
        foreach($sourceItem in $sourceItems) {
            try {
                if ($sourceItem.PSIsContainer) {
                    # For folders, preserve the entire directory structure
                    Write-Verbose ("Adding folder and contents: " + $sourceItem.FullName)
                    $image.Root.AddTree($sourceItem.FullName, $true)
                } else {
                    # For individual files, add directly to root
                    Write-Verbose ("Adding file: " + $sourceItem.FullName)
                    $image.Root.AddTree($sourceItem.FullName, $false)
                }
            }
            catch {
                throw ("Failed to add " + $sourceItem.fullname + ". " + $_.exception.message)
            }
        }

        ## Create the ISO file
        Write-Verbose ("Writing ISO file to " + $targetFile)
        try {
            $result = $image.CreateResultImage()
            [ISOFile]::Create($targetFile.FullName, $result.ImageStream, $result.BlockSize, $result.TotalBlocks)
        }
        catch {
            throw ("Failed to write ISO file. " + $_.exception.Message)
        }

        Write-Verbose ("File complete.")
        return $targetFile
    }

    end {
        Write-Verbose ("Function complete.")
    }
}

# Function to open dialog to select files or folder
function Select-Files {
    # Create buttons for selection type
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Files or Folder"
    $form.Size = New-Object System.Drawing.Size(300,150)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $filesButton = New-Object System.Windows.Forms.Button
    $filesButton.Location = New-Object System.Drawing.Point(75,20)
    $filesButton.Size = New-Object System.Drawing.Size(150,30)
    $filesButton.Text = "Select Files"
    
    $folderButton = New-Object System.Windows.Forms.Button
    $folderButton.Location = New-Object System.Drawing.Point(75,60)
    $folderButton.Size = New-Object System.Drawing.Size(150,30)
    $folderButton.Text = "Select Folder"

    $script:result = $null  # Use script scope

    # Files button click handler
    $filesButton.Add_Click({
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.InitialDirectory = [System.Environment]::GetFolderPath('Desktop')
        $fileDialog.Filter = "All files (*.*)|*.*"
        $fileDialog.Multiselect = $true
        $fileDialog.Title = "Select Files to Include in ISO"

        if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:result = $fileDialog.FileNames
            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        }
    })

    # Folder button click handler
    $folderButton.Add_Click({
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "Select Folder to Include in ISO"
        $folderDialog.ShowNewFolderButton = $true
        
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:result = $folderDialog.SelectedPath
            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        }
    })

    # Add buttons to form
    $form.Controls.Add($filesButton)
    $form.Controls.Add($folderButton)

    # Show form as dialog
    $form.ShowDialog() | Out-Null

    return $script:result
}

# Function to open file dialog to select destination for saving the ISO
function Select-Destination {
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.InitialDirectory = [System.Environment]::GetFolderPath('Desktop')
    $saveFileDialog.Filter = "ISO files (*.iso)|*.iso"
    $saveFileDialog.DefaultExt = "iso"
    $saveFileDialog.AddExtension = $true
    $saveFileDialog.Title = "Select Destination to Save ISO"

    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $saveFileDialog.FileName
    } else {
        return $null
    }
}
