Function Set-Window {
    <#
        .SYNOPSIS
            Retrieve/Set the window size and coordinates of a process window.

        .DESCRIPTION
            Retrieve/Set the size (height,width) and coordinates (x,y) of a process window.

        .PARAMETER ProcessName
            Name of the process to determine the window characteristics. 
            (All processes if omitted).

        .PARAMETER Id
            Id of the process to determine the window characteristics. 

        .PARAMETER X
            Set the position of the window in pixels from the left.

        .PARAMETER Y
            Set the position of the window in pixels from the top.

        .PARAMETER Width
            Set the width of the window.

        .PARAMETER Height
            Set the height of the window.

        .PARAMETER Passthrough
            Returns the output object of the window.

        .NOTES
            Name:       Set-Window
            Authors:    Boe Prox, JosefZ, CanRanBan
            Version History:
                1.0//Boe Prox  - 11/24/2015 - Initial build
                1.1//JosefZ    - 19.05.2018 - Treats more process instances of supplied process name properly
                1.2//JosefZ    - 21.02.2019 - Parameter Id
                1.3//CanRanBan - 2023-01-12 - Formatting changes, removal of unused DLL import, addition of sources, additional documentation, meaningful warnings

        .LINK
            Current Source:
            https://github.com/CanRanBan/Collection-PowerShell-Scripts/blob/main/Set-Window.psm1

        .LINK
            Original Source:
            https://superuser.com/questions/1324007/setting-window-size-and-position-in-powershell-5-and-6

        .OUTPUTS
            None                                            Default is no output.
            System.Management.Automation.PSCustomObject     Returns the window properties if -Passthrough is used.

        .EXAMPLE
            Get-Process powershell | Set-Window -X 20 -Y 40 -Passthrough -Verbose
            VERBOSE: powershell (Id=11140, Handle=132410)

            Id          : 11140
            ProcessName : powershell
            Size        : 1134,781
            TopLeft     : 20,40
            BottomRight : 1154,821

            Description: Set the coordinates on the window for the process PowerShell.exe

        .EXAMPLE
            $windowArray = Set-Window -Passthrough
            WARNING: cmd (1096) is minimized! Coordinates will not be accurate.

            PS > $windowArray | Format-Table -AutoSize

              Id ProcessName    Size     TopLeft       BottomRight  
              -- -----------    ----     -------       -----------  
            1096 cmd            199,34   -32000,-32000 -31801,-31966
            4088 explorer       1280,50  0,974         1280,1024    
            6880 powershell     1280,974 0,0           1280,974     

            Description: Get the coordinates of all visible windows and save them into the $windowArray variable. Then, display them in a table view.

        .EXAMPLE
            Set-Window -Id $PID -Passthrough | Format-Table
​‌‍
              Id ProcessName Size     TopLeft BottomRight
              -- ----------- ----     ------- -----------
            7840 pwsh        1024,638 0,0     1024,638

            Description: Display the coordinates of the window for the current PowerShell session in a table view.

    #>
    [CmdletBinding(DefaultParameterSetName='Name')]
    Param (
        [Parameter(
            Mandatory=$false,
            ValueFromPipelineByPropertyName=$true,
            ParameterSetName='Name')]
        [string]$ProcessName='*',
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$false,
            ParameterSetName='Id')]
        [int]$Id,
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height,
        [switch]$Passthrough
    )
    Begin {
        Try { 
            [void][Window]
        } Catch {
        Add-Type @"
            using System;
            using System.Runtime.InteropServices;

            public class Window {
                [DllImport("user32.dll")]
                [return: MarshalAs(UnmanagedType.Bool)]
                public static extern bool GetWindowRect(
                    IntPtr hWnd,
                    out RECT lpRect
                );

                [DllImport("user32.dll")]
                [return: MarshalAs(UnmanagedType.Bool)]
                public static extern bool MoveWindow(
                    IntPtr handle,
                    int x,
                    int y,
                    int width,
                    int height,
                    bool redraw
                );
            }

            public struct RECT {
                public int Left;    // x position of upper-left corner
                public int Top;     // y position of upper-left corner
                public int Right;   // x position of lower-right corner
                public int Bottom;  // y position of lower-right corner
            }
"@
        }
    }
    Process {
        $Rectangle = New-Object RECT
        if ( $PSBoundParameters.ContainsKey('Id') ) {
            $Processes = Get-Process -Id $Id -ErrorAction SilentlyContinue
        } else {
            $Processes = Get-Process -Name "$ProcessName" -ErrorAction SilentlyContinue
        }
        if ( $null -eq $Processes ) {
            if ( $PSBoundParameters['Passthrough'] ) {
                Write-Warning 'No process found for used parameters.'
            }
        } else {
            $Processes | ForEach-Object {
                $Handle = $_.MainWindowHandle
                Write-Verbose "$($_.ProcessName) `(Id=$($_.Id), Handle=$Handle`)"
                if ( $Handle -eq [System.IntPtr]::Zero ) {
                    return
                }
                $Return = [Window]::GetWindowRect($Handle, [ref]$Rectangle)
                if (-NOT $PSBoundParameters.ContainsKey('X')) {
                    $X = $Rectangle.Left
                }
                if (-NOT $PSBoundParameters.ContainsKey('Y')) {
                    $Y = $Rectangle.Top
                }
                if (-NOT $PSBoundParameters.ContainsKey('Width')) {
                    $Width = $Rectangle.Right - $Rectangle.Left
                }
                if (-NOT $PSBoundParameters.ContainsKey('Height')) {
                    $Height = $Rectangle.Bottom - $Rectangle.Top
                }
                if ( $Return ) {
                    $Return = [Window]::MoveWindow($Handle, $X, $Y, $Width, $Height, $true)
                }
                if ( $PSBoundParameters['Passthrough'] ) {
                    $Rectangle = New-Object RECT
                    $Return = [Window]::GetWindowRect($Handle, [ref]$Rectangle)
                    if ( $Return ) {
                        $Width       = $Rectangle.Right  - $Rectangle.Left
                        $Height      = $Rectangle.Bottom - $Rectangle.Top
                        $Size        = New-Object System.Management.Automation.Host.Size        -ArgumentList $Width, $Height
                        $TopLeft     = New-Object System.Management.Automation.Host.Coordinates -ArgumentList $Rectangle.Left , $Rectangle.Top
                        $BottomRight = New-Object System.Management.Automation.Host.Coordinates -ArgumentList $Rectangle.Right, $Rectangle.Bottom
                        if ($Rectangle.Top    -lt 0 -AND
                            $Rectangle.Bottom -lt 0 -AND
                            $Rectangle.Left   -lt 0 -AND
                            $Rectangle.Right  -lt 0) {
                                Write-Warning "$($_.ProcessName) `($($_.Id)`) is minimized! Coordinates will not be accurate."
                        }
                        $Object = [PSCustomObject]@{
                            Id          = $_.Id
                            ProcessName = $_.ProcessName
                            Size        = $Size
                            TopLeft     = $TopLeft
                            BottomRight = $BottomRight
                        }
                        $Object
                    }
                }
            }
        }
    }
}