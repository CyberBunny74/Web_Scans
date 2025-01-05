# Import required module
Import-Module ImportExcel

# Define the subnet to scan
$subnet = "%SUBNET%"  # Replace with your subnet
$ports = %WEB_PORTS% #replace with the ports you wish to check
$timeout = 1000  # Timeout in milliseconds
$logFile = "%ERROR_LOG%" # Replace with your Error Log location and filename
$progressFile = "%PROGRESS_JSON%" # Replace with your Progress file name and location
$resultsFile = "%RESULTS_EXCEL%" # Replace with your results file and location

# Clear the log file if it exists
if (Test-Path $logFile) {
    Remove-Item $logFile
}

# Function to generate IP addresses in the subnet
function Get-IpRange {
    param (
        [string]$Subnet
    )
    $subnetParts = $Subnet.Split('/')
    $ip = [IPAddress]::Parse($subnetParts[0])
    $prefixLength = [int]$subnetParts[1]
    $bytes = $ip.GetAddressBytes()

    [Array]::Reverse($bytes)
    $start = [BitConverter]::ToUInt32($bytes, 0)
    [Array]::Reverse($bytes)

    $maxHosts = [math]::Pow(2, (32 - $prefixLength)) - 1

    $ips = @()
    for ($i = 1; $i -lt $maxHosts; $i++) {
        $bytes = [BitConverter]::GetBytes($start + $i)
        [Array]::Reverse($bytes)
        $ips += [IPAddress]::new($bytes)
    }
    return $ips
}

# Function to check port with timeout
function Test-Port {
    param (
        [string]$IPAddress,
        [int]$Port,
        [int]$Timeout
    )
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $asyncResult = $client.BeginConnect($IPAddress, $Port, $null, $null)
        $asyncResult.AsyncWaitHandle.WaitOne($Timeout) | Out-Null
        if ($client.Connected) {
            $client.EndConnect($asyncResult)
            $client.Close()
            return $true
        } else {
            $client.Close()
            return $false
        }
    } catch {
        Write-Error "Error testing port $Port on ($IPAddress): $($_.Exception.Message)"
        Add-Content -Path $logFile -Value "Error testing port $Port on ($IPAddress): $($_.Exception.Message)"
        return $false
    }
}

# Function to save progress
function Save-Progress {
    param (
        [array]$IpAddresses,
        [int]$CurrentIndex,
        [array]$Results
    )
    $progressData = @{
        IpAddresses = $IpAddresses
        CurrentIndex = $CurrentIndex
        Results = $Results
    }
    $progressData | ConvertTo-Json | Set-Content -Path $progressFile
}

# Function to load progress
function Load-Progress {
    if (Test-Path $progressFile) {
        return (Get-Content -Path $progressFile | ConvertFrom-Json)
    }
    return $null
}

# Load progress if available
$progressData = Load-Progress
if ($progressData -ne $null) {
    $ipAddresses = $progressData.IpAddresses
    $currentIndex = $progressData.CurrentIndex
    $results = $progressData.Results
} else {
    # Generate the list of IP addresses within the subnet
    $ipAddresses = Get-IpRange -Subnet $subnet
    $currentIndex = 0
    $results = @()
}

# Scan IP addresses
for ($i = $currentIndex; $i -lt $ipAddresses.Count; $i++) {
    $ip = $ipAddresses[$i]
    foreach ($port in $ports) {
        try {
            if (Test-Port -IPAddress $ip.IPAddressToString -Port $port -Timeout $timeout) {
                $hostname = ""
                try {
                    $hostname = [System.Net.Dns]::GetHostEntry($ip.IPAddressToString).HostName
                } catch {
                    $hostname = "N/A"
                }
                $result = [PSCustomObject]@{
                    IPAddress   = $ip.IPAddressToString
                    Hostname    = $hostname
                    Port        = $port
                    Status      = "Open"
                }
                $results += $result
                Write-Output $result
            }
        } catch {
            Write-Error "Error scanning $($ip.IPAddressToString) on port ($port): $($_.Exception.Message)"
            Add-Content -Path $logFile -Value "Error scanning $($ip.IPAddressToString) on port ($port): $($_.Exception.Message)"
        }
    }
    Save-Progress -IpAddresses $ipAddresses -CurrentIndex ($i + 1) -Results $results
}

# Display the results
$results | Format-Table -AutoSize

# Export the results to an Excel file
try {
    $results | Export-Excel -Path $resultsFile -AutoSize
    Remove-Item $progressFile -ErrorAction SilentlyContinue # Remove progress file after successful completion
} catch {
    Write-Error "Error exporting results to Excel: $($_.Exception.Message)"
    Add-Content -Path $logFile -Value "Error exporting results to Excel: $($_.Exception.Message)"
}
