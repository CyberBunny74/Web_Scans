# Import the required module
Import-Module ImportExcel

$computers = Import-Csv -Path "%INPUT_CSV%" # replace with the csv file you are using
$ports = 80,443
$results = @()

ForEach ($computer in $computers) {
    ForEach ($port in $ports) {
        $status = if (Test-NetConnection -ComputerName $computer.ComputerName -Port $port -InformationLevel Quiet -WarningAction SilentlyContinue) {
            "Open"
        } else {
            "Closed"
        }

        $results += [PSCustomObject]@{
            ComputerName = $computer.ComputerName
            Port         = $port
            Status       = $status
        }
    }
}

$results | Export-Excel -Path "%OUTPUT_EXCEL%" -AutoSize
