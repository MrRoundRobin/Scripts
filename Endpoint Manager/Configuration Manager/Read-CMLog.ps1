[cmdletbinding()]
param(
    [parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [string]$Path,
    
    [string]$Filter
)

if (-not (Test-Path -Path $Path)) {
    throw 'Log not found'
}

#Read file using StreamReader for big log files

$file = [System.io.File]::Open($path, 'Open', 'Read', 'ReadWrite')
$reader = New-Object System.IO.StreamReader($file)
[string]$LogFileRaw = $reader.ReadToEnd()
$reader.Close()
$file.Close()

$LogFileRaw -split'<!\[LOG\[' | ForEach-Object {
    if ([string]::IsNullOrEmpty($_)) { return }

    $parts = $_ -split '\]LOG\]!>'

    $meta = $parts[1].Trim().Trim('<>').Split(' ').Split('=').Trim('"')

    if (-not [string]::IsNullOrEmpty($Filter) -and $parts[0] -notmatch $Filter) {
        continue
    }

    try {
        $time = [datetime]::ParseExact($meta[$meta.IndexOf('date')+1]+($meta[$meta.IndexOf('time')+1].Split('-')[0]),"MM-dd-yyyyHH:mm:ss.fff",$null)
    } catch [System.Management.Automation.MethodInvocationException] {
        $time = $null
    }

    [PSCustomObject]@{
        LogText   = $parts[0]
        Type      = $meta[$meta.IndexOf('type')+1]
        Component = $meta[$meta.IndexOf('component')+1]
        DateTime  = $time
        Thread    = $meta[$meta.IndexOf('thread')+1]
        Context   = $meta[$meta.IndexOf('context')+1]
        File      = $meta[$meta.IndexOf('file')+1]
        LogFile   = $Path
    }
}
