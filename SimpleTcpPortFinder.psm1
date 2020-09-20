Function Get-ActiveTcpPortNumber
{
  <#
  .SYNOPSIS
  Retrieves a list of all active TCP Port numbers.

  .DESCRIPTION
  Retrieves a list of all active TCP Port numbers.
  Includes all listener TCP ports, and all active 'client' TCP ports.

  .OUTPUTS
  uint16[]

  .EXAMPLE
  Get-ActiveTcpPortNumber
  #>
  [CmdletBinding()]
  Param()
  Try
  {
    $portHash = New-Object -TypeName Collections.Generic.HashSet[uint16]
    $ipPropertyType = [Net.NetworkInformation.IPGlobalProperties]
    $ipProperties = $ipPropertyType::GetIPGlobalProperties()

    $listenerPorts = $ipProperties.GetActiveTcpListeners()
    $connectedPorts = $ipProperties.GetActiveTcpConnections()

    foreach($serverPort in $listenerPorts)
    {
      [void]$portHash.Add($serverPort.Port)
    }
    foreach ($clientPort in $connectedPorts)
    {
      [void]$portHash.Add($clientPort.LocalEndPoint.Port)
    }
    return $portHash
  }
  Catch
  {
    $emsg = [string]::Format(
      'Unexpected error while obtaining active TCP ports. Message: {0}',
      $_.Exception.Message
    )
    Write-Error -Exception $_.Exception `
      -Message $emsg `
      -Category $_.CategoryInfo.Category
  }
}

Function ConvertTo-RandomArray
{
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory=$true, Position=0)]
    [uint16[]]$Array
  )
  $n = $Array.Length
  $copy = New-Object -TypeName UInt16[] -ArgumentList $n
  [array]::Copy($Array, $copy, $n)
  $rng = New-Object -TypeName Random

  for ($i = 0; $i -lt ($n - 1); $i++)
  {
    $rand = $i + $rng.Next($n - $i)
    $temp = $copy[$rand]
    $copy[$rand] = $copy[$i]
    $copy[$i] = $temp
  }
  return $copy
}

Function Get-InactiveTcpPortNumber
{
  <#
  .SYNOPSIS
  Get inactive / available TCP port numbers.

  .DESCRIPTION
  Get inactive / available TCP port numbers for a provided range.

  .PARAMETER StartRange
  Start of the range of ports to query for.
  Defaults to the IANA suggested ephemeral port range.
  Default: 49152

  .PARAMETER EndRange
  Start of the range of ports to query for.
  Defaults to the IANA suggested ephemeral port range.
  Default: 65535

  .PARAMETER Random
  Switch to retrieve a random port from within the provided range.

  .OUTPUTS
  uint16

  .EXAMPLE
  Get-InactiveTcpPortNumber

  .EXAMPLE
  Get-InactiveTcpPortNumber -Random

  .EXAMPLE
  Get-InactiveTcpPortNumber -StartRange 1024 -EndRange 5000 -Random
  #>
  [CmdletBinding()]
  Param(
    [Parameter(Position=0)]
    [uint16]$StartRange = 49152,

    [Parameter(Position=1)]
    [uint16]$EndRange = 65535,

    [Parameter(Position=2)]
    [switch]$Random
  )
  if ($StartRange -gt $EndRange)
  {
    # Switch values if reversed
    $temp = $StartRange
    $StartRange = $EndRange
    $EndRange = $temp
  }
  Try
  {
    [uint16[]]$portRange = $StartRange..$EndRange
    $activePorts = Get-ActiveTcpPortNumber -ErrorAction Stop

    if ($Random.ToBool())
    {
      $portRange = ConvertTo-RandomArray -Array $portRange
    }
    for ($i = 0; $i -lt $portRange.Count; $i++)
    {
      if ($portRange[$i] -notin $activePorts)
      {
        return $portRange[$i]
      }
    }
    $emsg = [string]::Format(
      'Failed to obtain inactive / available port in range: [{0}]',
      "$StartRange..$EndRange"
    )
    $except = New-Object -TypeName System.InvalidOperationException `
      -ArgumentList $emsg
    Write-Error -Exception $except -Category InvalidOperation -ErrorAction Stop
  }
  Catch
  {
    Write-Error -ErrorRecord $_
  }
}
