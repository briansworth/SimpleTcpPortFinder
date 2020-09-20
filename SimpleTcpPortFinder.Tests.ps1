$modName = 'SimpleTcpPortFinder'
Get-Module -Name $modName -ErrorAction SilentlyContinue | Remove-Module
Import-Module -Name $modName

InModuleScope -ModuleName $modName {
  Describe 'ConvertTo-RandomArray' {
    BeforeAll {
      [uint16[]]$uintArr = @(1..20)
    }
    It 'Returns uint16 array' {
      $result = ConvertTo-RandomArray -Array $uintArr

      $result | Should -HaveCount $uintArr.Length
      $result | Should -BeOfType [uint16]
    }

    It 'Returns randomized array' {
      $result = ConvertTo-RandomArray -Array $uintArr

      $success = $false
      for ($i = 0; $i -lt $result.Count; $i++)
      {
        if ($result[$i] -ne $uintArr[$i])
        {
          $success = $true
          break
        }
      }

      $success | Should -Be $true
    }
  }
}

Describe 'Get-ActiveTcpPortNumber' {
  It 'Returns uint16 array' {
    $result = Get-ActiveTcpPortNumber

    $result | Should -BeOfType [uint16]
    $result.Length | Should -BeGreaterThan 1
  }

  It 'Throws if unexpected error occurs' {
    $emsg = 'Unexpected'
    $exceptionType = [Microsoft.PowerShell.Commands.WriteErrorException]
    Mock -CommandName New-Object -MockWith {
      Write-Error -Message $emsg -ErrorAction Stop
    }

    {Get-ActiveTcpPortNumber -ErrorAction Stop} |
      Should -Throw $emsg -ExceptionType $exceptionType
  }
}

Describe 'Get-InactiveTcpPortNumber' {
  BeforeAll {
    $range = @(22..25)
  }
  BeforeEach {
    Mock -CommandName Get-ActiveTcpPortNumber -MockWith {
      return $range
    }
  }

  It 'Returns port number' {
    $result = Get-InactiveTcpPortNumber -StartRange 1 -EndRange 2

    $result | Should -BeExactly 1
  }

  It 'Returns port number with -Random' {
    $result = Get-InactiveTcpPortNumber -StartRange 1 -EndRange 2 -Random

    $result | Should -BeGreaterOrEqual 1
  }

  It 'Runs successfully with StartRange > EndRange' {
    $result = Get-InactiveTcpPortNumber -StartRange 2 -EndRange 1

    $result | Should -BeExactly 1
  }

  It 'Throws when unable to find inactive port' {
    $s = $range[0]
    $e = $range[-1]

    {Get-InactiveTcpPortNumber -StartRange $s -EndRange $e -ErrorAction Stop} |
      Should -Throw -ExceptionType InvalidOperationException
  }
}
