<# : -- batch launcher (PowerShell reads this block as a comment) --
@echo off
setlocal
set "M365LC_SELF=%~f0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Expression ([IO.File]::ReadAllText($env:M365LC_SELF, [Text.Encoding]::UTF8))"
echo.
pause
exit /b
#>
#
# m365-license-check
# Microsoft 365(개인/Family) Office 라이선스 진단 및 캐시 초기화 - 파일 하나로 동작
#
# 실행 방법 (둘 중 하나)
#   1) PowerShell에 붙여넣기:
#      irm https://raw.githubusercontent.com/jeiel85/m365-license-check/main/m365-license-check.cmd | iex
#   2) 이 파일을 내려받아 더블클릭
#
# 하는 일
#   - 진단(읽기 전용): 설치 제품, 공유 컴퓨터 인증, 로그인 계정 수, 구독 라이선스 상태
#     (Office에 함께 설치되는 Microsoft 도구 vnextdiag.ps1 사용)
#   - 초기화(묻고 나서만): 라이선스 캐시를 백업한 뒤 지우고 Word를 다시 열어 재로그인으로 새로 받게 함
#     로그인 계정 정보는 건드리지 않으며, 어떤 정보도 외부로 보내지 않음
#
# MIT License. Microsoft와 관련 없는 비공식 도구입니다.

& {
    $ErrorActionPreference = 'Stop'
    $Version      = '1.0.0'
    $OfficeCommon = 'HKCU:\Software\Microsoft\Office\16.0\Common'
    $LicensingKey = "$OfficeCommon\Licensing"
    $OfficeLocal  = Join-Path $env:LOCALAPPDATA 'Microsoft\Office'
    $LicensesDir  = Join-Path $OfficeLocal 'Licenses'
    $OfficeProcs  = 'WINWORD','EXCEL','POWERPNT','OUTLOOK','ONENOTE','MSACCESS','MSPUB','VISIO','WINPROJ'

    function Write-Section([string]$Title) {
        Write-Host ''
        Write-Host "== $Title ==" -ForegroundColor Cyan
    }

    function Confirm-Yes([string]$Prompt) {
        $answer = Read-Host "$Prompt [y/N]"
        return ($answer -match '^\s*[yY]')
    }

    function Get-OfficeInstall {
        $cfg = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
        if (-not $cfg) { return $null }
        [pscustomobject]@{
            Products      = $cfg.ProductReleaseIds
            Version       = $cfg.VersionToReport
            Platform      = $cfg.Platform
            SharedLicense = $cfg.SharedComputerLicensing
            InstallPath   = $cfg.InstallationPath
        }
    }

    function Get-SignedInAccounts {
        $root = "$OfficeCommon\Identity\Identities"
        if (-not (Test-Path $root)) { return @() }
        @(Get-ChildItem $root | ForEach-Object {
            $p = Get-ItemProperty $_.PSPath
            if ($p.EmailAddress) { [pscustomobject]@{ Email = $p.EmailAddress; Provider = $p.ProviderId } }
        })
    }

    function Find-VNextDiag([string]$InstallPath) {
        $candidates = @()
        if ($InstallPath) { $candidates += (Join-Path $InstallPath 'root\Office16\vnextdiag.ps1') }
        $candidates += 'C:\Program Files\Microsoft Office\root\Office16\vnextdiag.ps1'
        $candidates += 'C:\Program Files (x86)\Microsoft Office\root\Office16\vnextdiag.ps1'
        foreach ($c in $candidates) { if (Test-Path $c) { return $c } }
        return $null
    }

    # vnextdiag.ps1 -action list 출력의 "vNext licenses found" 구역에서 JSON 블록을 뽑는다.
    function Get-VNextLicenses([string]$DiagPath) {
        $raw = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $DiagPath -action list 2>&1 | Out-String
        $section = $raw
        $start = $raw.IndexOf('vNext licenses found')
        if ($start -ge 0) {
            $section = $raw.Substring($start)
            $end = $section.IndexOf('Device licenses found')
            if ($end -ge 0) { $section = $section.Substring(0, $end) }
        }
        $licenses = @()
        foreach ($m in [regex]::Matches($section, '(?s)\{.*?\}')) {
            try { $licenses += ($m.Value | ConvertFrom-Json) } catch { }
        }
        return ,$licenses
    }

    function Get-RunningOffice {
        @(Get-Process -Name $OfficeProcs -ErrorAction SilentlyContinue)
    }

    # 반환값: 'ok' | 'needs-reset' | 'other'
    function Invoke-Diagnose {
        $findings = New-Object System.Collections.Generic.List[string]
        $needsReset = $false

        Write-Section 'Office 설치'
        $install = Get-OfficeInstall
        if (-not $install) {
            Write-Host 'Microsoft 365 앱(Click-to-Run) 설치를 찾지 못했습니다.' -ForegroundColor Yellow
            Write-Host 'account.microsoft.com/services 에서 Office를 설치한 뒤 다시 실행하세요.'
            return 'other'
        }
        $install | Format-List | Out-Host
        if ($install.Products -notmatch 'O365') {
            $findings.Add("설치된 제품($($install.Products))이 Microsoft 365 구독용이 아닙니다(영구 라이선스 등). 구독 계정으로 로그인해도 이 설치본은 구독 라이선스를 쓰지 않으니 account.microsoft.com/services 에서 구독용으로 다시 설치하세요.")
        }
        if ($install.SharedLicense -eq '1') {
            $findings.Add('공유 컴퓨터 정품 인증(Shared Computer Licensing)이 켜져 있습니다. 개인/Family 구독에서는 정상 동작하지 않을 수 있습니다.')
        }

        Write-Section 'Office 로그인 계정'
        $accounts = Get-SignedInAccounts
        if ($accounts.Count -eq 0) {
            Write-Host '(레지스트리에서 확인 불가 - Word의 파일 > 계정 화면에서 확인하세요)'
        } else {
            $accounts | Format-Table -AutoSize | Out-Host
            if ($accounts.Count -gt 1) {
                $findings.Add("Office에 계정이 $($accounts.Count)개 로그인돼 있습니다. 구독이 없거나 취소된 계정의 라이선스를 잡았을 수 있으니 구독 계정만 남기고 로그아웃해 보세요.")
            }
        }

        Write-Section '구독 라이선스'
        $diag = Find-VNextDiag $install.InstallPath
        if ($diag) {
            $licenses = Get-VNextLicenses $diag
            if ($licenses.Count -eq 0) {
                Write-Host '이 PC에 받아 둔 구독 라이선스가 없습니다.' -ForegroundColor Yellow
                $findings.Add('구독 라이선스를 아직 받지 않았습니다. Word를 열고 파일 > 계정에서 구독 계정으로 로그인하세요.')
            } else {
                $licenses | Select-Object Product, Type, LicenseState, EntitlementStatus, ReasonCode, NotBefore, NotAfter |
                    Format-List | Out-Host
                $now = (Get-Date).ToUniversalTime()
                foreach ($l in $licenses) {
                    $bad = @()
                    if ($l.LicenseState -and $l.LicenseState -ne 'Licensed') { $bad += "LicenseState=$($l.LicenseState)" }
                    if ($l.EntitlementStatus -and $l.EntitlementStatus -ne 'Provisioned') { $bad += "EntitlementStatus=$($l.EntitlementStatus)" }
                    if ($l.ReasonCode) { $bad += "ReasonCode=$($l.ReasonCode)" }
                    if ($l.NotAfter) {
                        try { if (([datetime]$l.NotAfter).ToUniversalTime() -lt $now) { $bad += "NotAfter 지남($($l.NotAfter))" } } catch { }
                    }
                    if ($bad.Count -gt 0) {
                        $needsReset = $true
                        $findings.Add("$($l.Product) 라이선스 이상: $($bad -join ', '). 계정 페이지(account.microsoft.com/services)에서 구독이 살아 있다면 이 PC의 라이선스 캐시가 낡은 것입니다.")
                    }
                }
            }
        } else {
            Write-Host 'Office 진단 도구(vnextdiag.ps1)를 찾지 못해 라이선스 파일 목록만 표시합니다.' -ForegroundColor Yellow
            Get-ChildItem $LicensesDir -Recurse -File -ErrorAction SilentlyContinue |
                Select-Object FullName, LastWriteTime | Format-Table -AutoSize | Out-Host
        }

        Write-Section '판정'
        if ($findings.Count -eq 0) {
            Write-Host '이상 없음: 구독 라이선스가 정상(Licensed / Provisioned)입니다.' -ForegroundColor Green
            return 'ok'
        }
        foreach ($f in $findings) { Write-Host "- $f" -ForegroundColor Yellow }
        if ($needsReset) { return 'needs-reset' }
        return 'other'
    }

    function Close-OfficeApps {
        $procs = Get-RunningOffice
        if ($procs.Count -eq 0) { return $true }

        Write-Host 'Office 앱을 닫습니다:'
        $procs | Select-Object Name, Id, MainWindowTitle | Format-Table -AutoSize | Out-Host
        foreach ($p in $procs) { try { [void]$p.CloseMainWindow() } catch { } }
        Start-Sleep -Seconds 5

        $procs = Get-RunningOffice
        if ($procs.Count -eq 0) { return $true }

        Write-Host '정상 종료되지 않은 Office 앱이 있습니다(라이선스 안내 창이 떠 있으면 흔히 이렇습니다).' -ForegroundColor Yellow
        $procs | Select-Object Name, Id, MainWindowTitle | Format-Table -AutoSize | Out-Host
        if (-not (Confirm-Yes '강제 종료할까요? 저장하지 않은 편집 내용은 사라집니다.')) { return $false }
        $procs | Stop-Process -Force
        Start-Sleep -Seconds 2
        return ((Get-RunningOffice).Count -eq 0)
    }

    # 반환값: 성공 시 백업 경로, 중단 시 $null
    function Invoke-Reset {
        Write-Section '라이선스 캐시 초기화'
        if (-not (Close-OfficeApps)) {
            Write-Host 'Office 앱이 열려 있어 중단했습니다. 모두 닫은 뒤 다시 실행하세요.' -ForegroundColor Red
            return $null
        }

        $backup = Join-Path $OfficeLocal ('license-backup-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
        New-Item -ItemType Directory -Path $backup -Force | Out-Null
        if (Test-Path $LicensesDir) { Copy-Item $LicensesDir (Join-Path $backup 'Licenses') -Recurse }
        if (Test-Path $LicensingKey) {
            & reg.exe export 'HKCU\Software\Microsoft\Office\16.0\Common\Licensing' (Join-Path $backup 'Licensing.reg') /y | Out-Null
            if ($LASTEXITCODE -ne 0) { throw "레지스트리 백업 실패(reg export 종료 코드 $LASTEXITCODE). 아무것도 지우지 않고 중단합니다." }
        }
        Write-Host "백업 완료: $backup" -ForegroundColor Green

        if (Test-Path $LicensesDir)  { Remove-Item $LicensesDir -Recurse -Force }
        if (Test-Path $LicensingKey) { Remove-Item $LicensingKey -Recurse -Force }
        if ((Test-Path $LicensesDir) -or (Test-Path $LicensingKey)) { throw '캐시 삭제가 끝나지 않았습니다. 위 경로를 확인하세요.' }
        Write-Host '라이선스 캐시를 지웠습니다.' -ForegroundColor Green
        return $backup
    }

    try {
        Write-Host ''
        Write-Host "m365-license-check v$Version  -  Microsoft 365 Office 라이선스 진단" -ForegroundColor White
        Write-Host '어떤 정보도 외부로 보내지 않습니다. 진단은 읽기만 합니다.'

        if ($env:OS -ne 'Windows_NT') { Write-Host 'Windows에서만 동작합니다.' -ForegroundColor Red; return }

        $state = Invoke-Diagnose
        if ($state -eq 'other' -and -not (Get-OfficeInstall)) { return }

        Write-Host ''
        if ($state -eq 'needs-reset') {
            $doReset = Confirm-Yes '라이선스 캐시를 백업 후 초기화하고 새로 받을까요?'
        } else {
            Write-Host 'Word에 여전히 "정기 가입이 취소되었습니다"가 뜬다면 캐시 초기화로 새 라이선스를 받을 수 있습니다.'
            $doReset = Confirm-Yes '라이선스 캐시를 백업 후 초기화할까요?'
        }
        if (-not $doReset) {
            Write-Host '초기화하지 않고 마칩니다.'
        } else {
            $backup = Invoke-Reset
            if ($backup) {
                Start-Process winword.exe
                Write-Host ''
                Write-Host '방금 열린 Word에서 구독 계정으로 로그인하세요.' -ForegroundColor Cyan
                Write-Host '(로그인 창이 없으면 파일 > 계정 > 로그인 또는 "라이선스 업데이트")'
                [void](Read-Host '로그인을 마쳤으면 Enter를 누르세요. 다시 진단합니다')
                $after = Invoke-Diagnose
                if ($after -ne 'ok') {
                    Write-Host ''
                    Write-Host '로그인 직후에는 반영이 늦을 수 있으니 Word를 닫았다 열고 한 번 더 실행해 보세요.'
                }
                Write-Host ''
                Write-Host "되돌리기: Office를 닫고 $backup\Licenses 폴더를 $OfficeLocal 로 복사, Licensing.reg 를 더블클릭해 가져오기"
            }
        }

        Write-Host ''
        Write-Host '초기화 후에도 "취소"로 나오면 이 PC 문제가 아닙니다.' -ForegroundColor White
        Write-Host '  - Family 구성원: 구독 소유자가 account.microsoft.com/services 에서'
        Write-Host '    정기 결제 해제 / 결제 실패 / 공유 목록 제외 여부를 확인해야 합니다.'
        Write-Host '  - OneDrive는 무료 5GB로도 동작하므로 OneDrive가 된다고 Office 라이선스가 정상인 것은 아닙니다.'
    } catch {
        Write-Host ''
        Write-Host "오류: $($_.Exception.Message)" -ForegroundColor Red
    }
}
