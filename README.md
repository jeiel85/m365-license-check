# m365-license-check

Microsoft 365(개인·Family)를 구독 중인데 Word·Excel에 **"정기 가입이 취소되었습니다"** 가 뜰 때, 원인을 진단하고 라이선스 캐시를 초기화하는 Windows용 단일 파일 도구입니다.

**랜딩 페이지: https://jeiel85.github.io/m365-license-check/**

## 실행

PowerShell(또는 Windows 터미널)에 붙여넣고 Enter:

```powershell
irm https://raw.githubusercontent.com/jeiel85/m365-license-check/main/m365-license-check.cmd | iex
```

또는 [`m365-license-check.cmd`](m365-license-check.cmd)를 내려받아 더블클릭하세요. 관리자 권한은 필요 없습니다.

> 브라우저는 보안상 웹페이지에서 PC의 스크립트를 바로 실행할 수 없어서, "복사 후 붙여넣기" 또는 "내려받아 더블클릭"이 가장 짧은 경로입니다.

## 하는 일

| 단계 | 동작 | 변경 여부 |
|---|---|---|
| 진단 | 설치 제품이 구독용인지, 공유 컴퓨터 인증, 로그인 계정 수, 구독 라이선스 상태·만료일 (Office에 포함된 Microsoft 도구 `vnextdiag.ps1` 사용) | 읽기만 |
| 초기화 | Office 앱 종료(강제 종료는 따로 확인) → 라이선스 캐시 백업 → 삭제 → Word 실행 → 재로그인 후 자동 재진단 | **물어본 뒤에만** |

- 지우는 대상: `%LOCALAPPDATA%\Microsoft\Office\Licenses`, `HKCU\Software\Microsoft\Office\16.0\Common\Licensing`
- 백업 위치: `%LOCALAPPDATA%\Microsoft\Office\license-backup-<날짜시간>\` (되돌리는 방법을 화면에 출력)
- 로그인 계정 정보는 지우지 않고, 어떤 정보도 외부로 보내지 않습니다.

## 그래도 "취소"로 나오면

이 PC가 아니라 구독 쪽 문제입니다. Family 구성원이라면 구독 소유자가 [account.microsoft.com/services](https://account.microsoft.com/services)에서 정기 결제 해제·결제 실패·공유 목록 제외 여부를 확인해야 합니다.
OneDrive는 무료 5GB로도 동작하므로, OneDrive가 된다고 Office 라이선스가 정상인 것은 아닙니다.

회사·학교 계정(Microsoft 365 비즈니스/엔터프라이즈)은 대상이 아닙니다. 조직 관리자에게 문의하세요.

## 파일 하나로 두 가지 실행 방식

`m365-license-check.cmd`는 배치 파일이면서 동시에 PowerShell 스크립트입니다. 첫 줄 `<# :`부터 `#>`까지는 배치가 실행하는 런처이고, PowerShell에게는 주석입니다. 그래서 더블클릭(배치)과 `irm | iex`(PowerShell) 모두 같은 파일로 동작합니다.

## 라이선스

MIT. Microsoft와 관련 없는 비공식 도구이며, Microsoft 365·Word·OneDrive는 Microsoft Corporation의 상표입니다.
