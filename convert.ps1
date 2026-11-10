<#
  파워포인트(.ppt/.pptx), 한글(.hwp/.hwpx), 엑셀(.xls/.xlsx) 파일을, 그 PC에
  실제 설치된 프로그램을 COM 자동화로 띄워 같은 내용의 PDF로 저장합니다.
  만들어진 PDF는 note가 이미 가진 "PDF 열기" 경로로 그대로 들어갑니다.
#>
param(
  [Parameter(Mandatory=$true)][ValidateSet('ppt','hwp','excel')][string]$Kind,
  [Parameter(Mandatory=$true)][string]$InputPath,
  [Parameter(Mandatory=$true)][string]$OutputPath
)
$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $InputPath)) {
  throw "원본 파일을 찾을 수 없습니다: $InputPath"
}

if ($Kind -eq 'ppt') {
  $app = New-Object -ComObject PowerPoint.Application
  try {
    # Open(FileName, ReadOnly, Untitled, WithWindow) — 창을 띄우지 않습니다.
    $pres = $app.Presentations.Open($InputPath, $true, $true, $false)
    try {
      $pres.SaveAs($OutputPath, 32)   # ppSaveAsPDF
    } finally {
      $pres.Close()
    }
  } finally {
    $app.Quit()
  }
} elseif ($Kind -eq 'hwp') {
  $hwp = New-Object -ComObject HWPFrame.HwpObject
  try {
    # 보안 모듈 등록 — 없으면 자동화 중 "보안 경고" 창이 뜰 수 있습니다.
    try { $hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample') } catch {}
    try { $hwp.XHwpWindows.Item(0).Visible = $false } catch {}
    # 형식 확인창(호환 형식 변환 등)이 숨은 채로 뜨면 자동화가 영영 멈추므로,
    # 전부 "예/확인"으로 자동 응답하게 해 둡니다.
    try { $hwp.SetMessageBoxMode(0x00000020) } catch {}
    # .hwpx(새 형식)와 .hwp(옛 형식)는 Open 의 두 번째 인자(형식)가 다릅니다 —
    # 확장자로 맞는 쪽을 골라야, hwpx를 HWP로 잘못 열다 멈추는 일이 없습니다.
    $fmt = if ($InputPath -match '\.hwpx$') { 'HWPX' } else { 'HWP' }
    $opened = $hwp.Open($InputPath, $fmt, 'forceopen:true')
    if (-not $opened) { throw '한글 파일을 열지 못했습니다.' }
    $hwp.SaveAs($OutputPath, 'PDF', '')
  } finally {
    $hwp.Quit()
  }
} else {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  try {
    $wb = $excel.Workbooks.Open($InputPath, 0, $true)
    try {
      $wb.ExportAsFixedFormat(0, $OutputPath)   # 0 = xlTypePDF
    } finally {
      $wb.Close($false)
    }
  } finally {
    $excel.Quit()
  }
}

if (-not (Test-Path -LiteralPath $OutputPath)) {
  throw 'PDF 파일이 만들어지지 않았습니다.'
}
