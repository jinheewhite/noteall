<#
  파워포인트(.ppt/.pptx) 또는 한글(.hwp/.hwpx) 파일을, 그 PC에 실제 설치된
  프로그램을 COM 자동화로 띄워 같은 내용의 PDF로 저장합니다.
  만들어진 PDF는 note가 이미 가진 "PDF 열기" 경로로 그대로 들어갑니다.
#>
param(
  [Parameter(Mandatory=$true)][ValidateSet('ppt','hwp')][string]$Kind,
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
} else {
  $hwp = New-Object -ComObject HWPFrame.HwpObject
  try {
    # 보안 모듈 등록 — 없으면 자동화 중 "보안 경고" 창이 뜰 수 있습니다.
    try { $hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample') } catch {}
    try { $hwp.XHwpWindows.Item(0).Visible = $false } catch {}
    $opened = $hwp.Open($InputPath, 'HWP', 'forceopen:true')
    if (-not $opened) { throw '한글 파일을 열지 못했습니다.' }
    $hwp.SaveAs($OutputPath, 'PDF', '')
  } finally {
    $hwp.Quit()
  }
}

if (-not (Test-Path -LiteralPath $OutputPath)) {
  throw 'PDF 파일이 만들어지지 않았습니다.'
}
