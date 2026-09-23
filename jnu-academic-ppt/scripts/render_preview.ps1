<#
  把 .pptx 每页导出成 PNG，供肉眼复核。需要本机装有 PowerPoint。

  用法：
    powershell -ExecutionPolicy Bypass -File render_preview.ps1 -Path "汇报.pptx" -Out ".jnuppt/preview"
    powershell -ExecutionPolicy Bypass -File render_preview.ps1 -Path "汇报.pptx" -Out "prev" -Contact

  -Contact 会额外拼一张九宫格总览图（需要 Python + Pillow）。
#>
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [string]$Out = "preview",
  [int]$W = 1280,
  [int]$H = 720,
  [switch]$Contact,
  [string]$Python = ""
)
$ErrorActionPreference = "Stop"

$Path = (Resolve-Path $Path).Path
if (-not (Test-Path $Out)) { New-Item -ItemType Directory -Force -Path $Out | Out-Null }
$Out = (Resolve-Path $Out).Path

$app = $null
$pres = $null
# PowerPoint 是单实例：用户已经开着时 New-Object 拿到的就是用户那个进程。
# 只有本脚本自己拉起的实例才 Quit，否则会把用户正在编辑的文件一起关掉。
$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
try {
  $app = New-Object -ComObject PowerPoint.Application
  $pres = $app.Presentations.Open($Path, $true, $false, $false)   # ReadOnly, 不开窗口
  $n = $pres.Slides.Count
  for ($i = 1; $i -le $n; $i++) {
    $pres.Slides.Item($i).Export((Join-Path $Out ("s{0:d2}.png" -f $i)), "PNG", $W, $H)
  }
  Write-Output "已导出 $n 页 -> $Out"
} catch {
  throw "PowerPoint 预览导出失败：$($_.Exception.Message)。脚本不会结束你已打开的 PowerPoint；请先确认本机已安装桌面版 PowerPoint。"
} finally {
  if ($pres) {
    $pres.Close()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
  }
  if ($app) {
    if (-not $wasRunning -and $app.Presentations.Count -eq 0) { $app.Quit() }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
  }
}

if ($Contact) {
  if (-not $Python) {
    $cmd = Get-Command python -ErrorAction SilentlyContinue
    if (-not $cmd) { $cmd = Get-Command py -ErrorAction SilentlyContinue }
    if (-not $cmd) { throw "生成总览图需要 Python + Pillow；请用 -Python 指定解释器。" }
    $Python = $cmd.Source
  }
  $py = @'
import sys, glob, os
from PIL import Image, ImageDraw
d, cols = sys.argv[1], 3
fs = sorted(glob.glob(os.path.join(d, "*.png")))
tw, th, pad, lab = 640, 360, 6, 18
rows = (len(fs) + cols - 1) // cols
sheet = Image.new("RGB", (cols*(tw+pad)+pad, rows*(th+pad+lab)+pad), (230, 230, 235))
dr = ImageDraw.Draw(sheet)
for i, f in enumerate(fs):
    im = Image.open(f).convert("RGB").resize((tw, th))
    x, y = pad+(i % cols)*(tw+pad), pad+(i//cols)*(th+pad+lab)
    sheet.paste(im, (x, y+lab)); dr.text((x+4, y+3), os.path.basename(f), fill=(20, 20, 20))
out = os.path.join(d, "_contact.png"); sheet.save(out); print("总览图：" + out)
'@
  $tmp = [System.IO.Path]::GetTempFileName() + ".py"
  Set-Content -Path $tmp -Value $py -Encoding utf8
  & $Python $tmp $Out
  Remove-Item $tmp -Force
}
