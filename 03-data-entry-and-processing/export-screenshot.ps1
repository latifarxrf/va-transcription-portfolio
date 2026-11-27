# export-screenshot.ps1
# Exports the "Perlu Dicek" sheet range (A1:M8) as a PNG for the README.
# Run:  powershell -STA -NoProfile -ExecutionPolicy Bypass -File export-screenshot.ps1
#
# Pipeline: Excel print-renders the range to a 1-page PDF (fixed 52% zoom so all
# 13 columns land inside the page), poppler (WSL) converts PDF -> PNG, then the
# white page margin is trimmed with System.Drawing.

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$xlsx = Join-Path $here 'invoice-tracker-qc.xlsx'
$pdf  = Join-Path $here 'tmp-perlu-dicek.pdf'
$raw  = Join-Path $here 'tmp-perlu-dicek-raw.png'
$png  = Join-Path $here 'invoice-tracker-qc-perlu-dicek.png'

function To-WslPath([string]$p) {
  $p = $p -replace '\\','/'
  if ($p -match '^([A-Za-z]):/(.*)$') { return '/mnt/' + $Matches[1].ToLower() + '/' + $Matches[2] }
  return $p
}

if (-not (Test-Path $xlsx)) { throw "workbook not found: $xlsx" }
foreach ($f in @($png,$pdf,$raw)) { if (Test-Path $f) { Remove-Item $f -Force } }

Add-Type -AssemblyName System.Drawing

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($xlsx)
$ws = $wb.Worksheets.Item('Perlu Dicek')
$ws.Activate()

$ws.PageSetup.PrintArea = '$A$1:$M$8'
$ws.PageSetup.Orientation = 2           # landscape
$ws.PageSetup.Zoom = 52                 # fixed scale: whole A:M fits on one page
$ws.PageSetup.LeftMargin = 10
$ws.PageSetup.RightMargin = 10
$ws.PageSetup.TopMargin = 10
$ws.PageSetup.BottomMargin = 10
$ws.PageSetup.CenterHorizontally = $true
$ws.PageSetup.PrintGridlines = $false
$ws.Cells.Item(1,1).Select() | Out-Null
$ws.ExportAsFixedFormat(0, $pdf)
$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

$wslPdf = To-WslPath $pdf
$prefixW = To-WslPath ($raw -replace '\.png$','')
& wsl -e bash -lc "pdftoppm -png -r 170 -singlefile '$wslPdf' '$prefixW'"
if (-not (Test-Path $raw)) { throw "pdftoppm did not produce $raw" }

# ---- trim white margins (compiled helper: fast pixel scan)
$trimSrc = @'
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
public static class ImgTrim {
  public static string Trim(string inPath, string outPath, int threshold, int pad) {
    using (Bitmap src = new Bitmap(inPath)) {
      int w = src.Width, h = src.Height;
      Rectangle rect = new Rectangle(0, 0, w, h);
      BitmapData data = src.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
      int stride = data.Stride;
      byte[] buf = new byte[stride * h];
      Marshal.Copy(data.Scan0, buf, 0, buf.Length);
      src.UnlockBits(data);
      int minX = w, minY = h, maxX = -1, maxY = -1;
      for (int y = 0; y < h; y++) {
        int row = y * stride;
        for (int x = 0; x < w; x++) {
          int i = row + x * 3;
          byte b = buf[i], g = buf[i + 1], r = buf[i + 2];
          if (r < threshold || g < threshold || b < threshold) {
            if (x < minX) minX = x;
            if (x > maxX) maxX = x;
            if (y < minY) minY = y;
            if (y > maxY) maxY = y;
          }
        }
      }
      if (maxX < 0) { minX = 0; minY = 0; maxX = w - 1; maxY = h - 1; }
      minX = Math.Max(0, minX - pad); minY = Math.Max(0, minY - pad);
      maxX = Math.Min(w - 1, maxX + pad); maxY = Math.Min(h - 1, maxY + pad);
      using (Bitmap dst = new Bitmap(maxX - minX + 1, maxY - minY + 1)) {
        using (Graphics g2 = Graphics.FromImage(dst)) {
          g2.DrawImage(src, new Rectangle(0, 0, dst.Width, dst.Height),
                           new Rectangle(minX, minY, dst.Width, dst.Height), GraphicsUnit.Pixel);
        }
        dst.Save(outPath, ImageFormat.Png);
        return dst.Width + "x" + dst.Height;
      }
    }
  }
}
'@
Add-Type -TypeDefinition $trimSrc -ReferencedAssemblies System.Drawing
$dim = [ImgTrim]::Trim($raw, $png, 245, 14)
Remove-Item $raw -Force
Remove-Item $pdf -Force

$check = [System.Drawing.Image]::FromFile($png)
Write-Output ("OK " + $png + " size=" + $check.Width + "x" + $check.Height + " bytes=" + (Get-Item $png).Length)
$check.Dispose()
