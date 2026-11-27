# build-invoice-tracker.ps1
# Builds the "Invoice Tracker & QC" demo workbook with live Excel formulas.
# Run:  powershell -NoProfile -ExecutionPolicy Bypass -File build-invoice-tracker.ps1
# Requires: Excel desktop (COM). All data is SIMULATED.
#
# Write strategy (important):
#   - Static content is written in ONE bulk Range.Value2 assignment per block
#     (object[,] array). Per-cell numeric puts are avoided because the PowerShell
#     COM adapter caches a property signature per member and rejects mixed types.
#   - All computed cells are written through .Formula (string), which is stable.
#   - Dates are written as OADate numbers; display format is set via NumberFormat.

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$xlsx = Join-Path $here 'invoice-tracker-qc.xlsx'
$png  = Join-Path $here 'invoice-tracker-qc-perlu-dicek.png'

function BGR([int]$r,[int]$g,[int]$b) { return $r + ($g -shl 8) + ($b -shl 16) }
$CLR_INPUT   = BGR 0 0 255
$CLR_LINK    = BGR 0 128 0
$CLR_FORMULA = BGR 0 0 0
$CLR_HEADER  = BGR 31 78 121
$CLR_BAND    = BGR 221 235 247
$CLR_YELLOW  = BGR 255 242 204
$CLR_FLAG    = BGR 255 199 206
$CLR_OK      = BGR 198 239 206
$CLR_GREY    = BGR 242 242 242
$FMT_USD  = '"US$"#,##0.00;"US$"-#,##0.00'
$FMT_DATE = 'dd-mmm-yyyy'

function New-Grid([int]$rows,[int]$cols) {
  $lengths = New-Object 'int[]' 2
  $lengths[0] = $rows
  $lengths[1] = $cols
  # comma prevents PowerShell from unrolling the 2-D array on return
  return ,([Array]::CreateInstance([object], $lengths))
}
function D([string]$iso) {
  return ([datetime]::ParseExact($iso,'yyyy-MM-dd',$null)).ToOADate()
}
function Style-Header($ws,[string]$addr) {
  $r = $ws.Range($addr)
  $r.Font.Name='Arial'; $r.Font.Size=10; $r.Font.Bold=$true; $r.Font.Color=BGR 255 255 255
  $r.Interior.Color=$CLR_HEADER
  $r.HorizontalAlignment=-4108; $r.VerticalAlignment=-4108
  $r.WrapText=$true
}
function Style-Body($ws,[string]$addr) {
  $r = $ws.Range($addr)
  $r.Font.Name='Arial'; $r.Font.Size=10
  $r.VerticalAlignment=-4160
  try { $r.Borders.LineStyle=1; $r.Borders.Weight=2; $r.Borders.ColorIndex=16 } catch {}
}
function Style-Title($ws,[string]$addr,[int]$size) {
  $r = $ws.Range($addr)
  $r.Font.Name='Arial'; $r.Font.Size=$size; $r.Font.Bold=$true; $r.Font.Color=$CLR_HEADER
}
function Style-Note($ws,[string]$addr) {
  $r = $ws.Range($addr)
  $r.Font.Name='Arial'; $r.Font.Size=9; $r.Font.Italic=$true; $r.Font.Color=BGR 89 89 89
}

# ---------------------------------------------------------------- source data
# Simulated 3-way match source (10 rows from invoice-reconciliation-3-way-matching.csv
# + 1 deliberately added case where the GRN document does not exist yet).
$sim = @(
  @('REC-2026-001','PO-89421','2026-10-15','Amazon Web Services',         4520.00,'GRN-0981','2026-10-31','Service Rendered',   'INV-AWS-882194','2026-11-01',4520.00,'Exact match','Finance Lead'),
  @('REC-2026-002','PO-89422','2026-10-18','Google Workspace Enterprise', 1840.00,'GRN-0982','2026-10-31','Service Rendered',   'INV-GOOG-41209','2026-11-01',1840.00,'Exact match','Finance Lead'),
  @('REC-2026-003','PO-89423','2026-10-20','Apex Office Equipment Ltd',   3250.00,'GRN-0985','2026-10-28','Goods Received 100%','INV-APX-77412','2026-10-29',3250.00,'Exact match','Ops Manager'),
  @('REC-2026-004','PO-89424','2026-10-22','WeWork Shared Spaces',        5600.00,'GRN-0988','2026-10-31','Facility Verified',  'INV-WW-994102','2026-11-02',5600.00,'Exact match','Finance Lead'),
  @('REC-2026-005','PO-89425','2026-10-25','FastTrack Global Couriers',    890.00,'GRN-0991','2026-10-30','Deliveries Complete','INV-FTC-10941','2026-10-31', 915.00,'Fuel surcharge adjusted per contract','Logistics Lead'),
  @('REC-2026-006','PO-89426','2026-10-27','Slack Technologies',          2400.00,'GRN-0994','2026-10-31','Service Rendered',   'INV-SLK-33819','2026-11-01',2400.00,'Exact match','IT Director'),
  @('REC-2026-007','PO-89427','2026-10-28','Delta Airlines Corporate',    4120.00,'GRN-0997','2026-10-29','Ticket Issued',      'INV-DAL-78219','2026-10-29',4120.00,'Exact match','Executive VA'),
  @('REC-2026-008','PO-89428','2026-10-30','Zendesk Support Platform',    1650.00,'GRN-1002','2026-10-31','Service Rendered',   'INV-ZEN-90184','2026-11-01',1650.00,'Exact match','CS Lead'),
  @('REC-2026-009','PO-89429','2026-11-02','Metro Catering Services',    1420.00,'GRN-1005','2026-11-05','Event Completed',    'INV-MCS-55102','2026-11-06',1420.00,'Exact match','Office Admin'),
  @('REC-2026-010','PO-89430','2026-11-04','Canva Enterprise Design',      600.00,'GRN-1008','2026-11-04','Seats Activated',    'INV-CNV-10294','2026-11-04', 600.00,'Exact match','Marketing Lead'),
  @('REC-2026-011','PO-89431','2026-11-06','Northwind Office Supplies',   1250.00,'',         '','Belum ada GRN',     'INV-NWS-48210','2026-11-07',1250.00,'Kasus tambahan (simulasi): berita terima barang belum tersedia saat rekonsiliasi','Ops Manager')
)
$nSim = $sim.Count

# ---------------------------------------------------------------- start Excel
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()

function Set-Sheet($wb, [string]$name, [int]$after) {
  $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($after))
  $ws.Name = $name
  return $ws
}
$wsPend = $wb.Worksheets.Item(1); $wsPend.Name = 'Pendahuluan'
$wsData = Set-Sheet $wb 'Data Simulasi' 1
$wsMat  = Set-Sheet $wb 'Pencocokan Invoice' 2
$wsFlag = Set-Sheet $wb 'Perlu Dicek' 3
$wsSum  = Set-Sheet $wb 'Ringkasan ke Pemilik' 4

# ================================================================ 1. Pendahuluan
$pendRows = @(
  [pscustomobject]@{ A='Invoice Tracker & QC - 3-Way Match (PO / GRN / Invoice)'; B=$null }
  [pscustomobject]@{ A='SIMULASI - seluruh data, vendor, dan nomor dokumen di workbook ini sintetis (latihan terstruktur), bukan pekerjaan klien.'; B=$null }
  [pscustomobject]@{ A=''; B=$null }
  [pscustomobject]@{ A='1. Apa isi workbook ini'; B=$null }
  [pscustomobject]@{ A='Tujuan'; B='Menunjukkan proses QA 3-way match (Purchase Order vs Goods Receipt Note vs Invoice), cara selisih dihitung, dan bagaimana transaksi yang perlu ditinjau disiapkan untuk bisnis.' }
  [pscustomobject]@{ A='Isi'; B='Tab Data Simulasi (input) - Tab Pencocokan Invoice (proses) - Tab Perlu Dicek (output) - Tab Ringkasan ke Pemilik (contoh hasil kerja VA).' }
  [pscustomobject]@{ A='Cara pakai'; B='Isi hanya kolom biru pada tab Data Simulasi. Seluruh angka, selisih, status, dan daftar Perlu Dicek dihitung ulang otomatis oleh formula.' }
  [pscustomobject]@{ A='Sumber data simulasi'; B='03-data-entry-and-processing/invoice-reconciliation-3-way-matching.csv (10 transaksi) + 1 kasus tambahan tanpa GRN.' }
  [pscustomobject]@{ A='Referensi'; B='AlexTheAnalyst/Excel-Tutorial - Data Cleaning (praktik merapikan format dan menyiapkan data di Excel), diterapkan pada kasus sendiri di sini.' }
  [pscustomobject]@{ A=''; B=$null }
  [pscustomobject]@{ A='2. Asumsi yang saya buat (dokumentasi)'; B=$null }
  [pscustomobject]@{ A='Asumsi 1'; B='Sumber CSV asli tidak memuat kolom nilai GRN, sehingga nilai barang/service diterima diambil dari nilai PO (barang/service diterima penuh sesuai grn_status). Kolom ini berformula dan bisa diganti manual untuk penerimaan sebagian/diskon.' }
  [pscustomobject]@{ A='Asumsi 2'; B='Toleransi selisih default US$ 0,00 (kontrol di tab Pencocokan Invoice, sel B3). Selisih di atas toleransi memicu status Perlu Dicek.' }
  [pscustomobject]@{ A='Asumsi 3'; B='Selisih antar dokumen dihitung dalam USD, tidak memperhitungkan pajak atau biaya bank karena data sumber tidak memisahkannya.' }
  [pscustomobject]@{ A=''; B=$null }
  [pscustomobject]@{ A='3. Legenda warna'; B=$null }
  [pscustomobject]@{ A='Biru'; B='Data input yang diketik manual (nilai sumber).' }
  [pscustomobject]@{ A='Hijau'; B='Formula yang mengambil nilai dari tab lain.' }
  [pscustomobject]@{ A='Hitam'; B='Formula di dalam tab yang sama.' }
  [pscustomobject]@{ A='Kuning'; B='Sel yang perlu diisi atau diubah (parameter dan kolom status tinjauan).' }
  [pscustomobject]@{ A='Merah muda'; B='Baris yang masuk daftar Perlu Dicek.' }
  [pscustomobject]@{ A=''; B=$null }
  [pscustomobject]@{ A='4. Cara saya menemukan selisih dan menyiapkan tindak lanjut'; B=$null }
  [pscustomobject]@{ A='Langkah 1'; B='Bandingkan invoice vs PO dan invoice vs nilai GRN dengan formula (kolom K-M pada tab Pencocokan Invoice).' }
  [pscustomobject]@{ A='Langkah 2'; B='Cek kelengkapan dokumen: nomor dan status GRN harus ada. GRN kosong berarti transaksi tidak bisa dinyatakan cocok walau angkanya sama.' }
  [pscustomobject]@{ A='Langkah 3'; B='Tandai transaksi berstatus Perlu Dicek, lalu susun daftar pada tab Perlu Dicek: alasan, tindak lanjut, dan PIC yang perlu dihubungi.' }
  [pscustomobject]@{ A='Langkah 4'; B='Ringkas untuk pemilik bisnis pada tab Ringkasan ke Pemilik (nomor invoice, alasan, tindakan, total selisih) agar tim tahu apa yang perlu diputuskan.' }
)
$g = New-Grid $pendRows.Count 2
for ($i=0; $i -lt $pendRows.Count; $i++) {
  $g.SetValue($pendRows[$i].A,$i,0)
  if ($pendRows[$i].B -ne $null) { $g.SetValue($pendRows[$i].B,$i,1) }
}
$wsPend.Range('A1').Resize($pendRows.Count,2).Value2 = $g
Style-Title $wsPend 'A1' 16
$wsPend.Range('A2').Font.Name='Arial'; $wsPend.Range('A2').Font.Size=10; $wsPend.Range('A2').Font.Italic=$true; $wsPend.Range('A2').Font.Color=BGR 192 0 0; $wsPend.Range('A2').Font.Bold=$true
foreach ($rr in @(4,11,16,23)) {
  $c = $wsPend.Range("A$rr")
  $c.Font.Name='Arial'; $c.Font.Size=11; $c.Font.Bold=$true; $c.Font.Color=$CLR_HEADER
}
$wsPend.Range('A5:A27').Font.Name='Arial'; $wsPend.Range('A5:A27').Font.Size=10; $wsPend.Range('A5:A27').Font.Bold=$true
$wsPend.Range('B5:B27').Font.Name='Arial'; $wsPend.Range('B5:B27').Font.Size=10
$wsPend.Range('B5:B27').WrapText=$true
$wsPend.Range('A5:A27').VerticalAlignment=-4160
$wsPend.Range('B5:B27').VerticalAlignment=-4160
$wsPend.Columns.Item(1).ColumnWidth=26
$wsPend.Columns.Item(2).ColumnWidth=105
$wsPend.Range('A17:A21').Interior.Color = $CLR_BAND

# ================================================================ 2. Data Simulasi (input)
$hdr = @('reconciliation_id','po_number','po_date','vendor_name','po_amount_usd','grn_number','grn_date','grn_status','received_value_usd','invoice_number','invoice_date','invoice_amount_usd','catatan_sumber','pic_internal')
$g = New-Grid ($nSim+1) 14
for ($i=0; $i -lt $hdr.Count; $i++) { $g.SetValue($hdr[$i],0,$i) }
for ($r=0; $r -lt $nSim; $r++) {
  $s = $sim[$r]; $row = $r + 1
  $g.SetValue($s[0],$row,0); $g.SetValue($s[1],$row,1); $g.SetValue((D $s[2]),$row,2)
  $g.SetValue($s[3],$row,3); $g.SetValue([double]$s[4],$row,4)
  if ($s[5] -ne '') { $g.SetValue($s[5],$row,5) }
  if ($s[6] -ne '') { $g.SetValue((D $s[6]),$row,6) }
  $g.SetValue($s[7],$row,7)
  # col 8 (received_value_usd) = formula, written later
  $g.SetValue($s[8],$row,9); $g.SetValue((D $s[9]),$row,10)
  $g.SetValue([double]$s[10],$row,11); $g.SetValue($s[11],$row,12); $g.SetValue($s[12],$row,13)
}
$wsData.Range('A1').Resize($nSim+1,14).Value2 = $g
$lastData = $nSim + 1
for ($r=0; $r -lt $nSim; $r++) {
  $row = $r + 2
  $wsData.Cells.Item($row,9).Formula = ('=IF(F{0}="","",E{0})' -f $row)
}
Style-Header $wsData 'A1:N1'
$wsData.Rows.Item(1).RowHeight = 30
Style-Body $wsData ("A2:N{0}" -f $lastData)
$wsData.Range("A2:H{0}" -f $lastData).Font.Color = $CLR_INPUT
$wsData.Range("J2:N{0}" -f $lastData).Font.Color = $CLR_INPUT
$wsData.Range("I2:I{0}" -f $lastData).Font.Color = $CLR_FORMULA
$wsData.Range("C2:C{0}" -f $lastData).NumberFormat = $FMT_DATE
$wsData.Range("G2:G{0}" -f $lastData).NumberFormat = $FMT_DATE
$wsData.Range("K2:K{0}" -f $lastData).NumberFormat = $FMT_DATE
$wsData.Range("E2:E{0}" -f $lastData).NumberFormat = $FMT_USD
$wsData.Range("I2:I{0}" -f $lastData).NumberFormat = $FMT_USD
$wsData.Range("L2:L{0}" -f $lastData).NumberFormat = $FMT_USD
$wsData.Range("A2:N{0}" -f $lastData).Interior.Color = $CLR_BAND
$wsData.Range("A12:N12").Interior.Color = $CLR_FLAG
Style-Note $wsData ("A{0}" -f ($lastData+2))
$wsData.Range(("A{0}" -f ($lastData+2))).Value2 = 'Baris terakhir (REC-2026-011) adalah kasus tambahan untuk demonstrasi: dokumen GRN belum ada sehingga transaksi wajib ditinjau meski angka invoice = PO.'
$wsData.Range(("A{0}" -f ($lastData+3))).Value2 = 'Kolom I berformula = nilai PO selama GRN ada (asumsi: barang/service diterima penuh). Untuk penerimaan sebagian atau diskon, ganti formulanya dengan nilai aktual.'
Style-Note $wsData ("A{0}" -f ($lastData+3))
$dataWidths = @(16,12,12,30,14,12,12,20,18,18,12,18,46,15)
for ($i=0; $i -lt $dataWidths.Count; $i++) { $wsData.Columns.Item($i+1).ColumnWidth = $dataWidths[$i] }

# ================================================================ 3. Pencocokan Invoice (proses)
$g = New-Grid 10 3
$g.SetValue('Pencocokan Invoice - 3-Way Match (PO / GRN / Invoice)',0,0)
$g.SetValue('SIMULASI. Semua selisih, status, dan ringkasan dihitung oleh formula dari tab Data Simulasi.',1,0)
$g.SetValue('Toleransi selisih (USD)',2,0)
$g.SetValue([double]0,2,1)
$g.SetValue('Selisih di atas toleransi (atau GRN belum ada) = status Perlu Dicek.',2,2)
$g.SetValue('Ringkasan cepat (formula)',4,0)
$g.SetValue('Total transaksi direkonsiliasi',5,0)
$g.SetValue('Transaksi cocok (3-way match)',6,0)
$g.SetValue('Transaksi perlu dicek',7,0)
$g.SetValue('Total nilai invoice (USD)',8,0)
$g.SetValue('Total selisih invoice vs PO (USD)',9,0)
$wsMat.Range('A1').Resize(10,3).Value2 = $g

$sumFormulas = @(
 '=COUNTA($A$13:$A$23)',
 '=COUNTIF($O$13:$O$23,"Match")',
 '=COUNTIF($O$13:$O$23,"Perlu Dicek")',
 '=SUM($J$13:$J$23)',
 '=SUM($K$13:$K$23)'
)
for ($i=0; $i -lt $sumFormulas.Count; $i++) { $wsMat.Cells.Item(6+$i,2).Formula = $sumFormulas[$i] }

$mHdr = @('reconciliation_id','vendor_name','po_number','po_amount_usd','grn_number','grn_status','received_value_usd','invoice_number','invoice_date','invoice_amount_usd','var_invoice_vs_po_usd','var_receipt_vs_invoice_usd','cek_nilai','cek_dokumen','status_qc','tindak_lanjut','no_flag','catatan_sumber','pic_internal')
$g = New-Grid 1 19
for ($i=0; $i -lt $mHdr.Count; $i++) { $g.SetValue($mHdr[$i],0,$i) }
$wsMat.Range('A12').Resize(1,19).Value2 = $g
Style-Header $wsMat 'A12:S12'
$wsMat.Rows.Item(12).RowHeight = 32

# formula templates: {0} = baris pada tab ini, {1} = baris pada tab Data Simulasi
$fA = @'
='Data Simulasi'!A{1}
'@
$fB = @'
='Data Simulasi'!D{1}
'@
$fC = @'
='Data Simulasi'!B{1}
'@
$fD = @'
='Data Simulasi'!E{1}
'@
$fE = @'
=IF('Data Simulasi'!F{1}="","(belum ada)",'Data Simulasi'!F{1})
'@
$fF = @'
='Data Simulasi'!H{1}
'@
$fG = @'
=IF('Data Simulasi'!I{1}="","",'Data Simulasi'!I{1})
'@
$fH = @'
='Data Simulasi'!J{1}
'@
$fI = @'
='Data Simulasi'!K{1}
'@
$fJ = @'
='Data Simulasi'!L{1}
'@
$fK = @'
=ROUND(J{0}-D{0},2)
'@
$fL = @'
=IF(G{0}="","n/a",ROUND(J{0}-G{0},2))
'@
$fM = @'
=IF(G{0}="","n/a",IF(ABS(L{0})<=$B$3,"Sesuai","Selisih"))
'@
$fN = @'
=IF('Data Simulasi'!F{1}="","GRN belum ada","GRN ada")
'@
$fO = @'
=IF(OR(N{0}="GRN belum ada",M{0}="Selisih"),"Perlu Dicek","Match")
'@
$fP = @'
=IF(O{0}="Match","Tidak ada tindak lanjut - arsipkan sebagai transaksi cocok",IF(N{0}="GRN belum ada","Minta salinan GRN / berita terima dari vendor ("&B{0}&") sebelum menyetujui pembayaran","Konfirmasi ke vendor: minta faktur revisi atau keterangan tambahan untuk selisih US$"&TEXT(K{0},"#,##0.00")))
'@
$fQ = @'
=IF(O{0}="Perlu Dicek",COUNTIF($O$13:O{0},"Perlu Dicek"),"")
'@
$fR = @'
='Data Simulasi'!M{1}
'@
$fS = @'
='Data Simulasi'!N{1}
'@

for ($r=0; $r -lt $nSim; $r++) {
  $row = 13 + $r
  $src = 2 + $r
  $wsMat.Cells.Item($row,1).Formula  = $fA -f $row,$src
  $wsMat.Cells.Item($row,2).Formula  = $fB -f $row,$src
  $wsMat.Cells.Item($row,3).Formula  = $fC -f $row,$src
  $wsMat.Cells.Item($row,4).Formula  = $fD -f $row,$src
  $wsMat.Cells.Item($row,5).Formula  = $fE -f $row,$src
  $wsMat.Cells.Item($row,6).Formula  = $fF -f $row,$src
  $wsMat.Cells.Item($row,7).Formula  = $fG -f $row,$src
  $wsMat.Cells.Item($row,8).Formula  = $fH -f $row,$src
  $wsMat.Cells.Item($row,9).Formula  = $fI -f $row,$src
  $wsMat.Cells.Item($row,10).Formula = $fJ -f $row,$src
  $wsMat.Cells.Item($row,11).Formula = $fK -f $row
  $wsMat.Cells.Item($row,12).Formula = $fL -f $row
  $wsMat.Cells.Item($row,13).Formula = $fM -f $row
  $wsMat.Cells.Item($row,14).Formula = $fN -f $row,$src
  $wsMat.Cells.Item($row,15).Formula = $fO -f $row
  $wsMat.Cells.Item($row,16).Formula = $fP -f $row
  $wsMat.Cells.Item($row,17).Formula = $fQ -f $row
  $wsMat.Cells.Item($row,18).Formula = $fR -f $row,$src
  $wsMat.Cells.Item($row,19).Formula = $fS -f $row,$src
}
$lastMat = 12 + $nSim
Style-Body $wsMat ("A13:S{0}" -f $lastMat)
$wsMat.Range("A13:J{0}" -f $lastMat).Font.Color = $CLR_LINK
$wsMat.Range("K13:Q{0}" -f $lastMat).Font.Color = $CLR_FORMULA
$wsMat.Range("R13:S{0}" -f $lastMat).Font.Color = $CLR_LINK
$wsMat.Range("D13:D{0}" -f $lastMat).NumberFormat = $FMT_USD
$wsMat.Range("G13:G{0}" -f $lastMat).NumberFormat = $FMT_USD
$wsMat.Range("J13:L{0}" -f $lastMat).NumberFormat = $FMT_USD
$wsMat.Range("I13:I{0}" -f $lastMat).NumberFormat = $FMT_DATE
$wsMat.Range("A13:S{0}" -f $lastMat).Interior.Color = $CLR_BAND
$wsMat.Range("P13:P{0}" -f $lastMat).WrapText = $true
$wsMat.Range("R13:R{0}" -f $lastMat).WrapText = $true
$mWidths = @(16,26,12,14,13,20,16,18,12,16,15,16,11,14,12,42,8,40,14)
for ($i=0; $i -lt $mWidths.Count; $i++) { $wsMat.Columns.Item($i+1).ColumnWidth = $mWidths[$i] }
Style-Title $wsMat 'A1' 14
$wsMat.Range('A2').Font.Name='Arial'; $wsMat.Range('A2').Font.Size=10; $wsMat.Range('A2').Font.Italic=$true
$wsMat.Range('A3').Font.Name='Arial'; $wsMat.Range('A3').Font.Size=10; $wsMat.Range('A3').Font.Bold=$true
$wsMat.Range('B3').Font.Name='Arial'; $wsMat.Range('B3').Font.Size=10; $wsMat.Range('B3').Font.Color=$CLR_INPUT
$wsMat.Range('B3').Interior.Color = $CLR_YELLOW
$wsMat.Range('B3').NumberFormat = $FMT_USD
$wsMat.Range('C3').Font.Name='Arial'; $wsMat.Range('C3').Font.Size=9; $wsMat.Range('C3').Font.Italic=$true
$wsMat.Range('A5').Font.Name='Arial'; $wsMat.Range('A5').Font.Size=11; $wsMat.Range('A5').Font.Bold=$true; $wsMat.Range('A5').Font.Color=$CLR_HEADER
$wsMat.Range('A6:A10').Font.Name='Arial'; $wsMat.Range('A6:A10').Font.Size=10; $wsMat.Range('A6:A10').Font.Bold=$true
$wsMat.Range('B6:B10').Font.Name='Arial'; $wsMat.Range('B6:B10').Font.Size=10; $wsMat.Range('B6:B10').Font.Bold=$true
$wsMat.Range('B9:B10').NumberFormat = $FMT_USD
$wsMat.Range('B6:B8').HorizontalAlignment = -4131
$wsMat.Range('A6:B10').Interior.Color = $CLR_GREY
Style-Note $wsMat ("A{0}" -f ($lastMat+2))
$wsMat.Range(("A{0}" -f ($lastMat+2))).Value2 = 'Kolom Q hanya untuk mengurutkan daftar Perlu Dicek (helper). Abaikan kolom ini saat membaca laporan.'
Style-Note $wsMat ("A{0}" -f ($lastMat+3))
$wsMat.Range(("A{0}" -f ($lastMat+3))).Value2 = 'Baris merah muda = transaksi berstatus Perlu Dicek (diwarnai otomatis oleh formula status_qc di kolom O).'

# ================================================================ 4. Perlu Dicek (output)
$g = New-Grid 2 13
$g.SetValue('Output workbook: baris diambil otomatis dari tab Pencocokan Invoice. Kolom kuning = diisi manual setelah menghubungi vendor/PIC.',0,0)
$fHdr = @('no','reconciliation_id','vendor_name','invoice_number','invoice_date','po_amount_usd','received_value_usd','invoice_amount_usd','selisih_vs_po_usd','jenis_masalah','tindak_lanjut','pic_yang_harus_dihubungi','status_tinjauan')
for ($i=0; $i -lt $fHdr.Count; $i++) { $g.SetValue($fHdr[$i],1,$i) }
$wsFlag.Range('A2').Resize(2,13).Value2 = $g
$wsFlag.Range('A1').Formula = '="Perlu Dicek - "&COUNTIF(''Pencocokan Invoice''!$O$13:$O$23,"Perlu Dicek")&" transaksi perlu ditinjau sebelum pembayaran (SIMULASI)"'
Style-Title $wsFlag 'A1' 14
$wsFlag.Range('A2').Font.Name='Arial'; $wsFlag.Range('A2').Font.Size=10; $wsFlag.Range('A2').Font.Italic=$true
Style-Header $wsFlag 'A3:M3'
$wsFlag.Rows.Item(3).RowHeight = 32

$gA = @'
=IF(COUNTIF('Pencocokan Invoice'!$Q$13:$Q$23,{1})=0,"",{1})
'@
$gB = @'
=IF($A{0}="","",INDEX('Pencocokan Invoice'!$A$13:$A$23,MATCH($A{0},'Pencocokan Invoice'!$Q$13:$Q$23,0)))
'@
$gC = @'
=IF($A{0}="","",INDEX('Pencocokan Invoice'!$B$13:$B$23,MATCH($A{0},'Pencocokan Invoice'!$Q$13:$Q$23,0)))
'@
$gD = @'
=IF($A{0}="","",INDEX('Pencocokan Invoice'!$H$13:$H$23,MATCH($A{0},'Pencocokan Invoice'!$Q$13:$Q$23,0)))
'@
$gE = @'
=IF($A{0}="","",INDEX('Pencocokan Invoice'!$I$13:$I$23,MATCH($A{0},'Pencocokan Invoice'!$Q$13:$Q$23,0)))
'@
$gF = @'
=IF($A{0}="","",INDEX('Pencocokan Invoice'!$D$13:$D$23,MATCH($A{0},'Pencocokan Invoice'!$Q$13:$Q$23,0)))
'@
$gG = @'
=IF($A{0}="","",IF(INDEX('Pencocokan Invoice'!$G$13:$G$23,MATCH($A{0},'Pencocokan Invoice'!$Q$13:$Q$23,0))="","(belum ada)",INDEX('Pencocokan Invoice'!$G$13:$G$23,MATCH($A{0},'Pencocokan Invoice'!$Q$13:$Q$23,0))))
'@
$gH = @'
=IF($A{0}="","",INDEX('Pencocokan Invoice'!$J$13:$J$23,MATCH($A{0},'Pencocokan Invoice'!$Q$13:$Q$23,0)))
'@
$gI = @'
=IF($A{0}="","",INDEX('Pencocokan Invoice'!$K$13:$K$23,MATCH($A{0},'Pencocokan Invoice'!$Q$13:$Q$23,0)))
'@
$gJ = @'
=IF($A{0}="","",IF(INDEX('Pencocokan Invoice'!$N$13:$N$23,MATCH($A{0},'Pencocokan Invoice'!$Q$13:$Q$23,0))="GRN belum ada","Dokumen GRN belum ada - angka sama dengan PO (selisih US$"&TEXT($I{0},"#,##0.00")&"), tetapi dokumen penerimaan tidak terlampir","Selisih US$"&TEXT($I{0},"#,##0.00")))
'@
$gK = @'
=IF($A{0}="","",INDEX('Pencocokan Invoice'!$P$13:$P$23,MATCH($A{0},'Pencocokan Invoice'!$Q$13:$Q$23,0)))
'@
$gL = @'
=IF($A{0}="","",INDEX('Pencocokan Invoice'!$B$13:$B$23,MATCH($A{0},'Pencocokan Invoice'!$Q$13:$Q$23,0))&" + PIC internal: "&INDEX('Pencocokan Invoice'!$S$13:$S$23,MATCH($A{0},'Pencocokan Invoice'!$Q$13:$Q$23,0)))
'@
$gM = @'
=IF($A{0}="","","Menunggu konfirmasi vendor")
'@

$slots = 5
for ($i=0; $i -lt $slots; $i++) {
  $row = 4 + $i
  $n = $i + 1
  $wsFlag.Cells.Item($row,1).Formula  = $gA -f $row,$n
  $wsFlag.Cells.Item($row,2).Formula  = $gB -f $row
  $wsFlag.Cells.Item($row,3).Formula  = $gC -f $row
  $wsFlag.Cells.Item($row,4).Formula  = $gD -f $row
  $wsFlag.Cells.Item($row,5).Formula  = $gE -f $row
  $wsFlag.Cells.Item($row,6).Formula  = $gF -f $row
  $wsFlag.Cells.Item($row,7).Formula  = $gG -f $row
  $wsFlag.Cells.Item($row,8).Formula  = $gH -f $row
  $wsFlag.Cells.Item($row,9).Formula  = $gI -f $row
  $wsFlag.Cells.Item($row,10).Formula = $gJ -f $row
  $wsFlag.Cells.Item($row,11).Formula = $gK -f $row
  $wsFlag.Cells.Item($row,12).Formula = $gL -f $row
  $wsFlag.Cells.Item($row,13).Formula = $gM -f $row
}
$lastFlag = 3 + $slots
Style-Body $wsFlag ("A4:M{0}" -f $lastFlag)
$wsFlag.Range("A4:I{0}" -f $lastFlag).Font.Color = $CLR_LINK
$wsFlag.Range("J4:L{0}" -f $lastFlag).Font.Color = $CLR_FORMULA
$wsFlag.Range("M4:M{0}" -f $lastFlag).Font.Color = $CLR_INPUT
$wsFlag.Range("M4:M{0}" -f $lastFlag).Interior.Color = $CLR_YELLOW
$wsFlag.Range("E4:E{0}" -f $lastFlag).NumberFormat = $FMT_DATE
$wsFlag.Range("F4:I{0}" -f $lastFlag).NumberFormat = $FMT_USD
$wsFlag.Range("A4:M{0}" -f $lastFlag).Interior.Color = $CLR_BAND
$wsFlag.Range("J4:K{0}" -f $lastFlag).WrapText = $true
$wsFlag.Range("K4:K{0}" -f $lastFlag).Interior.Color = $CLR_FLAG
$flagWidths = @(5,16,26,18,12,14,16,16,15,30,46,28,24)
for ($i=0; $i -lt $flagWidths.Count; $i++) { $wsFlag.Columns.Item($i+1).ColumnWidth = $flagWidths[$i] }
$noteRow = $lastFlag + 2
$wsFlag.Range(("A{0}" -f $noteRow)).Value2 = 'Cara mengisi: kolom L (PIC) dan M (status tinjauan) diisi manual setelah menghubungi vendor/PIC. Baris kosong berarti tidak ada transaksi lain yang perlu ditinjau.'
$wsFlag.Range(("A{0}" -f ($noteRow+1))).Value2 = 'Semua baris, nomor invoice, dan nominal di atas adalah data simulasi untuk demonstrasi proses QA.'
Style-Note $wsFlag ("A{0}" -f $noteRow)
Style-Note $wsFlag ("A{0}" -f ($noteRow+1))
$wsFlag.Rows.Item(4).Resize($slots,13).RowHeight = 44

# ================================================================ 5. Ringkasan ke Pemilik
$sumRows = @(
  [pscustomobject]@{ A='Contoh Hasil Kerja VA - Ringkasan ke Pemilik Bisnis'; B=$null }
  [pscustomobject]@{ A='SIMULASI. Semua nominal dan nomor invoice di bawah diambil dari tab Pencocokan Invoice lewat formula, jadi ringkasan selalu cocok dengan data.'; B=$null }
  [pscustomobject]@{ A=''; B=$null }
  [pscustomobject]@{ A='Contoh 1 - Email ringkas ke pemilik bisnis'; B=$null }
  [pscustomobject]@{ A='Subjek'; B=$null }
  [pscustomobject]@{ A='Yth'; B='Pemilik Bisnis / Finance Lead' }
  [pscustomobject]@{ A='Isi email'; B=$null }
  [pscustomobject]@{ A=''; B=$null }
  [pscustomobject]@{ A='Contoh 2 - Pesan singkat (WhatsApp / chat internal)'; B=$null }
  [pscustomobject]@{ A=''; B=$null }
  [pscustomobject]@{ A=''; B=$null }
  [pscustomobject]@{ A='Cakupan pekerjaan yang bisa saya kerjakan dalam trial 5 jam'; B=$null }
  [pscustomobject]@{ A='1. Merapikan format spreadsheet: header, penamaan kolom, format tanggal, mata uang, dan konsistensi penulisan.'; B=$null }
  [pscustomobject]@{ A='2. Pencocokan 3-way (PO vs GRN vs Invoice) dengan formula yang bisa diaudit, plus pemeriksaan kelengkapan dokumen.'; B=$null }
  [pscustomobject]@{ A='3. Menghitung selisih per transaksi dan menandai yang di luar toleransi.'; B=$null }
  [pscustomobject]@{ A='4. Menyusun daftar Perlu Dicek: alasan, tindak lanjut, dan PIC yang perlu dihubungi.'; B=$null }
  [pscustomobject]@{ A='5. Ringkasan singkat untuk pemilik bisnis, lalu serah terima file dan 1x revisi minor.'; B=$null }
  [pscustomobject]@{ A=''; B=$null }
  [pscustomobject]@{ A='Catatan: file ini adalah simulasi. Saat trial, angka dihitung ulang dari data asli Anda, bukan dari dataset latihan ini.'; B=$null }
)
$g = New-Grid $sumRows.Count 2
for ($i=0; $i -lt $sumRows.Count; $i++) {
  $g.SetValue($sumRows[$i].A,$i,0)
  if ($sumRows[$i].B -ne $null) { $g.SetValue($sumRows[$i].B,$i,1) }
}
$wsSum.Range('A1').Resize($sumRows.Count,2).Value2 = $g

$wsSum.Range('B5').Formula = @'
="Ringkasan rekonsiliasi invoice - "&COUNTIF('Pencocokan Invoice'!$O$13:$O$23,"Perlu Dicek")&" transaksi perlu ditinjau sebelum November ditutup"
'@
$wsSum.Range('B7').Formula = @'
="Selamat siang,"&CHAR(10)&"Rekonsiliasi invoice periode 1-7 November 2026 sudah selesai: "&COUNTA('Pencocokan Invoice'!$A$13:$A$23)&" invoice, total US$"&TEXT(SUM('Pencocokan Invoice'!$J$13:$J$23),"#,##0.00")&", "&COUNTIF('Pencocokan Invoice'!$O$13:$O$23,"Match")&" cocok tanpa selisih."&CHAR(10)&"Dua transaksi perlu ditinjau sebelum pembayaran:"&CHAR(10)&"1) "&INDEX('Pencocokan Invoice'!$H$13:$H$23,MATCH(1,'Pencocokan Invoice'!$Q$13:$Q$23,0))&" ("&INDEX('Pencocokan Invoice'!$B$13:$B$23,MATCH(1,'Pencocokan Invoice'!$Q$13:$Q$23,0))&") - "&INDEX('Pencocokan Invoice'!$N$13:$N$23,MATCH(1,'Pencocokan Invoice'!$Q$13:$Q$23,0))&", invoice US$"&TEXT(INDEX('Pencocokan Invoice'!$J$13:$J$23,MATCH(1,'Pencocokan Invoice'!$Q$13:$Q$23,0)),"#,##0.00")&" vs PO US$"&TEXT(INDEX('Pencocokan Invoice'!$D$13:$D$23,MATCH(1,'Pencocokan Invoice'!$Q$13:$Q$23,0)),"#,##0.00")&", selisih US$"&TEXT(INDEX('Pencocokan Invoice'!$K$13:$K$23,MATCH(1,'Pencocokan Invoice'!$Q$13:$Q$23,0)),"#,##0.00")&" ("&INDEX('Pencocokan Invoice'!$R$13:$R$23,MATCH(1,'Pencocokan Invoice'!$Q$13:$Q$23,0))&"). Tindakan: "&INDEX('Pencocokan Invoice'!$P$13:$P$23,MATCH(1,'Pencocokan Invoice'!$Q$13:$Q$23,0))&" PIC internal: "&INDEX('Pencocokan Invoice'!$S$13:$S$23,MATCH(1,'Pencocokan Invoice'!$Q$13:$Q$23,0))&"."&CHAR(10)&"2) "&INDEX('Pencocokan Invoice'!$H$13:$H$23,MATCH(2,'Pencocokan Invoice'!$Q$13:$Q$23,0))&" ("&INDEX('Pencocokan Invoice'!$B$13:$B$23,MATCH(2,'Pencocokan Invoice'!$Q$13:$Q$23,0))&") - "&INDEX('Pencocokan Invoice'!$N$13:$N$23,MATCH(2,'Pencocokan Invoice'!$Q$13:$Q$23,0))&", invoice US$"&TEXT(INDEX('Pencocokan Invoice'!$J$13:$J$23,MATCH(2,'Pencocokan Invoice'!$Q$13:$Q$23,0)),"#,##0.00")&" vs PO US$"&TEXT(INDEX('Pencocokan Invoice'!$D$13:$D$23,MATCH(2,'Pencocokan Invoice'!$Q$13:$Q$23,0)),"#,##0.00")&", selisih US$"&TEXT(INDEX('Pencocokan Invoice'!$K$13:$K$23,MATCH(2,'Pencocokan Invoice'!$Q$13:$Q$23,0)),"#,##0.00")&". Tindakan: "&INDEX('Pencocokan Invoice'!$P$13:$P$23,MATCH(2,'Pencocokan Invoice'!$Q$13:$Q$23,0))&" PIC internal: "&INDEX('Pencocokan Invoice'!$S$13:$S$23,MATCH(2,'Pencocokan Invoice'!$Q$13:$Q$23,0))&"."&CHAR(10)&"Total selisih yang belum dijelaskan: US$"&TEXT(SUMIF('Pencocokan Invoice'!$O$13:$O$23,"Perlu Dicek",'Pencocokan Invoice'!$K$13:$K$23),"#,##0.00")&". Mohon konfirmasi tindakan di atas, setelah itu saya lanjutkan sisanya."&CHAR(10)&"Salam,"&CHAR(10)&"Latifar"
'@
$wsSum.Range('B10').Formula = @'
="Halo Pak/Bu, update rekonsiliasi invoice: "&COUNTIF('Pencocokan Invoice'!$O$13:$O$23,"Perlu Dicek")&" transaksi perlu ditinjau ("&INDEX('Pencocokan Invoice'!$H$13:$H$23,MATCH(1,'Pencocokan Invoice'!$Q$13:$Q$23,0))&" dan "&INDEX('Pencocokan Invoice'!$H$13:$H$23,MATCH(2,'Pencocokan Invoice'!$Q$13:$Q$23,0))&"), total selisih US$"&TEXT(SUMIF('Pencocokan Invoice'!$O$13:$O$23,"Perlu Dicek",'Pencocokan Invoice'!$K$13:$K$23),"#,##0.00")&". Rinciannya sudah saya kirim di email. Boleh saya lanjut setelah konfirmasi?"
'@
Style-Title $wsSum 'A1' 16
$wsSum.Range('A2').Font.Name='Arial'; $wsSum.Range('A2').Font.Size=10; $wsSum.Range('A2').Font.Italic=$true
foreach ($rr in @(4,9,12)) {
  $c = $wsSum.Range("A$rr")
  $c.Font.Name='Arial'; $c.Font.Size=12; $c.Font.Bold=$true; $c.Font.Color=$CLR_HEADER
}
$wsSum.Range('A5:A7').Font.Name='Arial'; $wsSum.Range('A5:A7').Font.Size=10; $wsSum.Range('A5:A7').Font.Bold=$true
$wsSum.Range('B5:B7').Font.Name='Arial'; $wsSum.Range('B5:B7').Font.Size=10
$wsSum.Range('B7').WrapText = $true
$wsSum.Range('B7').RowHeight = 230
$wsSum.Range('B7').VerticalAlignment = -4160
$wsSum.Range('B10').WrapText = $true
$wsSum.Range('B10').RowHeight = 60
$wsSum.Range('B10').VerticalAlignment = -4160
$wsSum.Range('A13:A17').Font.Name='Arial'; $wsSum.Range('A13:A17').Font.Size=10
$wsSum.Range('A13:A17').WrapText = $true
$wsSum.Range('A13:A17').VerticalAlignment = -4160
$wsSum.Range('A13:A17').RowHeight = 28
Style-Note $wsSum 'A19'
$wsSum.Columns.Item(1).ColumnWidth = 46
$wsSum.Columns.Item(2).ColumnWidth = 92

# ================================================================ cosmetics
foreach ($ws in @($wsPend,$wsData,$wsMat,$wsFlag,$wsSum)) {
  $ws.Cells.Font.Name = 'Arial'
  $ws.Cells.VerticalAlignment = -4160
  $ws.Activate()
  $xl.ActiveWindow.DisplayGridlines = $false
}
$wsPend.Tab.Color = BGR 128 128 128
$wsData.Tab.Color = BGR 31 78 121
$wsMat.Tab.Color  = BGR 0 128 0
$wsFlag.Tab.Color = BGR 192 0 0
$wsSum.Tab.Color  = BGR 112 48 160
try { $wsData.Activate(); $xl.ActiveWindow.FreezePanes = $false; $xl.ActiveWindow.SplitRow = 1; $xl.ActiveWindow.SplitColumn = 0; $xl.ActiveWindow.FreezePanes = $true } catch {}
try { $wsMat.Activate();  $xl.ActiveWindow.FreezePanes = $false; $xl.ActiveWindow.SplitRow = 12; $xl.ActiveWindow.SplitColumn = 0; $xl.ActiveWindow.FreezePanes = $true } catch {}
try { $wsFlag.Activate(); $xl.ActiveWindow.FreezePanes = $false; $xl.ActiveWindow.SplitRow = 3; $xl.ActiveWindow.SplitColumn = 0; $xl.ActiveWindow.FreezePanes = $true } catch {}
foreach ($ws in @($wsPend,$wsData,$wsMat,$wsFlag,$wsSum)) {
  try {
    $ws.PageSetup.Orientation = 2
    $ws.PageSetup.Zoom = $false
    $ws.PageSetup.FitToPagesWide = 1
    $ws.PageSetup.FitToPagesTall = $false
    $ws.PageSetup.LeftMargin = 30; $ws.PageSetup.RightMargin = 30
  } catch {}
}
$wsPend.Activate()

# ================================================================ recalc + save
$xl.CalculateFullRebuild()
if (Test-Path $xlsx) { Remove-Item $xlsx -Force }
$wb.SaveAs($xlsx, 51)
$xl.CalculateFullRebuild()
$wb.Save()

# warn flagged rows (dynamic fill based on computed status)
for ($r=13; $r -le $lastMat; $r++) {
  if ($wsMat.Cells.Item($r,15).Text -eq 'Perlu Dicek') {
    $wsMat.Range("A{0}:S{0}" -f $r).Interior.Color = $CLR_FLAG
    $wsMat.Range("A{0}:S{0}" -f $r).Font.Color = BGR 156 0 6
  } else {
    $wsMat.Range("A{0}:S{0}" -f $r).Interior.Color = $CLR_OK
  }
}
$wb.Save()

# ================================================================ verification readback
$report = [ordered]@{}
$report.file = $xlsx
$report.total_transactions = $wsMat.Range('B6').Value2
$report.cocok = $wsMat.Range('B7').Value2
$report.perlu_dicek = $wsMat.Range('B8').Value2
$report.total_invoice_usd = $wsMat.Range('B9').Value2
$report.total_selisih_usd = $wsMat.Range('B10').Value2
$report.toleransi = $wsMat.Range('B3').Value2
$flagRows = @()
for ($r=4; $r -le $lastFlag; $r++) {
  if ($wsFlag.Cells.Item($r,2).Text) {
    $flagRows += [ordered]@{
      no = $wsFlag.Cells.Item($r,1).Text
      id = $wsFlag.Cells.Item($r,2).Text
      vendor = $wsFlag.Cells.Item($r,3).Text
      invoice = $wsFlag.Cells.Item($r,4).Text
      selisih = $wsFlag.Cells.Item($r,9).Text
      masalah = $wsFlag.Cells.Item($r,10).Text
      tindakan = $wsFlag.Cells.Item($r,11).Text
      pic = $wsFlag.Cells.Item($r,12).Text
      status = $wsFlag.Cells.Item($r,13).Text
    }
  }
}
$report.perlu_dicek_rows = $flagRows
$errors = @()
foreach ($ws in @($wsPend,$wsData,$wsMat,$wsFlag,$wsSum)) {
  $ur = $ws.UsedRange
  for ($rr=1; $rr -le $ur.Rows.Count; $rr++) {
    for ($cc=1; $cc -le $ur.Columns.Count; $cc++) {
      $v = $ur.Cells.Item($rr,$cc).Text
      if ($v -and ($v -like '#*!' -or $v -eq '#N/A' -or $v -like '#VALUE*' -or $v -like '#REF*' -or $v -like '#NAME*' -or $v -like '#DIV*' -or $v -like '#NUM*' -or $v -like '#NULL*')) {
        $errors += ("{0}!R{1}C{2} = {3}" -f $ws.Name, $rr, $cc, $v)
      }
    }
  }
}
$report.formula_errors = $errors
$report.email_preview = $wsSum.Range('B7').Text
$report.sheets = @($wsPend.Name,$wsData.Name,$wsMat.Name,$wsFlag.Name,$wsSum.Name)

# ================================================================ screenshot of Perlu Dicek
try {
  $wsFlag.Activate()
  $range = $wsFlag.Range('A1:M8')
  $range.CopyPicture(1,2) | Out-Null
  $co = $wsFlag.ChartObjects().Add(700, 20, $range.Width, $range.Height)
  $co.Chart.Paste()
  $co.Chart.Export($png,'PNG') | Out-Null
  $co.Delete()
  Write-Output ("screenshot=" + $png)
} catch { Write-Output ("screenshot failed: " + $_.Exception.Message) }

$wb.Save()
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
$report | ConvertTo-Json -Depth 6
