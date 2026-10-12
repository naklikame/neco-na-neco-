Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationCore, WindowsBase

#region KONVERZNI FUNKCE

function Get-SafeTempPath {
    param([string]$Extension)
    $rand = [System.IO.Path]::GetRandomFileName() -replace '\..*$',''
    return Join-Path ([System.IO.Path]::GetTempPath()) ("knv_$rand$Extension")
}

$script:VbsPath = if ($PSScriptRoot) {
    Join-Path $PSScriptRoot 'convert_helper.vbs'
} else {
    Join-Path (Split-Path $MyInvocation.MyCommand.Path -Parent) 'convert_helper.vbs'
}

function Invoke-WordVbs {
    param([string]$Mode, [string]$SrcPath, [string]$DstPath)
    $proc = Start-Process "cscript.exe" `
        -ArgumentList "//NoLogo `"$($script:VbsPath)`" $Mode `"$SrcPath`" `"$DstPath`"" `
        -Wait -PassThru -NoNewWindow
    if ($proc.ExitCode -ne 0) { throw "VBScript selhal (exit $($proc.ExitCode))." }
    if (-not (Test-Path $DstPath)) { throw "PDF nebylo vytvoreno." }
}

function ConvertDocxToPdf {
    param([string]$SrcPath, [string]$DstPath)
    $tempIn  = Get-SafeTempPath '.docx'
    $tempOut = Get-SafeTempPath '.pdf'
    Copy-Item $SrcPath $tempIn -Force
    Unblock-File $tempIn -ErrorAction SilentlyContinue
    try { Invoke-WordVbs 'docx' $tempIn $tempOut }
    finally { Remove-Item $tempIn -ErrorAction SilentlyContinue }
    Move-Item $tempOut $DstPath -Force
}

function ConvertImageToPdf {
    param([string]$SrcPath, [string]$DstPath)
    $srcExt  = [System.IO.Path]::GetExtension($SrcPath)
    $tempIn  = Get-SafeTempPath $srcExt
    $tempOut = Get-SafeTempPath '.pdf'
    Copy-Item $SrcPath $tempIn -Force
    Unblock-File $tempIn -ErrorAction SilentlyContinue
    try { Invoke-WordVbs 'image' $tempIn $tempOut }
    finally { Remove-Item $tempIn -ErrorAction SilentlyContinue }
    Move-Item $tempOut $DstPath -Force
}

$script:ToolDir = Split-Path $script:VbsPath -Parent

# cwebp.exe = zapis WebP, dwebp.exe = cteni WebP (GDI+ WebP neumi)
function Get-WebpTool {
    param([string]$Name)
    $exe = Join-Path $script:ToolDir $Name
    if (Test-Path $exe) { return $exe }

    $ask = [System.Windows.Forms.MessageBox]::Show(
        "Pro WebP je potreba nastroj '$Name' (~500 KB) od Google.`n`nStahnout automaticky?",
        "Stahnout WebP nastroj", 4, 64)
    if ($ask -ne 6) { return $null }

    try {
        $zipUrl  = 'https://storage.googleapis.com/downloads.webmproject.org/releases/webp/libwebp-1.4.0-windows-x64.zip'
        $zipTmp  = [System.IO.Path]::ChangeExtension([System.IO.Path]::GetTempFileName(), '.zip')
        $extTmp  = Join-Path ([System.IO.Path]::GetTempPath()) 'cwebp_dl'
        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile($zipUrl, $zipTmp)
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        if (Test-Path $extTmp) { Remove-Item $extTmp -Recurse -Force }
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipTmp, $extTmp)
        foreach ($n in @('cwebp.exe','dwebp.exe')) {
            $found = Get-ChildItem $extTmp -Recurse -Filter $n | Select-Object -First 1
            if ($found) { Copy-Item $found.FullName (Join-Path $script:ToolDir $n) -Force }
        }
        Remove-Item $zipTmp,$extTmp -Recurse -ErrorAction SilentlyContinue
        if (Test-Path $exe) { return $exe }
        throw "$Name nenalezen v archivu."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Stazeni selhalo: $_`n`nStahnete rucne: https://developers.google.com/speed/webp/download a vlozte $Name vedle skriptu.", "Chyba", 0, 16)|Out-Null
        return $null
    }
}

function Invoke-WebpTool {
    param([string]$Name,[string]$ArgString,[string]$OutPath)
    $exe = Get-WebpTool $Name
    if (-not $exe) { throw "WebP nastroj neni dostupny." }
    $proc = Start-Process $exe -ArgumentList $ArgString -Wait -PassThru -NoNewWindow
    if ($proc.ExitCode -ne 0) { throw "$Name selhal (exit $($proc.ExitCode))." }
    if (-not (Test-Path $OutPath)) { throw "Vystupni soubor nebyl vytvoren." }
}

# Telefony ukladaji pixely nalezato + EXIF znacku orientace (0x0112).
# Otocime pixely podle znacky a znacku smazeme, aby se neotocilo dvakrat.
function Repair-ExifOrientation {
    param([System.Drawing.Image]$Img)
    if ($Img.PropertyIdList -notcontains 0x0112) { return }
    $o = [BitConverter]::ToUInt16($Img.GetPropertyItem(0x0112).Value, 0)
    $rf = switch ($o) {
        2 {'RotateNoneFlipX'}   3 {'Rotate180FlipNone'} 4 {'Rotate180FlipX'}
        5 {'Rotate90FlipX'}     6 {'Rotate90FlipNone'}  7 {'Rotate270FlipX'}
        8 {'Rotate270FlipNone'} default {$null}
    }
    if ($rf) { $Img.RotateFlip([System.Drawing.RotateFlipType]::$rf) }
    $Img.RemovePropertyItem(0x0112)
}

function ConvertImageToWebp {
    param([string]$SrcPath, [string]$DstPath)
    # cwebp EXIF orientaci ignoruje -> pixely otocit pres GDI+ a predat mu PNG
    $tempMid = Get-SafeTempPath '.png'
    $tempOut = Get-SafeTempPath '.webp'
    try {
        $bmp = [System.Drawing.Bitmap]::new($SrcPath)
        try {
            Repair-ExifOrientation $bmp
            $bmp.Save($tempMid,[System.Drawing.Imaging.ImageFormat]::Png)
        } finally { $bmp.Dispose() }
        Invoke-WebpTool 'cwebp.exe' "-quiet -q 90 `"$tempMid`" -o `"$tempOut`"" $tempOut
        Move-Item $tempOut $DstPath -Force
    } finally { Remove-Item $tempMid,$tempOut -ErrorAction SilentlyContinue }
}

# HEIC/HEIF: GDI+ neumi, cte se pres WIC (vyzaduje "HEIF Image Extensions" z MS Store).
# Dekoder uz aplikuje otoceni, vystupni PNG je tedy spravne orientovane.
function ConvertHeicToPng {
    param([string]$SrcPath,[string]$DstPath)
    $fs = [System.IO.File]::OpenRead($SrcPath)
    try {
        try {
            $dec = [System.Windows.Media.Imaging.BitmapDecoder]::Create($fs,
                [System.Windows.Media.Imaging.BitmapCreateOptions]::None,
                [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
        } catch {
            throw "HEIC nelze precist - nainstalujte z Microsoft Store 'HEIF Image Extensions' a 'HEVC Video Extensions'."
        }
        $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
        $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($dec.Frames[0], $null, $null, $null))
        $out = [System.IO.File]::Create($DstPath)
        try { $enc.Save($out) } finally { $out.Close() }
    } finally { $fs.Close() }
}

function CompressImage {
    param([string]$SrcPath,[string]$DstPath,[int]$Quality,[int]$MaxW,[int]$MaxH)
    $srcExt = [System.IO.Path]::GetExtension($SrcPath).ToLower()
    $dstExt = [System.IO.Path]::GetExtension($DstPath).ToLower()
    $temps  = @()
    try {
        # WebP vstup: nejdriv dekodovat do docasneho PNG
        $loadPath = $SrcPath
        if ($srcExt -eq '.webp') {
            $tIn = Get-SafeTempPath '.webp'; $tPng = Get-SafeTempPath '.png'
            $temps += $tIn,$tPng
            Copy-Item $SrcPath $tIn -Force
            Invoke-WebpTool 'dwebp.exe' "-quiet `"$tIn`" -o `"$tPng`"" $tPng
            $loadPath = $tPng
        } elseif ($srcExt -in @('.heic','.heif')) {
            $tPng = Get-SafeTempPath '.png'
            $temps += $tPng
            ConvertHeicToPng -SrcPath $SrcPath -DstPath $tPng
            $loadPath = $tPng
        }
        $bmp = [System.Drawing.Bitmap]::new($loadPath)
        try {
            Repair-ExifOrientation $bmp
            $w = $bmp.Width; $h = $bmp.Height
            if (($MaxW -gt 0 -and $w -gt $MaxW) -or ($MaxH -gt 0 -and $h -gt $MaxH)) {
                $rW = if ($MaxW -gt 0) { $MaxW/$w } else { 1.0 }
                $rH = if ($MaxH -gt 0) { $MaxH/$h } else { 1.0 }
                $r  = [Math]::Min($rW,$rH)
                $nW = [int]($w*$r); $nH = [int]($h*$r)
                $rsz = New-Object System.Drawing.Bitmap($nW,$nH)
                $g   = [System.Drawing.Graphics]::FromImage($rsz)
                $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $g.DrawImage($bmp,0,0,$nW,$nH); $g.Dispose(); $bmp.Dispose(); $bmp = $rsz
            }
            if ($dstExt -in @('.jpg','.jpeg')) {
                $jpgEnc = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() |
                          Where-Object { $_.MimeType -eq 'image/jpeg' } | Select-Object -First 1
                $ep = New-Object System.Drawing.Imaging.EncoderParameters(1)
                $ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter(
                    [System.Drawing.Imaging.Encoder]::Quality,[long]$Quality)
                $bmp.Save($DstPath,$jpgEnc,$ep)
            } elseif ($dstExt -eq '.webp') {
                # WebP vystup: bezeztratove PNG -> cwebp s danou kvalitou
                $tMid = Get-SafeTempPath '.png'; $tOut = Get-SafeTempPath '.webp'
                $temps += $tMid,$tOut
                $bmp.Save($tMid,[System.Drawing.Imaging.ImageFormat]::Png)
                Invoke-WebpTool 'cwebp.exe' "-quiet -q $Quality `"$tMid`" -o `"$tOut`"" $tOut
                Move-Item $tOut $DstPath -Force
            } else { $bmp.Save($DstPath,[System.Drawing.Imaging.ImageFormat]::Png) }
        } finally { $bmp.Dispose() }
    } finally {
        foreach ($t in $temps) { Remove-Item $t -ErrorAction SilentlyContinue }
    }
}

function Get-RenamedName {
    param([string]$OldName,[string]$NewName,[string]$FindTxt,[string]$ReplaceTxt,
          [string]$Prefix,[string]$Suffix,[bool]$UseNum,[int]$Num,[int]$Pad)
    $ext = [System.IO.Path]::GetExtension($OldName)
    if ($NewName) {
        $base = $NewName
    } else {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($OldName)
        if ($FindTxt) { $base = $base.Replace($FindTxt,$ReplaceTxt) }
    }
    $numStr = if ($UseNum) { $Num.ToString().PadLeft($Pad,'0') } else { '' }
    return "$Prefix$base$Suffix$numStr$ext"
}

function Invoke-WordVbsEx {
    param([string]$ArgString)
    $proc = Start-Process "cscript.exe" `
        -ArgumentList "//NoLogo `"$($script:VbsPath)`" $ArgString" `
        -Wait -PassThru -NoNewWindow
    if ($proc.ExitCode -ne 0) { throw "VBScript selhal (exit $($proc.ExitCode))." }
}

function MergePdfs {
    param([string[]]$Inputs, [string]$DstPath)
    $temps = @()
    foreach ($inp in $Inputs) {
        $t = Get-SafeTempPath '.pdf'
        Copy-Item $inp $t -Force
        Unblock-File $t -ErrorAction SilentlyContinue
        $temps += $t
    }
    $tempOut = Get-SafeTempPath '.pdf'
    try {
        $joined = $temps -join '|'
        Invoke-WordVbsEx "merge `"$joined`" `"$tempOut`""
        if (-not (Test-Path $tempOut)) { throw "PDF nebylo vytvoreno." }
    } finally {
        foreach ($t in $temps) { Remove-Item $t -ErrorAction SilentlyContinue }
    }
    Move-Item $tempOut $DstPath -Force
}

function ExtractPdfPages {
    param([string]$SrcPath, [string]$DstPath, [int]$FromPage, [int]$ToPage)
    $tempIn  = Get-SafeTempPath '.pdf'
    $tempOut = Get-SafeTempPath '.pdf'
    Copy-Item $SrcPath $tempIn -Force
    Unblock-File $tempIn -ErrorAction SilentlyContinue
    try {
        Invoke-WordVbsEx "extract `"$tempIn`" `"$tempOut`" $FromPage $ToPage"
        if (-not (Test-Path $tempOut)) { throw "PDF nebylo vytvoreno." }
    } finally {
        Remove-Item $tempIn -ErrorAction SilentlyContinue
    }
    Move-Item $tempOut $DstPath -Force
}

#endregion

#region BARVY + FONTY + HELPERS

$C_NAVY   = [System.Drawing.Color]::FromArgb(15,23,42)
$C_ACCENT = [System.Drawing.Color]::FromArgb(37,99,235)
$C_GREEN  = [System.Drawing.Color]::FromArgb(22,163,74)
$C_RED    = [System.Drawing.Color]::FromArgb(220,38,38)
$C_ORANGE = [System.Drawing.Color]::FromArgb(217,119,6)
$C_BG     = [System.Drawing.Color]::White
$C_BGALT  = [System.Drawing.Color]::FromArgb(241,245,249)
$C_BORDER = [System.Drawing.Color]::FromArgb(203,213,225)
$C_TEXT   = [System.Drawing.Color]::FromArgb(15,23,42)
$C_SUB    = [System.Drawing.Color]::FromArgb(100,116,139)

$FNT_N = New-Object System.Drawing.Font('Segoe UI',9)
$FNT_B = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
$FNT_H = New-Object System.Drawing.Font('Segoe UI',12,[System.Drawing.FontStyle]::Bold)
$FNT_S = New-Object System.Drawing.Font('Segoe UI',8)
$FNT_M = New-Object System.Drawing.Font('Consolas',8)

function New-Lbl {
    param([string]$T,[int]$X,[int]$Y,[bool]$B=$false,$Clr=$null)
    $l = New-Object System.Windows.Forms.Label
    $l.Text=$T; $l.Location=New-Object System.Drawing.Point($X,$Y); $l.AutoSize=$true
    $l.Font=if($B){$FNT_B}else{$FNT_N}; $l.ForeColor=if($Clr){$Clr}else{$C_TEXT}
    return $l
}

function New-Btn {
    param([string]$T,[int]$X,[int]$Y,[int]$W=130,[int]$H=28,[bool]$P=$false)
    $b = New-Object System.Windows.Forms.Button
    $b.Text=$T; $b.Location=New-Object System.Drawing.Point($X,$Y)
    $b.Size=New-Object System.Drawing.Size($W,$H); $b.FlatStyle='Flat'
    if ($P) {
        $b.BackColor=$C_ACCENT; $b.ForeColor=[System.Drawing.Color]::White; $b.Font=$FNT_B
        $b.FlatAppearance.BorderColor=$C_ACCENT
    } else {
        $b.BackColor=$C_BG; $b.ForeColor=$C_SUB; $b.Font=$FNT_N
        $b.FlatAppearance.BorderColor=$C_BORDER
    }
    return $b
}

function New-Txt {
    param([int]$X,[int]$Y,[int]$W,[string]$Val='',$Clr=$null)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location=New-Object System.Drawing.Point($X,$Y); $t.Size=New-Object System.Drawing.Size($W,24)
    $t.BorderStyle='FixedSingle'; $t.Font=$FNT_N
    if ($Val){ $t.Text=$Val }
    if ($Clr){ $t.ForeColor=$Clr }
    return $t
}

function New-FileList {
    param([int]$Y,[int]$H=100)
    $lb = New-Object System.Windows.Forms.ListBox
    $lb.Location=New-Object System.Drawing.Point(14,$Y); $lb.Size=New-Object System.Drawing.Size(554,$H)
    $lb.BorderStyle='FixedSingle'; $lb.HorizontalScrollbar=$true
    $lb.BackColor=$C_BG; $lb.ForeColor=$C_TEXT; $lb.Font=$FNT_S; $lb.AllowDrop=$true
    return $lb
}

function New-Log {
    param([int]$Y,[int]$X=14,[int]$W=554,[int]$H=54)
    $rt = New-Object System.Windows.Forms.RichTextBox
    $rt.Location=New-Object System.Drawing.Point($X,$Y); $rt.Size=New-Object System.Drawing.Size($W,$H)
    $rt.ReadOnly=$true; $rt.BorderStyle='None'; $rt.BackColor=$C_BGALT
    $rt.ForeColor=$C_SUB; $rt.Font=$FNT_M
    return $rt
}

function New-Sep { param([int]$Y)
    $p = New-Object System.Windows.Forms.Panel
    $p.Location=New-Object System.Drawing.Point(14,$Y); $p.Size=New-Object System.Drawing.Size(554,1)
    $p.BackColor=$C_BORDER; return $p
}

function Add-SecHdr {
    param($Pnl,[string]$T,[int]$Y)
    $bar = New-Object System.Windows.Forms.Panel
    $bar.Location=New-Object System.Drawing.Point(14,$Y); $bar.Size=New-Object System.Drawing.Size(4,16)
    $bar.BackColor=$C_ACCENT
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text=$T; $lbl.Location=New-Object System.Drawing.Point(24,($Y-2))
    $lbl.AutoSize=$true; $lbl.Font=$FNT_B; $lbl.ForeColor=$C_TEXT
    $Pnl.Controls.Add($bar); $Pnl.Controls.Add($lbl)
}

function AppendLog {
    param($R,[string]$M,$Clr=$null)
    $R.SelectionStart=$R.TextLength; $R.SelectionLength=0
    if ($Clr){$R.SelectionColor=$Clr}
    $R.AppendText("$M`n"); $R.SelectionColor=$C_SUB; $R.ScrollToCaret(); $form.Refresh()
}

function Add-ToList {
    param($Lst,[string[]]$Paths,[string[]]$Exts)
    foreach ($p in $Paths) {
        $e = [System.IO.Path]::GetExtension($p).ToLower()
        if ($Exts.Count -eq 0 -or $e -in $Exts) {
            if (-not $Lst.Items.Contains($p)) { $Lst.Items.Add($p)|Out-Null }
        }
    }
}

function Register-DD {
    param($Lst,[string[]]$Exts)
    $Lst.Add_DragEnter({
        $_.Effect = if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {'Copy'} else {'None'}
    }.GetNewClosure())
    $Lst.Add_DragDrop({
        Add-ToList $Lst $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop) $Exts
    }.GetNewClosure())
}

function Open-Folder { param([string]$D); Start-Process explorer.exe $D }

#endregion

#region FORM + NAVIGACE

$form = New-Object System.Windows.Forms.Form
$form.Text='Nastroje pro soubory'; $form.Font=$FNT_N
$form.Size=New-Object System.Drawing.Size(582,700)
$form.MinimumSize=New-Object System.Drawing.Size(582,700)
$form.StartPosition='CenterScreen'; $form.FormBorderStyle='FixedDialog'
$form.MaximizeBox=$false; $form.BackColor=$C_BG

# --- Header ---
$pnlHdr = New-Object System.Windows.Forms.Panel
$pnlHdr.Location=New-Object System.Drawing.Point(0,0); $pnlHdr.Size=New-Object System.Drawing.Size(582,52)
$pnlHdr.BackColor=$C_NAVY
$lblHdr = New-Object System.Windows.Forms.Label
$lblHdr.Text='  Nastroje pro soubory'; $lblHdr.Dock='Fill'
$lblHdr.ForeColor=[System.Drawing.Color]::White; $lblHdr.Font=$FNT_H; $lblHdr.TextAlign='MiddleLeft'
$pnlHdr.Controls.Add($lblHdr)

# --- Nav bar ---
$pnlNav = New-Object System.Windows.Forms.Panel
$pnlNav.Location=New-Object System.Drawing.Point(0,52); $pnlNav.Size=New-Object System.Drawing.Size(582,44)
$pnlNav.BackColor=$C_BGALT

$sepNav = New-Object System.Windows.Forms.Panel
$sepNav.Location=New-Object System.Drawing.Point(0,95); $sepNav.Size=New-Object System.Drawing.Size(582,1)
$sepNav.BackColor=$C_BORDER

function New-NavBtn {
    param([string]$T,[int]$X)
    $b = New-Object System.Windows.Forms.Button
    $b.Text=$T; $b.Location=New-Object System.Drawing.Point($X,5)
    $b.Size=New-Object System.Drawing.Size(136,34); $b.FlatStyle='Flat'
    $b.FlatAppearance.BorderSize=0; $b.BackColor=$C_BGALT
    $b.ForeColor=$C_SUB; $b.Font=$FNT_N
    $b.Cursor=[System.Windows.Forms.Cursors]::Hand
    return $b
}

$navC = New-NavBtn 'Konvertor' 8
$navR = New-NavBtn 'Prejmenovani' 148
$navK = New-NavBtn 'Komprese' 288
$navP = New-NavBtn 'PDF operace' 428
$pnlNav.Controls.AddRange(@($navC,$navR,$navK,$navP))

# --- Content panels ---
$pnlC = New-Object System.Windows.Forms.Panel
$pnlC.Location=New-Object System.Drawing.Point(0,96); $pnlC.Size=New-Object System.Drawing.Size(582,604)
$pnlC.BackColor=$C_BG; $pnlC.Visible=$true

$pnlR = New-Object System.Windows.Forms.Panel
$pnlR.Location=New-Object System.Drawing.Point(0,96); $pnlR.Size=New-Object System.Drawing.Size(582,604)
$pnlR.BackColor=$C_BG; $pnlR.Visible=$false

$pnlK = New-Object System.Windows.Forms.Panel
$pnlK.Location=New-Object System.Drawing.Point(0,96); $pnlK.Size=New-Object System.Drawing.Size(582,604)
$pnlK.BackColor=$C_BG; $pnlK.Visible=$false

$pnlP = New-Object System.Windows.Forms.Panel
$pnlP.Location=New-Object System.Drawing.Point(0,96); $pnlP.Size=New-Object System.Drawing.Size(582,604)
$pnlP.BackColor=$C_BG; $pnlP.Visible=$false

$form.Controls.AddRange(@($pnlHdr,$pnlNav,$sepNav,$pnlC,$pnlR,$pnlK,$pnlP))

$script:navBtns = @($navC,$navR,$navK,$navP)
$script:cntPnls = @($pnlC,$pnlR,$pnlK,$pnlP)

function Set-ActiveNav {
    param([int]$I)
    for ($j=0;$j -lt 4;$j++) {
        $script:cntPnls[$j].Visible=($j -eq $I)
        if ($j -eq $I) {
            $script:navBtns[$j].BackColor=[System.Drawing.Color]::FromArgb(219,234,254)
            $script:navBtns[$j].ForeColor=[System.Drawing.Color]::FromArgb(29,78,216)
            $script:navBtns[$j].Font=$FNT_B
        } else {
            $script:navBtns[$j].BackColor=$C_BGALT
            $script:navBtns[$j].ForeColor=$C_SUB
            $script:navBtns[$j].Font=$FNT_N
        }
    }
}
$navC.Add_Click({Set-ActiveNav 0})
$navR.Add_Click({Set-ActiveNav 1})
$navK.Add_Click({Set-ActiveNav 2})
$navP.Add_Click({Set-ActiveNav 3})
Set-ActiveNav 0

#endregion

#region PANEL: KONVERTOR

Add-SecHdr $pnlC 'Soubory ke konverzi' 16
$cLst = New-FileList -Y 36
$cBtnBrowse = New-Btn 'Vybrat soubory...' 14 144
$cBtnClear  = New-Btn 'Vymazat' 152 144 80
$pnlC.Controls.AddRange(@($cLst,$cBtnBrowse,$cBtnClear))
$pnlC.Controls.Add((New-Sep 178))

Add-SecHdr $pnlC 'Nastaveni' 188
$pnlC.Controls.Add((New-Lbl 'Cilovy format:' 14 210))

$cCmb = New-Object System.Windows.Forms.ComboBox
$cCmb.Location=New-Object System.Drawing.Point(112,207); $cCmb.Size=New-Object System.Drawing.Size(180,26)
$cCmb.DropDownStyle='DropDownList'; $cCmb.FlatStyle='Flat'; $cCmb.Font=$FNT_N
$cCmb.Items.Add('PDF  (.pdf)')|Out-Null
$cCmb.Items.Add('WebP  (.webp)')|Out-Null
$cCmb.SelectedIndex=0

$pnlC.Controls.Add($cCmb)
$pnlC.Controls.Add((New-Lbl 'Vystupni slozka:' 14 243))
$cTxtOut = New-Txt 14 261 486 '(stejna slozka jako zdrojovy soubor)' $C_SUB
$cBtnOut = New-Btn '...' 506 259 46 26
$pnlC.Controls.AddRange(@($cTxtOut,$cBtnOut))
$pnlC.Controls.Add((New-Sep 294))

$cBtnConv = New-Btn 'PREVEST' 14 304 120 34 $true
$cProg = New-Object System.Windows.Forms.ProgressBar
$cProg.Location=New-Object System.Drawing.Point(144,310); $cProg.Size=New-Object System.Drawing.Size(424,22); $cProg.Style='Continuous'
$cLog = New-Log -Y 348
$pnlC.Controls.AddRange(@($cBtnConv,$cProg,$cLog))

#endregion

#region PANEL: PREJMENOVANI

Add-SecHdr $pnlR 'Soubory k prejmenovani' 16
$rLst = New-FileList -Y 36
$rBtnBrowse = New-Btn 'Vybrat soubory...' 14 144
$rBtnClear  = New-Btn 'Vymazat' 152 144 80
$pnlR.Controls.AddRange(@($rLst,$rBtnBrowse,$rBtnClear))
$pnlR.Controls.Add((New-Sep 178))

Add-SecHdr $pnlR 'Transformace nazvu' 188
$pnlR.Controls.Add((New-Lbl 'Novy nazev:' 14 210))
$rTxtNew = New-Txt 106 207 462; $rTxtNew.Font=$FNT_B
$rLblHint = New-Lbl '(prazdne = zachovat puvodni nazev)' 106 233 $false $C_SUB; $rLblHint.Font=$FNT_S
$pnlR.Controls.AddRange(@($rTxtNew,$rLblHint))

$pnlR.Controls.Add((New-Lbl 'Najit:' 14 257))
$rTxtFind = New-Txt 60 254 180
$pnlR.Controls.Add((New-Lbl 'Nahradit:' 252 257))
$rTxtRepl = New-Txt 318 254 250
$pnlR.Controls.AddRange(@($rTxtFind,$rTxtRepl))

$pnlR.Controls.Add((New-Lbl 'Prefix:' 14 290))
$rTxtPre = New-Txt 60 287 180
$pnlR.Controls.Add((New-Lbl 'Suffix:' 252 290))
$rTxtSuf = New-Txt 318 287 250
$pnlR.Controls.AddRange(@($rTxtPre,$rTxtSuf))

$rChkNum = New-Object System.Windows.Forms.CheckBox
$rChkNum.Text='Cislovani'; $rChkNum.Location=New-Object System.Drawing.Point(14,320); $rChkNum.AutoSize=$true; $rChkNum.Font=$FNT_N
$pnlR.Controls.Add((New-Lbl 'Start:' 116 322))
$rNumStart = New-Object System.Windows.Forms.NumericUpDown
$rNumStart.Location=New-Object System.Drawing.Point(156,319); $rNumStart.Size=New-Object System.Drawing.Size(58,24)
$rNumStart.Minimum=0; $rNumStart.Maximum=9999; $rNumStart.Value=1; $rNumStart.Enabled=$false; $rNumStart.Font=$FNT_N
$pnlR.Controls.Add((New-Lbl 'Cislic:' 228 322))
$rNumPad = New-Object System.Windows.Forms.NumericUpDown
$rNumPad.Location=New-Object System.Drawing.Point(270,319); $rNumPad.Size=New-Object System.Drawing.Size(46,24)
$rNumPad.Minimum=1; $rNumPad.Maximum=6; $rNumPad.Value=2; $rNumPad.Enabled=$false; $rNumPad.Font=$FNT_N
$pnlR.Controls.AddRange(@($rChkNum,$rNumStart,$rNumPad))
$pnlR.Controls.Add((New-Sep 350))

Add-SecHdr $pnlR 'Nahled' 360
$rLV = New-Object System.Windows.Forms.ListView
$rLV.Location=New-Object System.Drawing.Point(14,380); $rLV.Size=New-Object System.Drawing.Size(554,118)
$rLV.View='Details'; $rLV.FullRowSelect=$true; $rLV.GridLines=$true; $rLV.BorderStyle='FixedSingle'; $rLV.Font=$FNT_S
$cOld=New-Object System.Windows.Forms.ColumnHeader; $cOld.Text='Puvodni nazev'; $cOld.Width=266
$cNew=New-Object System.Windows.Forms.ColumnHeader; $cNew.Text='Novy nazev'; $cNew.Width=284
$rLV.Columns.AddRange(@($cOld,$cNew))
$pnlR.Controls.Add($rLV)
$pnlR.Controls.Add((New-Sep 506))

$rBtnRen = New-Btn 'PREJMENOVANI' 14 516 148 34 $true
$rLog = New-Log -Y 520 -X 172 -W 396 -H 44
$pnlR.Controls.AddRange(@($rBtnRen,$rLog))

#endregion

#region PANEL: KOMPRESE

Add-SecHdr $pnlK 'Obrazky ke kompresi  (jpg, png, webp, heic)' 16
$kLst = New-FileList -Y 36
$kBtnBrowse = New-Btn 'Vybrat obrazky...' 14 144
$kBtnClear  = New-Btn 'Vymazat' 152 144 80
$pnlK.Controls.AddRange(@($kLst,$kBtnBrowse,$kBtnClear))
$pnlK.Controls.Add((New-Sep 178))

Add-SecHdr $pnlK 'Moznosti komprese' 188
$pnlK.Controls.Add((New-Lbl 'Kvalita:' 14 210 $true))
$kTrack = New-Object System.Windows.Forms.TrackBar
$kTrack.Location=New-Object System.Drawing.Point(110,206); $kTrack.Size=New-Object System.Drawing.Size(380,32)
$kTrack.Minimum=10; $kTrack.Maximum=100; $kTrack.Value=70; $kTrack.TickFrequency=10
$kLblQ = New-Lbl '70 %' 500 210 $true
$pnlK.Controls.AddRange(@($kTrack,$kLblQ))

$kChkRes = New-Object System.Windows.Forms.CheckBox
$kChkRes.Text='Max rozliseni:'; $kChkRes.Location=New-Object System.Drawing.Point(14,248); $kChkRes.AutoSize=$true; $kChkRes.Font=$FNT_N
$kTxtW = New-Txt 142 245 62 '1920'; $kTxtW.Enabled=$false
$pnlK.Controls.Add((New-Lbl 'x' 210 248))
$kTxtH = New-Txt 222 245 62 '1080'; $kTxtH.Enabled=$false
$pnlK.Controls.Add((New-Lbl 'px' 290 248))
$pnlK.Controls.AddRange(@($kChkRes,$kTxtW,$kTxtH))

$pnlK.Controls.Add((New-Lbl 'Ulozit jako:' 14 278))
$kRbSame = New-Object System.Windows.Forms.RadioButton; $kRbSame.Text='Stejny format'; $kRbSame.Location=New-Object System.Drawing.Point(112,276); $kRbSame.AutoSize=$true; $kRbSame.Checked=$true; $kRbSame.Font=$FNT_N
$kRbJpg  = New-Object System.Windows.Forms.RadioButton; $kRbJpg.Text='JPG'; $kRbJpg.Location=New-Object System.Drawing.Point(244,276); $kRbJpg.AutoSize=$true; $kRbJpg.Font=$FNT_N
$kRbPng  = New-Object System.Windows.Forms.RadioButton; $kRbPng.Text='PNG'; $kRbPng.Location=New-Object System.Drawing.Point(300,276); $kRbPng.AutoSize=$true; $kRbPng.Font=$FNT_N
$kRbWebp = New-Object System.Windows.Forms.RadioButton; $kRbWebp.Text='WebP'; $kRbWebp.Location=New-Object System.Drawing.Point(356,276); $kRbWebp.AutoSize=$true; $kRbWebp.Font=$FNT_N
$pnlK.Controls.AddRange(@($kRbSame,$kRbJpg,$kRbPng,$kRbWebp))

$pnlK.Controls.Add((New-Lbl 'Vystupni slozka:' 14 308))
$kTxtOut = New-Txt 14 326 490 '(prepsat originaly)' $C_SUB
$kBtnOut = New-Btn '...' 510 324 46 26
$pnlK.Controls.AddRange(@($kTxtOut,$kBtnOut))

$kChkSfx = New-Object System.Windows.Forms.CheckBox
$kChkSfx.Text="Pridat '_k' do nazvu"; $kChkSfx.Location=New-Object System.Drawing.Point(14,358); $kChkSfx.AutoSize=$true; $kChkSfx.Checked=$true; $kChkSfx.Font=$FNT_N
$pnlK.Controls.Add($kChkSfx)
$pnlK.Controls.Add((New-Sep 382))

$kBtnComp = New-Btn 'KOMPRIMOVAT' 14 392 148 34 $true
$kProg = New-Object System.Windows.Forms.ProgressBar
$kProg.Location=New-Object System.Drawing.Point(172,398); $kProg.Size=New-Object System.Drawing.Size(396,22); $kProg.Style='Continuous'
$kLog = New-Log -Y 436
$pnlK.Controls.AddRange(@($kBtnComp,$kProg,$kLog))

#endregion

#region PANEL: PDF

# ---- Sekce 1: Slucovani PDF ----
Add-SecHdr $pnlP 'Slucovani PDF do jednoho souboru' 16
$pMergeLst = New-FileList -Y 36 -H 90
$pnlP.Controls.Add($pMergeLst)

$pMergeBtnBrowse = New-Btn 'Vybrat PDF...' 14 134 120 28
$pMergeBtnClear  = New-Btn 'Vymazat' 142 134 70 28
$pMergeBtnUp     = New-Btn '^ Nahoru' 220 134 80 28
$pMergeBtnDown   = New-Btn 'v Dolu' 308 134 70 28
$pnlP.Controls.AddRange(@($pMergeBtnBrowse,$pMergeBtnClear,$pMergeBtnUp,$pMergeBtnDown))

$pnlP.Controls.Add((New-Lbl 'Vystupni slozka:' 14 170))
$pMergeTxtOut = New-Txt 14 188 490 '(slozka prvniho PDF ze seznamu)' $C_SUB
$pMergeBtnOut = New-Btn '...' 510 186 46 26
$pnlP.Controls.AddRange(@($pMergeTxtOut,$pMergeBtnOut))
$pnlP.Controls.Add((New-Sep 220))

$pMergeBtnRun = New-Btn 'SLOUCIT PDF' 14 228 148 34 $true
$pnlP.Controls.Add($pMergeBtnRun)

# ---- Vizualni oddelovac sekci ----
$pDivider = New-Object System.Windows.Forms.Panel
$pDivider.Location=New-Object System.Drawing.Point(0,278); $pDivider.Size=New-Object System.Drawing.Size(582,3)
$pDivider.BackColor=$C_BORDER
$pnlP.Controls.Add($pDivider)

# ---- Sekce 2: Vyriz stranky ----
Add-SecHdr $pnlP 'Vyriz stranky z PDF' 292

$pnlP.Controls.Add((New-Lbl 'Vstupni PDF:' 14 318))
$pExtractTxtSrc = New-Txt 14 336 490 '' $C_TEXT
$pExtractBtnSrc = New-Btn 'Vybrat...' 510 334 58 26
$pnlP.Controls.AddRange(@($pExtractTxtSrc,$pExtractBtnSrc))

$pnlP.Controls.Add((New-Lbl 'Stranky:  od' 14 373))
$pExtractNumFrom = New-Object System.Windows.Forms.NumericUpDown
$pExtractNumFrom.Location=New-Object System.Drawing.Point(96,370); $pExtractNumFrom.Size=New-Object System.Drawing.Size(58,26)
$pExtractNumFrom.Minimum=1; $pExtractNumFrom.Maximum=9999; $pExtractNumFrom.Value=1; $pExtractNumFrom.Font=$FNT_N
$pnlP.Controls.Add((New-Lbl 'do' 162 373))
$pExtractNumTo = New-Object System.Windows.Forms.NumericUpDown
$pExtractNumTo.Location=New-Object System.Drawing.Point(186,370); $pExtractNumTo.Size=New-Object System.Drawing.Size(58,26)
$pExtractNumTo.Minimum=1; $pExtractNumTo.Maximum=9999; $pExtractNumTo.Value=1; $pExtractNumTo.Font=$FNT_N
$pnlP.Controls.AddRange(@($pExtractNumFrom,$pExtractNumTo))

$pnlP.Controls.Add((New-Lbl 'Vystupni slozka:' 14 406))
$pExtractTxtOut = New-Txt 14 424 490 '(stejna slozka jako zdrojovy PDF)' $C_SUB
$pExtractBtnOut = New-Btn '...' 510 422 46 26
$pnlP.Controls.AddRange(@($pExtractTxtOut,$pExtractBtnOut))
$pnlP.Controls.Add((New-Sep 456))

$pExtractBtnRun = New-Btn 'VYREZAT STRANKY' 14 464 160 34 $true
$pnlP.Controls.Add($pExtractBtnRun)

# ---- Spolecny log ----
$pLog = New-Log -Y 508 -H 60
$pnlP.Controls.Add($pLog)

$pHint = New-Lbl 'Poznamka: Word pri zpracovani PDF prevadi obsah do editovatelneho formatu - slozite layouty nemusi byt 100% presne.' 14 578 $false $C_SUB
$pHint.Font=$FNT_S; $pHint.AutoSize=$false; $pHint.Size=New-Object System.Drawing.Size(554,24)
$pnlP.Controls.Add($pHint)

#endregion

#region LOGIKA: KONVERTOR

$script:cCustomOut = $null

$cBtnBrowse.Add_Click({
    $ofd=New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Multiselect=$true
    $ofd.Filter='Podporovane|*.docx;*.jpg;*.jpeg;*.png|Word|*.docx|Obrazky|*.jpg;*.jpeg;*.png|Vse|*.*'
    if ($ofd.ShowDialog($form) -eq 'OK') {
        Add-ToList $cLst $ofd.FileNames @('.docx','.jpg','.jpeg','.png')
    }
})
$cBtnClear.Add_Click({ $cLst.Items.Clear(); $cLog.Clear(); $cProg.Value=0 })
Register-DD $cLst @('.docx','.jpg','.jpeg','.png')

$cBtnOut.Add_Click({
    $fbd=New-Object System.Windows.Forms.FolderBrowserDialog
    if ($fbd.ShowDialog($form) -eq 'OK') {
        $script:cCustomOut=$fbd.SelectedPath; $cTxtOut.Text=$fbd.SelectedPath; $cTxtOut.ForeColor=$C_TEXT
    }
})

$cBtnConv.Add_Click({
    if ($cLst.Items.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Vyberte soubory.','',0,48)|Out-Null; return }
    $fmt   = $cCmb.SelectedItem
    $files = @($cLst.Items)
    $cProg.Maximum=$files.Count; $cProg.Value=0; $cLog.Clear(); $cBtnConv.Enabled=$false
    $ok=0; $fail=0
    foreach ($src in $files) {
        $sName = Split-Path $src -Leaf
        $sExt  = [System.IO.Path]::GetExtension($src).ToLower()
        $oDir  = if ($script:cCustomOut) {$script:cCustomOut} else {Split-Path $src -Parent}
        $base  = [System.IO.Path]::GetFileNameWithoutExtension($src)
        $dExt  = if ($fmt -like '*webp*') {'.webp'} else {'.pdf'}
        if ($dExt -eq '.webp' -and $sExt -eq '.docx') {
            AppendLog $cLog "Preskoceno [$sName] - Word na WebP neni mozny" $C_ORANGE
            $cProg.Value++; continue
        }
        $valid = ($dExt -eq '.pdf' -and $sExt -in @('.docx','.jpg','.jpeg','.png')) -or
                 ($dExt -eq '.webp' -and $sExt -in @('.jpg','.jpeg','.png'))
        if (-not $valid) { AppendLog $cLog "Preskoceno [$sName]" $C_ORANGE; $cProg.Value++; continue }
        $dst = Join-Path $oDir ($base+$dExt)
        AppendLog $cLog "Probiha: $sName ..." $C_SUB
        try {
            if ($dExt -eq '.pdf' -and $sExt -eq '.docx') { ConvertDocxToPdf -SrcPath $src -DstPath $dst }
            elseif ($dExt -eq '.pdf')                     { ConvertImageToPdf -SrcPath $src -DstPath $dst }
            else                                           { ConvertImageToWebp -SrcPath $src -DstPath $dst }
            AppendLog $cLog "OK: $sName -> $(Split-Path $dst -Leaf)" $C_GREEN; $ok++
        } catch { AppendLog $cLog "CHYBA [$sName]: $_" $C_RED; $fail++ }
        $cProg.Value++
    }
    $cBtnConv.Enabled=$true
    $sum="Hotovo: $ok uspesne"; if($fail -gt 0){$sum+=", $fail chyb"}
    AppendLog $cLog $sum $C_TEXT
    if ($ok -gt 0) {
        $dir = if($script:cCustomOut){$script:cCustomOut}else{Split-Path $files[0] -Parent}
        if ([System.Windows.Forms.MessageBox]::Show("$sum`nOtevrit slozku?","Hotovo",4,64) -eq 6) { Open-Folder $dir }
    }
})

#endregion

#region LOGIKA: PREJMENOVANI

$rBtnBrowse.Add_Click({
    $ofd=New-Object System.Windows.Forms.OpenFileDialog; $ofd.Multiselect=$true; $ofd.Filter='Vsechny soubory|*.*'
    if ($ofd.ShowDialog($form) -eq 'OK') { Add-ToList $rLst $ofd.FileNames @(); UpdateRenamePreview }
})
$rBtnClear.Add_Click({ $rLst.Items.Clear(); $rLV.Items.Clear(); $rLog.Clear() })
Register-DD $rLst @()
$rLst.Add_DragDrop({ UpdateRenamePreview })

$rChkNum.Add_CheckedChanged({ $rNumStart.Enabled=$rChkNum.Checked; $rNumPad.Enabled=$rChkNum.Checked; UpdateRenamePreview })
$rTxtNew.Add_TextChanged({UpdateRenamePreview})
$rTxtFind.Add_TextChanged({UpdateRenamePreview})
$rTxtRepl.Add_TextChanged({UpdateRenamePreview})
$rTxtPre.Add_TextChanged({UpdateRenamePreview})
$rTxtSuf.Add_TextChanged({UpdateRenamePreview})
$rNumStart.Add_ValueChanged({UpdateRenamePreview})
$rNumPad.Add_ValueChanged({UpdateRenamePreview})
$rLst.Add_SelectedIndexChanged({UpdateRenamePreview})

function UpdateRenamePreview {
    $rLV.Items.Clear(); $num=[int]$rNumStart.Value
    foreach ($p in $rLst.Items) {
        $oName = Split-Path $p -Leaf
        $nName = Get-RenamedName -OldName $oName -NewName $rTxtNew.Text `
            -FindTxt $rTxtFind.Text -ReplaceTxt $rTxtRepl.Text `
            -Prefix $rTxtPre.Text -Suffix $rTxtSuf.Text `
            -UseNum $rChkNum.Checked -Num $num -Pad ([int]$rNumPad.Value)
        $item=New-Object System.Windows.Forms.ListViewItem($oName)
        $item.SubItems.Add($nName)|Out-Null
        if ($oName -ne $nName) { $item.ForeColor=$C_ACCENT }
        $rLV.Items.Add($item)|Out-Null
        if ($rChkNum.Checked){$num++}
    }
}

$rBtnRen.Add_Click({
    if ($rLst.Items.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Vyberte soubory.','',0,48)|Out-Null; return }
    if ([System.Windows.Forms.MessageBox]::Show("Prejmenovani $($rLst.Items.Count) souboru?","Potvrdit",4,32) -ne 6) { return }
    $rLog.Clear(); $ok=0; $fail=0; $num=[int]$rNumStart.Value
    foreach ($p in @($rLst.Items)) {
        $oName = Split-Path $p -Leaf
        $nName = Get-RenamedName -OldName $oName -NewName $rTxtNew.Text `
            -FindTxt $rTxtFind.Text -ReplaceTxt $rTxtRepl.Text `
            -Prefix $rTxtPre.Text -Suffix $rTxtSuf.Text `
            -UseNum $rChkNum.Checked -Num $num -Pad ([int]$rNumPad.Value)
        if ($rChkNum.Checked){$num++}
        if ($oName -eq $nName){continue}
        try {
            Rename-Item $p (Join-Path (Split-Path $p -Parent) $nName) -ErrorAction Stop
            AppendLog $rLog "OK: $oName -> $nName" $C_GREEN; $ok++
        } catch { AppendLog $rLog "CHYBA [$oName]: $_" $C_RED; $fail++ }
    }
    $sum="Hotovo: $ok prejmen."; if($fail -gt 0){$sum+=", $fail chyb"}
    AppendLog $rLog $sum $C_TEXT
    $rLst.Items.Clear(); UpdateRenamePreview
})

#endregion

#region LOGIKA: KOMPRESE

$script:kCustomOut = $null

$kBtnBrowse.Add_Click({
    $ofd=New-Object System.Windows.Forms.OpenFileDialog; $ofd.Multiselect=$true; $ofd.Filter='Obrazky|*.jpg;*.jpeg;*.png;*.webp;*.heic;*.heif|Vse|*.*'
    if ($ofd.ShowDialog($form) -eq 'OK') { Add-ToList $kLst $ofd.FileNames @('.jpg','.jpeg','.png','.webp','.heic','.heif') }
})
$kBtnClear.Add_Click({ $kLst.Items.Clear(); $kLog.Clear(); $kProg.Value=0 })
Register-DD $kLst @('.jpg','.jpeg','.png','.webp','.heic','.heif')
$kTrack.Add_ValueChanged({ $kLblQ.Text="$($kTrack.Value) %" })
$kChkRes.Add_CheckedChanged({ $kTxtW.Enabled=$kChkRes.Checked; $kTxtH.Enabled=$kChkRes.Checked })
$kBtnOut.Add_Click({
    $fbd=New-Object System.Windows.Forms.FolderBrowserDialog
    if ($fbd.ShowDialog($form) -eq 'OK') {
        $script:kCustomOut=$fbd.SelectedPath; $kTxtOut.Text=$fbd.SelectedPath; $kTxtOut.ForeColor=$C_TEXT; $kChkSfx.Checked=$false
    }
})

$kBtnComp.Add_Click({
    if ($kLst.Items.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Vyberte obrazky.','',0,48)|Out-Null; return }
    $q=$kTrack.Value; $mW=0; $mH=0
    if ($kChkRes.Checked){$mW=[int]$kTxtW.Text; $mH=[int]$kTxtH.Text}
    $files=@($kLst.Items)
    $kProg.Maximum=$files.Count; $kProg.Value=0; $kLog.Clear(); $kBtnComp.Enabled=$false
    $ok=0; $fail=0
    foreach ($src in $files) {
        $sName = Split-Path $src -Leaf
        $sExt  = [System.IO.Path]::GetExtension($src).ToLower()
        $base  = [System.IO.Path]::GetFileNameWithoutExtension($src)
        $dExt  = if($kRbJpg.Checked){'.jpg'}elseif($kRbPng.Checked){'.png'}elseif($kRbWebp.Checked){'.webp'}else{$sExt}
        if ($dExt -in @('.heic','.heif')) {
            AppendLog $kLog "Preskoceno [$sName] - do HEIC ukladat nelze, zvolte JPG / PNG / WebP" $C_ORANGE
            $kProg.Value++; continue
        }
        $oDir  = if($script:kCustomOut){$script:kCustomOut}else{Split-Path $src -Parent}
        $sfx   = if($kChkSfx.Checked -and -not $script:kCustomOut){'_k'}else{''}
        $dst   = Join-Path $oDir ($base+$sfx+$dExt)
        # PNG bez konverze na JPG nelze smysluplne komprimovat
        if ($sExt -eq '.png' -and $dExt -eq '.png' -and -not $kChkRes.Checked) {
            AppendLog $kLog "Preskoceno [$sName] - PNG je bezztrátový format, zmenšení není možné bez konverze na JPG" $C_ORANGE
            $kProg.Value++; continue
        }
        AppendLog $kLog "Komprimuje: $sName ..." $C_SUB
        try {
            $before = (Get-Item $src).Length
            CompressImage -SrcPath $src -DstPath $dst -Quality $q -MaxW $mW -MaxH $mH
            $after = (Get-Item $dst).Length
            $sameFmt = ($sExt -replace 'jpeg','jpg') -eq ($dExt -replace 'jpeg','jpg')
            if ($after -ge $before -and $sameFmt) {
                Copy-Item $src $dst -Force
                AppendLog $kLog "OK: $sName  (originál je uz optimalne zkomprimovany, zachovano)" $C_ORANGE
            } else {
                $pct = [int](100 - ($after / $before * 100))
                AppendLog $kLog "OK: $sName  ($([int]($before/1KB)) KB -> $([int]($after/1KB)) KB, -$pct%)" $C_GREEN
            }
            $ok++
        } catch { AppendLog $kLog "CHYBA [$sName]: $_" $C_RED; $fail++ }
        $kProg.Value++
    }
    $kBtnComp.Enabled=$true
    $sum="Hotovo: $ok uspesne"; if($fail -gt 0){$sum+=", $fail chyb"}
    AppendLog $kLog $sum $C_TEXT
    if ($ok -gt 0) {
        $dir=if($script:kCustomOut){$script:kCustomOut}else{Split-Path $files[0] -Parent}
        if ([System.Windows.Forms.MessageBox]::Show("$sum`nOtevrit slozku?","Hotovo",4,64) -eq 6) { Open-Folder $dir }
    }
})

#endregion

#region LOGIKA: PDF

$script:pMergeCustomOut   = $null
$script:pExtractCustomOut = $null

# --- Merge: tlacitka seznamu ---
$pMergeBtnBrowse.Add_Click({
    $ofd=New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Multiselect=$true; $ofd.Filter='PDF soubory|*.pdf|Vse|*.*'
    if ($ofd.ShowDialog($form) -eq 'OK') { Add-ToList $pMergeLst $ofd.FileNames @('.pdf') }
})
$pMergeBtnClear.Add_Click({ $pMergeLst.Items.Clear(); $pLog.Clear() })
Register-DD $pMergeLst @('.pdf')

$pMergeBtnUp.Add_Click({
    $i = $pMergeLst.SelectedIndex
    if ($i -le 0) { return }
    $item = $pMergeLst.Items[$i]
    $pMergeLst.Items.RemoveAt($i)
    $pMergeLst.Items.Insert($i-1, $item)
    $pMergeLst.SelectedIndex = $i-1
})
$pMergeBtnDown.Add_Click({
    $i = $pMergeLst.SelectedIndex
    if ($i -lt 0 -or $i -ge ($pMergeLst.Items.Count-1)) { return }
    $item = $pMergeLst.Items[$i]
    $pMergeLst.Items.RemoveAt($i)
    $pMergeLst.Items.Insert($i+1, $item)
    $pMergeLst.SelectedIndex = $i+1
})

$pMergeBtnOut.Add_Click({
    $fbd=New-Object System.Windows.Forms.FolderBrowserDialog
    if ($fbd.ShowDialog($form) -eq 'OK') {
        $script:pMergeCustomOut=$fbd.SelectedPath
        $pMergeTxtOut.Text=$fbd.SelectedPath; $pMergeTxtOut.ForeColor=$C_TEXT
    }
})

# --- Merge: akce ---
$pMergeBtnRun.Add_Click({
    if ($pMergeLst.Items.Count -lt 2) {
        [System.Windows.Forms.MessageBox]::Show('Pridejte alespon 2 PDF soubory.','',0,48)|Out-Null; return
    }
    $files = @($pMergeLst.Items)
    $oDir  = if ($script:pMergeCustomOut) { $script:pMergeCustomOut } else { Split-Path $files[0] -Parent }
    $dst   = Join-Path $oDir 'merged.pdf'
    $pLog.Clear()
    AppendLog $pLog "Slucuji $($files.Count) PDF souboru..." $C_SUB
    $pMergeBtnRun.Enabled=$false
    try {
        MergePdfs -Inputs $files -DstPath $dst
        AppendLog $pLog "OK: merged.pdf ulozeno do $oDir" $C_GREEN
        if ([System.Windows.Forms.MessageBox]::Show("Slouceno!`nOtevrit slozku?","Hotovo",4,64) -eq 6) { Open-Folder $oDir }
    } catch {
        AppendLog $pLog "CHYBA: $_" $C_RED
    } finally {
        $pMergeBtnRun.Enabled=$true
    }
})

# --- Extract: vyber souboru ---
$pExtractBtnSrc.Add_Click({
    $ofd=New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter='PDF soubory|*.pdf|Vse|*.*'
    if ($ofd.ShowDialog($form) -eq 'OK') {
        $pExtractTxtSrc.Text=$ofd.FileName; $pExtractTxtSrc.ForeColor=$C_TEXT
        $pExtractNumTo.Value=1; $pExtractNumFrom.Value=1
    }
})

$pExtractBtnOut.Add_Click({
    $fbd=New-Object System.Windows.Forms.FolderBrowserDialog
    if ($fbd.ShowDialog($form) -eq 'OK') {
        $script:pExtractCustomOut=$fbd.SelectedPath
        $pExtractTxtOut.Text=$fbd.SelectedPath; $pExtractTxtOut.ForeColor=$C_TEXT
    }
})

# --- Extract: akce ---
$pExtractBtnRun.Add_Click({
    $src = $pExtractTxtSrc.Text.Trim()
    if (-not $src -or -not (Test-Path $src)) {
        [System.Windows.Forms.MessageBox]::Show('Vyberte vstupni PDF soubor.','',0,48)|Out-Null; return
    }
    $fromP = [int]$pExtractNumFrom.Value
    $toP   = [int]$pExtractNumTo.Value
    if ($fromP -gt $toP) {
        [System.Windows.Forms.MessageBox]::Show("Stranka 'od' nesmi byt vetsi nez 'do'.","",0,48)|Out-Null; return
    }
    $oDir  = if ($script:pExtractCustomOut) { $script:pExtractCustomOut } else { Split-Path $src -Parent }
    $base  = [System.IO.Path]::GetFileNameWithoutExtension($src)
    $dst   = Join-Path $oDir "${base}_str${fromP}-${toP}.pdf"
    $pLog.Clear()
    AppendLog $pLog "Vyrizam stranky $fromP az $toP z: $(Split-Path $src -Leaf) ..." $C_SUB
    $pExtractBtnRun.Enabled=$false
    try {
        ExtractPdfPages -SrcPath $src -DstPath $dst -FromPage $fromP -ToPage $toP
        $outName = Split-Path $dst -Leaf
        AppendLog $pLog "OK: $outName ulozeno do $oDir" $C_GREEN
        if ([System.Windows.Forms.MessageBox]::Show("Stranky vyrizany!`nOtevrit slozku?","Hotovo",4,64) -eq 6) { Open-Folder $oDir }
    } catch {
        AppendLog $pLog "CHYBA: $_" $C_RED
    } finally {
        $pExtractBtnRun.Enabled=$true
    }
})

#endregion

[void]$form.ShowDialog()
