$file = "C:\Users\o_antonov\NEW\RN\modReconcile.bas"
$log = "C:\Users\o_antonov\NEW\RN\verify_cp1251.log"
$cp1251 = [System.Text.Encoding]::GetEncoding(1251)
$out = [System.IO.File]::ReadAllBytes($file)
$patterns = @{
    "Slovar" = @(0xD1,0xEB,0xEE,0xE2,0xE0,0xF0,0xFC)
    "Nefteprodukt" = @(0xED,0xE5,0xF4,0xF2,0xE5,0xEF,0xF0,0xEE,0xE4,0xF3,0xEA,0xF2)
    "Zachisleno" = @(0xC7,0xE0,0xF7,0xE8,0xF1,0xEB,0xE5,0xED,0xEE)
    "RN_Slovar" = @(0x52,0x4E,0x5F,0x53,0x6C,0x6F,0x76,0x61,0x72)
}
function Test-Bytes($haystack, $needle) {
    for ($i = 0; $i -le $haystack.Length - $needle.Length; $i++) {
        $ok = $true
        for ($j = 0; $j -lt $needle.Length; $j++) {
            if ($haystack[$i + $j] -ne $needle[$j]) { $ok = $false; break }
        }
        if ($ok) { return $true }
    }
    return $false
}
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("File size: $($out.Length) bytes")
$lines.Add("First 3 bytes: $($out[0]),$($out[1]),$($out[2])")
$allOk = $true
foreach ($key in $patterns.Keys) {
    if (Test-Bytes $out $patterns[$key]) { $lines.Add("OK bytes: $key") } else { $lines.Add("FAIL bytes: $key"); $allOk = $false }
}
$text = $cp1251.GetString($out)
$slovar = $cp1251.GetString([byte[]]@(0xD1,0xEB,0xEE,0xE2,0xE0,0xF0,0xFC))
$nefte = $cp1251.GetString([byte[]]@(0xED,0xE5,0xF4,0xF2,0xE5,0xEF,0xF0,0xEE,0xE4,0xF3,0xEA,0xF2))
$zach = $cp1251.GetString([byte[]]@(0xC7,0xE0,0xF7,0xE8,0xF1,0xEB,0xE5,0xED,0xEE))
$lines.Add("Decode Slovar: $($text.Contains($slovar))")
$lines.Add("Decode Nefteprodukt: $($text.Contains($nefte))")
$lines.Add("Decode Zachisleno: $($text.Contains($zach))")
$lines.Add("Total lines: $(($text -split "`r?`n").Count)")
if ($allOk) { $lines.Add("RESULT: CP1251 VALID") } else { $lines.Add("RESULT: CP1251 INVALID") }
$lines -join "`n" | Set-Content -Path $log -Encoding UTF8
$lines
