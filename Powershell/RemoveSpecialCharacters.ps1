function Remove-Chars {
param ([String]$src = [String]::Empty)
#replace diacritics
$normalized = $src.Normalize( [Text.NormalizationForm]::FormD )
$sb = new-object Text.StringBuilder
$normalized.ToCharArray() | % {
if( [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
[void]$sb.Append($_)
}
}
$sb=$sb.ToString()
#replace via code page conversion
$NonUnicodeEncoding = [System.Text.Encoding]::GetEncoding(850)
$UnicodeEncoding = [System.Text.Encoding]::Unicode
[Byte[]] $UnicodeBytes = $UnicodeEncoding.GetBytes($sb);
[Byte[]] $NonUnicodeBytes = [System.Text.Encoding]::Convert($UnicodeEncoding, $NonUnicodeEncoding , $UnicodeBytes);
[Char[]] $NonUnicodeChars = New-Object -TypeName “Char[]” -ArgumentList $($NonUnicodeEncoding.GetCharCount($NonUnicodeBytes, 0, $NonUnicodeBytes.Length))
[void] $NonUnicodeEncoding.GetChars($NonUnicodeBytes, 0, $NonUnicodeBytes.Length, $NonUnicodeChars, 0);
[String] $NonUnicodeString = New-Object String(,$NonUnicodeChars)


Return $NonUnicodeString


}
