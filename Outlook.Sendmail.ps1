
#
# Copyright (c) 2025 Huang Qinjin (huangqinjin@gmail.com)
#
# Distributed under the Boost Software License, Version 1.0.
#    (See accompanying file LICENSE_1_0.txt or copy at
#          https://www.boost.org/LICENSE_1_0.txt)
#

$lines = @()
while ($true) {
    $line = [Console]::ReadLine()
    if ($null -eq $line) {
        break
    }
    $lines += $line
}

function Get-HeadersContent($lines) {
    $headers = @{}
    $content = $null
    foreach ($line in $lines) {
        if ($null -ne $content) {
            # Body
            $content += $line
        }
        elseif ($line -eq '') {
            # End of Headers
            $content = @()
        }
        elseif ($line -match '^([^\s:]+):\s*(.*)$') {
            # Header Field Format
            # https://datatracker.ietf.org/doc/html/rfc5322#section-2.2
            $field = $matches[1]
            $headers[$field] = $matches[2]
        } 
        elseif ($line -match '^\s+') {
            # Header Field Unfolding
            # https://datatracker.ietf.org/doc/html/rfc5322#section-2.2.3
            $headers[$field] += $line
        }
    }
    return $headers, $content
}
function Get-AddressList($a) {
    # https://datatracker.ietf.org/doc/html/rfc5322#section-3.4
    return ($a -split ',' | ForEach-Object {
        $_ -replace '^.*<(.*)>.*$', '$1'
    }) -join ';'
}

function ConvertFrom-QuotedPrintable($content) {
    $sb = [System.Text.StringBuilder]::new()
    $stream = [System.IO.MemoryStream]::new()
    foreach ($line in $content) {
        $stream.SetLength(0)        
        $bytes = [System.Text.Encoding]::ASCII.GetBytes($line)
        for ($i = 0; $i -lt $bytes.Count; ++$i) {
            if ($bytes[$i] -eq [char]'=' -and $i -lt $bytes.Count - 1) {
                $stream.WriteByte([Convert]::ToByte([System.Text.Encoding]::ASCII.GetString($bytes, $i + 1, 2), 16))
                $i += 2
            } else {
                $stream.WriteByte($bytes[$i])
            }
        }
        $chars = [System.Text.Encoding]::UTF8.GetChars($stream.ToArray())
        if ($chars[$chars.Count - 1] -eq '=') {
            [void]$sb.Append($chars, 0, $chars.Count - 1)
        } else {
            [void]$sb.Append($chars).Append("`r`n")
        }
    }
    return $sb.ToString()
}

$headers, $content = Get-HeadersContent($lines)
$parts = @()
if ($headers['Content-Type'] -match 'multipart/alternative;\s*boundary=(.+)$') {
    $boundary = $matches[1]
    $part = $null
    foreach ($line in $content) {
        if ($line -eq "--$boundary" -or $line -eq "--$boundary--") {
            if ($part.Count -gt 0) {
                $r = @{}
                $r.headers, $r.content = Get-HeadersContent($part)
                $parts += $r
            }
            if ($line -eq "--$boundary--") {
                break
            }
            $part = @()
        }
        else {
            $part += $line
        }
    }
}
else {
    $parts += @{
        headers = $headers
        content = $content
    }
}

$subject = $headers['Subject']
# https://datatracker.ietf.org/doc/html/rfc5322#section-3.6.3
$to = Get-AddressList($headers['To'])
$cc = Get-AddressList($headers['Cc'])
$bcc = Get-AddressList($headers['Bcc'])

$Outlook = New-Object -ComObject Outlook.Application
$Mail = $Outlook.CreateItem(0)
$Mail.Subject = $subject
$Mail.To = $to
$Mail.CC = $cc
$Mail.BCC = $bcc

foreach ($part in $parts) {
    $headers = $part.headers
    $content = $part.content

    if ($headers['Content-Transfer-Encoding'] -eq 'quoted-printable') {
        $body = ConvertFrom-QuotedPrintable($content)
    } else {
        $body = $content -join "`r`n"
    }

    if ($headers['Content-Type'] -match 'text/html') {
        $Mail.HTMLBody = $body
    }
    else  {
        # olFormatPlain (Value: 1) for Plain Text
        # olFormatHTML (Value: 2) for HTML
        # olFormatRichText (Value: 3) for Rich Text
        $Mail.BodyFormat = 1
        $Mail.Body = $body
    }
}

$Mail.Send()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Mail) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
