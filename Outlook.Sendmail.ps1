
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
    if ($headers['Content-Type'] -match 'text/html') {
        $Mail.HTMLBody = $content -join "`r`n"
    }
    else  {
        # olFormatPlain (Value: 1) for Plain Text
        # olFormatHTML (Value: 2) for HTML
        # olFormatRichText (Value: 3) for Rich Text
        $Mail.BodyFormat = 1
        $Mail.Body = $content -join "`r`n"
    }
}

$Mail.Send()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Mail) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
