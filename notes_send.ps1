# Send one mail through the local Notes client (Lotus.NotesSession COM).
# Called from notify.py. Notes is often 32-bit, so notify.py starts the PowerShell
# whose bitness matches Notes. Input is JSON on stdin (UTF-8):
#   {"mode": "send" | "check", "password": "", "to": [...], "subject": "", "body": ""}
# Keep this file ASCII only (Windows PowerShell 5.1 reads BOM-less scripts as ANSI).
$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [Text.Encoding]::UTF8
    $stdin = [Console]::OpenStandardInput()
    $buf = New-Object IO.MemoryStream
    $stdin.CopyTo($buf)
    $m = [Text.Encoding]::UTF8.GetString($buf.ToArray()) | ConvertFrom-Json

    $session = New-Object -ComObject Lotus.NotesSession
    if ($m.password) { $session.Initialize([string]$m.password) } else { $session.Initialize() }
    $db = $session.GetDbDirectory('').OpenMailDatabase()
    if (-not $db -or -not $db.IsOpen) { throw 'Could not open the mail database of the current Notes user.' }

    if ($m.mode -eq 'check') {
        Write-Output ([string]$session.CommonUserName)
        exit 0
    }

    $doc = $db.CreateDocument()
    [void]$doc.ReplaceItemValue('Form', 'Memo')
    [void]$doc.ReplaceItemValue('SendTo', [object[]]@($m.to))
    [void]$doc.ReplaceItemValue('Subject', [string]$m.subject)
    $rt = $doc.CreateRichTextItem('Body')
    $lines = ([string]$m.body) -split "`r?`n"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $rt.AppendText($lines[$i])
        if ($i -lt $lines.Count - 1) { $rt.AddNewLine(1) }
    }
    $doc.SaveMessageOnSend = $true
    $doc.Send($false)
    Write-Output 'OK'
    exit 0
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
