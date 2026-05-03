param(
    [Parameter(Mandatory = $true)]
    [string]$PayloadPath
)

$ErrorActionPreference = "Stop"

function Write-Result {
    param([hashtable]$Result)
    $Result | ConvertTo-Json -Compress -Depth 5
}

try {
    $payload = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $attachmentPath = $payload.attachment_path

    if ($attachmentPath -and -not (Test-Path -LiteralPath $attachmentPath)) {
        Write-Result @{
            success = $false
            created = 0
            errors = @()
            error = "Attachment file was not found on the local machine."
        }
        exit 1
    }

    try {
        $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    }
    catch {
        $outlook = New-Object -ComObject Outlook.Application
    }

    $created = 0
    $errors = @()

    foreach ($recipient in $payload.recipients) {
        try {
            $mail = $outlook.CreateItem(0)
            $mail.To = [string]$recipient
            $mail.Subject = [string]$payload.subject
            $mail.Body = [string]$payload.body

            if ($attachmentPath) {
                [void]$mail.Attachments.Add((Resolve-Path -LiteralPath $attachmentPath).Path)
            }

            $mail.Display($false)
            $created += 1
        }
        catch {
            $errors += @{
                recipient = [string]$recipient
                error = $_.Exception.Message
            }
        }
    }

    Write-Result @{
        success = ($created -gt 0)
        created = $created
        errors = $errors
        error = $(if ($created -gt 0) { "" } else { "No Outlook draft was created." })
    }

    if ($created -gt 0) {
        exit 0
    }

    exit 1
}
catch {
    Write-Result @{
        success = $false
        created = 0
        errors = @()
        error = $_.Exception.Message
    }
    exit 1
}
