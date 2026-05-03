param(
    [Parameter(Mandatory = $true)]
    [string]$OutlookPath,

    [Parameter(Mandatory = $true)]
    [string]$Recipient,

    [Parameter(Mandatory = $true)]
    [string]$Subject,

    [Parameter(Mandatory = $true)]
    [string]$Body,

    [string]$AttachmentPath = ""
)

$ErrorActionPreference = "Stop"

$subjectEncoded = [System.Uri]::EscapeDataString($Subject)
$bodyEncoded = [System.Uri]::EscapeDataString($Body)
$messageArg = "${Recipient}?subject=${subjectEncoded}&body=${bodyEncoded}"

if ($AttachmentPath -and (Test-Path -LiteralPath $AttachmentPath)) {
    & $OutlookPath /c ipm.note /m $messageArg /a $AttachmentPath
}
else {
    & $OutlookPath /c ipm.note /m $messageArg
}

exit $LASTEXITCODE
