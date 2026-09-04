#Requires -Version 5.1

<#
.SYNOPSIS
    General-purpose self-signed certificate generator for certificate-based
    auth against an Entra App Registration. Run as needed, once per distinct
    Purpose - each run produces its own independent keypair.

.DESCRIPTION
    Certificate-based Entra app auth only needs the PUBLIC key uploaded to
    the app registration; Entra just verifies signatures with it. The
    PRIVATE key stays wherever the authenticating application runs, because
    that is the side that signs the auth request.

    This script produces two files, named after -Purpose:
      - A .pfx (password-protected, contains the PRIVATE key). Keep this
        wherever the authenticating app/service runs, encrypted at rest.
        Never hand this to the app registration's tenant/owner - they only
        ever need the .cer.
      - A .cer (PUBLIC key only). This is the file uploaded to the Entra app
        registration's certificate credentials. Safe to hand out - it
        contains no secret material.

    NOTE: Exchange Online app-only auth specifically requires a CSP (legacy)
    key provider, not the CNG provider Windows uses by default for
    New-SelfSignedCertificate. This script always requests the CSP provider
    so certs it produces work with Connect-ExchangeOnline -CertificateThumbprint
    as well as standard Graph/MSAL certificate auth. Skipping -Provider is the
    most common reason EXO app-only setup silently fails at the very last step.

.PARAMETER Purpose
    What this certificate is for, e.g. "Client XYZ Graph App" or "3Fold IT
    Exchange Online App-Only". Used as the certificate's Subject (CN=Purpose)
    and, sanitized, as the .pfx/.cer filename - so each Purpose gets its own
    files instead of overwriting a previous run's. Prompted for if not
    supplied.

.PARAMETER OutputDirectory
    Where to write the .pfx and .cer files. Defaults to the current directory.

.PARAMETER YearsValid
    Certificate lifetime. Defaults to 5 years. Whoever owns this needs to
    regenerate (and re-upload the new .cer wherever it was registered) before
    it expires.

.EXAMPLE
    .\New-3FoldExoCertificate.ps1 -Purpose "Client XYZ Graph App" -OutputDirectory C:\Secure\Certs
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$Purpose,
    [string]$OutputDirectory = 'C:\MSP-M365-Utility\Certificates',
    [int]$YearsValid = 5
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

$safePurpose = ($Purpose -replace '[^\w-]', '')
if (-not $safePurpose) {
    throw "Purpose must contain at least one letter, digit, underscore, or hyphen."
}

$pfxPath = Join-Path $OutputDirectory "$safePurpose.pfx"
$cerPath = Join-Path $OutputDirectory "$safePurpose.cer"

if ((Test-Path $pfxPath) -or (Test-Path $cerPath)) {
    Write-Host "A certificate for '$Purpose' already exists at $OutputDirectory." -ForegroundColor Yellow
    Write-Host "Regenerating replaces the private key - wherever the old .cer was" -ForegroundColor Yellow
    Write-Host "uploaded (e.g. an Entra app registration) will stop trusting it until" -ForegroundColor Yellow
    Write-Host "the new .cer is re-uploaded there. Continuing overwrites the files here." -ForegroundColor Yellow
    $confirm = Read-Host "Type YES to overwrite and generate a new keypair"
    if ($confirm -ne 'YES') {
        Write-Host "Aborted. Nothing was changed." -ForegroundColor Red
        Read-Host "`nPress Enter to exit"
        exit
    }
}

Write-Host "Generating certificate for '$Purpose' (CSP provider, not CNG - required for EXO app-only auth)..." -ForegroundColor Cyan

$cert = New-SelfSignedCertificate `
    -Subject "CN=$Purpose" `
    -CertStoreLocation "cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec KeyExchange `
    -KeyLength 2048 `
    -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" `
    -NotAfter (Get-Date).AddYears($YearsValid)

Write-Host "Certificate created. Thumbprint: $($cert.Thumbprint)" -ForegroundColor Green

$pfxPassword = Read-Host "Set a password to protect the .pfx (private key) file" -AsSecureString
$cert | Export-PfxCertificate -FilePath $pfxPath -Password $pfxPassword | Out-Null
$cert | Export-Certificate -FilePath $cerPath | Out-Null

# Remove it from the local cert store now that it's exported - this script's
# job is to produce the files, not to leave a copy of the private key sitting
# in this machine's certificate store indefinitely.
Remove-Item -Path "cert:\CurrentUser\My\$($cert.Thumbprint)" -Force

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host "  Purpose: $Purpose"
Write-Host "  Private key (keep wherever the app authenticates from, encrypted, never share): $pfxPath"
Write-Host "  Public cert (upload to the Entra app registration's certificate credentials): $cerPath"
Write-Host "  Thumbprint: $($cert.Thumbprint)"
Write-Host ""
Write-Host "Record the thumbprint somewhere durable - whatever connects as this app later" -ForegroundColor Yellow
Write-Host "(Connect-MgGraph, Connect-ExchangeOnline, MSAL, etc.) needs it, or the .pfx" -ForegroundColor Yellow
Write-Host "itself, to authenticate." -ForegroundColor Yellow
Write-Host ""
Read-Host "Press Enter to exit"
