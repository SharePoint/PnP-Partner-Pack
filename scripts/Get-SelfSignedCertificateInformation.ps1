Param(
   [Parameter(Mandatory=$true)]
   [string]$CertificateFile
)

$certPath = Resolve-Path ".\$($CertificateFile).cer"
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$cert.Import($certPath)
$rawCert = $cert.GetRawCertData()
$base64Cert = [System.Convert]::ToBase64String($rawCert)
$rawCertHash = $cert.GetCertHash()
$base64CertHash = [System.Convert]::ToBase64String($rawCertHash)
$KeyId = [System.Guid]::NewGuid().ToString() #$AppClientId
 $result = @(@{
    
      customKeyIdentifier = "$base64CertHash"
      keyId = "$KeyId"
      type = 'AsymmetricX509Cert'
      usage = 'Verify'
      value =   "$base64Cert"
     }
 )
($result |ConvertTo-Json) | Out-File "./thumb.txt" -Append:$false -Force -Encoding string
("Certificate Thumbprint: $($cert.Thumbprint)")| Out-File "./thumb.txt" -Append:$true -Force -Encoding string
Write-Host "Certificate Thumbprint:" $cert.Thumbprint

@{
  KeyCredentials = $result
  CertificateThumbprint = $cert.Thumbprint
  RawCert=$rawCert
}