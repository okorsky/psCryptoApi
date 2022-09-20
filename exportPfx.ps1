function Export-Certificate($certificateToExport)

{
    $certCollection = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2Collection
    $certChain = New-Object System.Security.Cryptography.X509Certificates.X509Chain
    $certChain.ChainPolicy.RevocationMode = "NoCheck"
    [void]$certChain.Build($certificateToExport)
    $certChain.ChainElements | ForEach-Object {[void]$certCollection.Add($_.Certificate)}
    $certChain.Reset()
    return $certCollection
}

$chain = Export-Certificate $cert
[array]::Reverse($chain)

$cert = Get-ChildItem -Path "Cert:\CurrentUser\My\914B66AB8CB7105EA8C54EBC3B028D6FAAA78764"
$certColl = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2Collection
$certColl.AddRange($chain)

$export = $certColl.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pkcs12, "test")

Set-Content -Encoding Byte -Value $export -Path "testchain.pfx"