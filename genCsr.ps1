$SubjectDN = New-Object -ComObject X509Enrollment.CX500DistinguishedName
$SubjectDN.Encode("CN=www.contoso.com", 0x0)

$SAN = New-Object -ComObject X509Enrollment.CX509ExtensionAlternativeNames
$IANs = New-Object -ComObject X509Enrollment.CAlternativeNames

"www.contoso.com", "owa.contoso.com" | ForEach-Object {
    # instantiate a IAlternativeName object
    $IAN = New-Object -ComObject X509Enrollment.CAlternativeName
    # initialize the object by using current element in the pipeline
    $IAN.InitializeFromString(0x3,$_)
    # add created object to an object collection of IAlternativeNames
    $IANs.Add($IAN)
}
# finally, initialize SAN extension from a collection of alternative names:
$SAN.InitializeEncode($IANs)

$PrivateKey = New-Object -ComObject X509Enrollment.CX509PrivateKey -Property @{
    ProviderName = "Microsoft RSA SChannel Cryptographic Provider"
    MachineContext = $true
    Length = 2048
    KeySpec = 1
    KeyUsage = [int][Security.Cryptography.X509Certificates.X509KeyUsageFlags]::KeyEncipherment
}

$PrivateKey.Create()

$KeyUsage = New-Object -ComObject X509Enrollment.CX509ExtensionKeyUsage
$KeyUsage.InitializeEncode([int][Security.Cryptography.X509Certificates.X509KeyUsageFlags]"DigitalSignature,KeyEncipherment")
$KeyUsage.Critical = $true

# create appropriate interface objects
$EKU = New-Object -ComObject X509Enrollment.CX509ExtensionEnhancedKeyUsage
$OIDs = New-Object -ComObject X509Enrollment.CObjectIDs
"Server Authentication", "Client Authentication" | ForEach-Object {
    # transform current element to an Oid object. This is necessary to retrieve OID value.
    # this step is not required when you pass OID values directly.
    $netOid = New-Object Security.Cryptography.Oid $_
    # instantiate a IObjectID object for current element.
    $OID = New-Object -ComObject X509Enrollment.CObjectID
    # initialize the object with current enhanced key usage
    $OID.InitializeFromValue($netOid.Value)
    # add the object to an object collection
    $OIDs.Add($OID)
}
# when all EKUs are processed, initialized the IX509ExtensionEnhancedKeyUsage with the IObjectIDs collection
$EKU.InitializeEncode($OIDs)

$PKCS10 = New-Object -ComObject X509Enrollment.CX509CertificateRequestPkcs10

# 0x2 argument for Context parameter indicates that the request is intended for computer (or machine context).
# strTemplateName parameter is optional and we pass just empty string.
$PKCS10.InitializeFromPrivateKey(0x2,$PrivateKey,"")

$PKCS10.Subject = $SubjectDN
$PKCS10.X509Extensions.Add($SAN)
$PKCS10.X509Extensions.Add($EKU)
$PKCS10.X509Extensions.Add($KeyUsage)

# instantiate IX509Enrollment object
$Request = New-Object -ComObject X509Enrollment.CX509Enrollment
# provide certificate friendly name:
$Request.CertificateFriendlyName = "My cool SSL cert"
# initialize the object from PKCS#10 object:
$Request.InitializeFromRequest($PKCS10)

$Base64 = $Request.CreateRequest(0x3)
#Set-Content $path -Value $Base64 -Encoding Ascii

