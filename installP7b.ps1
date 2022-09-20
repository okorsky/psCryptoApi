$enroll = New-Object -ComObject X509Enrollment.CX509Enrollment
$enroll.Initialize(0x1)
$raw = Get-Content -Encoding Byte "D:\dev\cont3.p7b"
$base64 = [Convert]::ToBase64String($raw)
$enroll.InstallResponse(0,$Base64,0x1,"")