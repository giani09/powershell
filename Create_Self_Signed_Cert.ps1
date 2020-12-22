# requires -Version 5

$Path = "$env:tmp\zertifikat.pfx"
$Password = Read-Host -Prompt 'Kennwort für Zertifikat' -AsSecureString

# Zertifikat anlegen
$cert = NewSelfSignedCertificate -KeyUsage DigitalSignature -KeySpec Signature -FriendlyName 'IT Abteilung' -Subject CN=TestAbteilung -KeyExportPolicy -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(5) -TextExtension @('2.5.29.37={text}1.3.6.15.5.7.3.3')
# Zertifikat in Datei exportieren:

$cert | Export-PfxCertificate -Password $Password -FilePath $Path

# Zertifikat aus Speicher löschen:
$cert | Remove-Item

#---- ----#

# Achtung: benötigt die Funktion New-SelfsignedCertificateEx
# Bezugsquelle: https://gallery.technet.microsoft.com/scriptcenter/
# Self-signed-certificate-5920a7c6
$Path = "$env:temp\zertifikat.pfx"

$Password = Read-Host -Prompt 'Kennwort für Zertifikat' -AsSecureString

New-SelfsignedCertificateEx -Exportable -Path $path -Password $password -Subject
'CN=Sicherheitsabteilung' -EKU '1.3.6.1.5.5.7.3.3' -KeySpec 'Signature' -KeyUsage
'DigitalSignature' -FriendlyName 'IT Sicherheit' -NotAfter (Get-Date).AddYears(5)

#---- ----#

# Zertifikat aus PFX-Datei laden
$path = "$env:temp\zertifikat.pfx"
$cert = Get-PfxCertificate -FilePath $Path
$cert

$cert | Select-Object -Property *

#---- -----#

# Zertifikat aus Zertifikatspeicher laden

Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert | Select-Object -Property Subject, Thumbprint

# Zertifikat in Skript einsetzen

$cert = Get-Item -Path Cert:\CurrentUser\My\$ThumbPrint