namespace OnlyExo365.Tests;

public sealed class PowerShellCryptographyCompatibilityTests
{
    [Fact]
    public async Task EmbeddedRunspace_CanSignVerifyAndEncryptCmsWithoutCertificateStore()
    {
        using var fixture = new LocalPowerShellFixture();
        var result = await fixture.Engine.ExecuteAsync("""
            $rsa = [System.Security.Cryptography.RSA]::Create(2048)
            $certificate = $null
            try {
                $request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
                    'CN=OnlyExo365 offline test', $rsa,
                    [System.Security.Cryptography.HashAlgorithmName]::SHA256,
                    [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
                $certificate = $request.CreateSelfSigned(
                    [DateTimeOffset]::UtcNow.AddMinutes(-1), [DateTimeOffset]::UtcNow.AddMinutes(5))
                $content = [System.Security.Cryptography.Pkcs.ContentInfo]::new(
                    [System.Text.Encoding]::UTF8.GetBytes('offline CMS payload'))
                $signed = [System.Security.Cryptography.Pkcs.SignedCms]::new($content, $false)
                $signed.ComputeSignature([System.Security.Cryptography.Pkcs.CmsSigner]::new($certificate))
                $decoded = [System.Security.Cryptography.Pkcs.SignedCms]::new()
                $decoded.Decode($signed.Encode())
                $decoded.CheckSignature($true)

                $enveloped = [System.Security.Cryptography.Pkcs.EnvelopedCms]::new($content)
                $enveloped.Encrypt([System.Security.Cryptography.Pkcs.CmsRecipient]::new($certificate))
                $decrypted = [System.Security.Cryptography.Pkcs.EnvelopedCms]::new()
                $decrypted.Decode($enveloped.Encode())
                $certificates = [System.Security.Cryptography.X509Certificates.X509Certificate2Collection]::new()
                [void]$certificates.Add($certificate)
                $decrypted.Decrypt($certificates)
                [pscustomobject]@{
                    SignedContent = [System.Text.Encoding]::UTF8.GetString($decoded.ContentInfo.Content)
                    DecryptedContent = [System.Text.Encoding]::UTF8.GetString($decrypted.ContentInfo.Content)
                }
            }
            finally {
                if ($null -ne $certificate) { $certificate.Dispose() }
                $rsa.Dispose()
            }
            """);

        Assert.True(result.Success, result.ErrorMessage);
        var output = Assert.Single(result.Output);
        Assert.Equal("offline CMS payload", output.Properties["SignedContent"]?.Value);
        Assert.Equal("offline CMS payload", output.Properties["DecryptedContent"]?.Value);
    }

    [Fact]
    public async Task EmbeddedRunspace_XmlSignatureAcceptsOriginalAndRejectsTamperedContent()
    {
        using var fixture = new LocalPowerShellFixture();
        var result = await fixture.Engine.ExecuteAsync("""
            $rsa = [System.Security.Cryptography.RSA]::Create(2048)
            try {
                $document = [System.Xml.XmlDocument]::new()
                $document.PreserveWhitespace = $true
                $document.LoadXml('<root><value>original</value></root>')
                $signed = [System.Security.Cryptography.Xml.SignedXml]::new($document)
                $signed.SigningKey = $rsa
                $reference = [System.Security.Cryptography.Xml.Reference]::new()
                $reference.Uri = ''
                $reference.AddTransform([System.Security.Cryptography.Xml.XmlDsigEnvelopedSignatureTransform]::new())
                $signed.AddReference($reference)
                $signed.ComputeSignature()
                [void]$document.DocumentElement.AppendChild($document.ImportNode($signed.GetXml(), $true))

                $verifier = [System.Security.Cryptography.Xml.SignedXml]::new($document)
                $verifier.LoadXml($document.GetElementsByTagName('Signature')[0])
                $originalValid = $verifier.CheckSignature($rsa)
                $document.SelectSingleNode('/root/value').InnerText = 'tampered'
                $verifier = [System.Security.Cryptography.Xml.SignedXml]::new($document)
                $verifier.LoadXml($document.GetElementsByTagName('Signature')[0])
                [pscustomobject]@{
                    OriginalValid = $originalValid
                    TamperedValid = $verifier.CheckSignature($rsa)
                }
            }
            finally { $rsa.Dispose() }
            """);

        Assert.True(result.Success, result.ErrorMessage);
        var output = Assert.Single(result.Output);
        Assert.Equal(true, output.Properties["OriginalValid"]?.Value);
        Assert.Equal(false, output.Properties["TamperedValid"]?.Value);
    }
}
