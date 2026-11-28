function Resolve-ReleaseTag {
    [CmdletBinding()]
    param([string]$EventName, [string]$PublishedTag, [string]$RequestedTag)

    $tag = if ($EventName -eq 'release') { $PublishedTag } else { $RequestedTag }
    if ([string]::IsNullOrWhiteSpace($tag) -or $tag -match '[\r\n]' -or $tag.StartsWith('-')) {
        throw 'A valid, single-line release tag is required.'
    }

    # Values remain native arguments, even when a legal Git tag contains shell syntax.
    & git check-ref-format "refs/tags/$tag" 2>$null
    if ($LASTEXITCODE -ne 0) { throw 'Release tag is not a valid Git reference.' }
    return $tag
}
