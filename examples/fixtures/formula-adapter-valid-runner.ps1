[Console]::In.ReadToEnd() | Out-Null
[ordered]@{
    status = 'Passed'
    writeBackAllowed = $false
    candidates = @(
        [ordered]@{
            raw = 'Q_{\\text{放}}=qm'
            score = 0.99
            status = 'Candidate'
            warnings = @('fixture-only')
        }
    )
    diagnostics = [ordered]@{ reason = 'Fixture runner completed.' }
} | ConvertTo-Json -Depth 8 -Compress
