[Console]::In.ReadToEnd() | Out-Null
Start-Sleep -Seconds 5
[ordered]@{ status = 'Passed'; writeBackAllowed = $false; candidates = @(); diagnostics = [ordered]@{ reason = 'Should not be returned.' } } | ConvertTo-Json -Compress
