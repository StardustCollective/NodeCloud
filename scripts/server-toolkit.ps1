    if ([string]::IsNullOrWhiteSpace($newUser)) {
        $AppendLog.Invoke("[ERROR] New Username is required.`r`n")
        return
    }

    # Clean OpenSSH known_hosts for this host/port so ssh.exe does not choke on old keys
    try {
        $knownHostsPath = Join-Path (Join-Path $env:USERPROFILE ".ssh") "known_hosts"
        if (Test-Path $knownHostsPath) {
            $allLines    = Get-Content -Path $knownHostsPath
            $escapedHost = [regex]::Escape($serverHost)
            $escapedPort = [regex]::Escape($port)

            # Match lines like:
            #   65.108.87.84 ssh-ed25519 ...
            #   [65.108.87.84]:22 ssh-ed25519 ...
            $pattern1 = "^$escapedHost\s"
            $pattern2 = "^\[$escapedHost\]:$escapedPort\s"

            $filtered = $allLines | Where-Object {
                $_ -notmatch $pattern1 -and $_ -notmatch $pattern2
            }

            if ($filtered.Count -ne $allLines.Count) {
                $filtered | Set-Content -Path $knownHostsPath
                $AppendLog.Invoke("Cleaned OpenSSH known_hosts entries for $serverHost (port $port).`r`n")
            }
        }
    } catch {
        $AppendLog.Invoke("[WARN] Failed to clean OpenSSH known_hosts: $($_.Exception.Message)`r`n")
    }

    # SSH key handling:
    # - If user selects a .ppk, convert it to an OpenSSH private key
