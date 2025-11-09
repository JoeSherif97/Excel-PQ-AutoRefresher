# ============================================
# Excel Auto Refresh Script (Recursive + Safe)
# ============================================

# --- CONFIG ---
$ExcelPath = "C:\ExcelTest"   # Main folder containing Excel files (and subfolders)
$LogFile   = "C:\ScriptPs\Work\ExcelRefreshLog.txt"

# --- FUNCTIONS ---
function Write-Log($Line) {
    $utf8 = New-Object System.Text.UTF8Encoding $false
    $folder = Split-Path $LogFile
    if (!(Test-Path $folder)) { New-Item -ItemType Directory -Path $folder | Out-Null }

    $sw = [System.IO.StreamWriter]::new($LogFile, $true, $utf8)
    $sw.WriteLine("[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Line")
    $sw.Close()
}

# --- MAIN ---
try {
    Write-Log "=== Starting Excel refresh process (recursive) ==="
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # ✅ Get all .xlsx files (including subfolders)
    $files = Get-ChildItem -Path $ExcelPath -Filter "*.xlsx" -File -Recurse
    $total = $files.Count
    if ($total -eq 0) {
        Write-Host "No Excel files found under $ExcelPath" -ForegroundColor Red
        Write-Log "No Excel files found — nothing to refresh."
        exit
    }

    $count = 0
    $refreshed = 0
    $skipped = 0

    foreach ($file in $files) {
        $count++
        Write-Host "`n[$count/$total] Checking: $($file.FullName)" -ForegroundColor Yellow
        Write-Log "Opening workbook: $($file.FullName)"

        try {
            $workbook = $excel.Workbooks.Open($file.FullName, $false, $false)

            # ✅ Skip if no Power Query connections
            if ($workbook.Connections.Count -eq 0) {
                Write-Host "  → No Power Query connections, skipped." -ForegroundColor DarkGray
                Write-Log "Skipped (no Power Query connections): $($file.FullName)"
                $workbook.Close($false)
                $skipped++
            }
            else {
                Write-Host "  → Refreshing..." -ForegroundColor Cyan
                Write-Log "Refreshing workbook: $($file.FullName)"

                # Disable background refresh for all connections
                foreach ($conn in $workbook.Connections) {
                    try {
                        if ($conn.OLEDBConnection) {
                            $conn.OLEDBConnection.BackgroundQuery = $false
                        }
                    } catch {}
                }

                # Refresh all connections synchronously
                $workbook.RefreshAll()
                
                # Wait until all connections finish
                $allDone = $false
                while (-not $allDone) {
                    $allDone = $true
                    foreach ($conn in $workbook.Connections) {
                        try {
                            if ($conn.OLEDBConnection.Refreshing) {
                                $allDone = $false
                                break
                            }
                        } catch {}
                    }
                    Start-Sleep -Seconds 2
                }

                $workbook.Save()
                $workbook.Close($false)
                Write-Log "Refreshed successfully: $($file.FullName)"
                $refreshed++
            }

            # ✅ Progress bar (properly formatted, no syntax errors)
            $percent = [math]::Round(($count / $total) * 100)
            $statusText = "Refreshed: $refreshed | Skipped: $skipped | Total: $total"
            Write-Progress -Activity "Refreshing Excel Workbooks (Recursive)" `
                           -Status $statusText `
                           -PercentComplete $percent
        }
        catch {
            Write-Log "Error processing $($file.FullName): $_"
        }
    }

    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Log "=== Excel refresh process finished (recursive) ==="
    Write-Host "`nAll Excel files processed sequentially!" -ForegroundColor Green
    Write-Host "Refreshed: $refreshed | Skipped: $skipped | Total: $total" -ForegroundColor Cyan
}
catch {
    Write-Log "Fatal error: $_"
}
