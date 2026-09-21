Import-Module Microsoft.Graph.Authentication -RequiredVersion 2.39.0 -ErrorAction Stop

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  TEST 2: CSP- & SOFTWARE-BESTAENDE DYNAMISCH ERMITTELN   " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# 1. Test: Subscriptions mit Filter
$testUris = @(
    "https://graph.microsoft.com/beta/directory/subscriptions?`$filter=serviceType eq 'Software'",
    "https://graph.microsoft.com/beta/directory/subscriptions?`$filter=purchaseChannel eq 'Reseller'",
    "https://graph.microsoft.com/v1.0/organization?`$select=id,displayName,assignedPlans",
    "https://graph.microsoft.com/beta/commerce/orders"
)

foreach ( $u in$testUris ) {
    Write-Host ""
    Write-Host "---> Abfrage: $u" -ForegroundColor Yellow
    try {
        $res = Invoke-MgGraphRequest -Method GET -Uri $u -ErrorAction Stop$cnt = 0
        if ($res.value) {
            $cnt =$res.value.Count
        } else {
            $cnt = 1
        }
        Write-Host "     HTTP 200 OK | Gefunden: $cnt Elemente" -ForegroundColor Green

        $items = @()
        if ($res.value) {
            $items = @($res.value)
        } else {
            $items = @($res)
        }

        foreach ( $it in $items ) {$name = ""
            if ($it.commercialRuntimeName) {
                $name =$it.commercialRuntimeName
            } elseif ($it.skuPartNumber) {
                $name =$it.skuPartNumber
            } elseif ($it.displayName) {
                $name =$it.displayName
            } else {
                $name =$it.id
            }
            Write-Host "     * $name (Typ: $($it.serviceType) | Kanal: $($it.purchaseChannel))" -ForegroundColor White
        }
    } catch {
        Write-Host "     Fehler: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# 2. Test: Alle 16 Roh-Objekte aus Endpunkt 1 inspizieren
Write-Host ""
Write-Host "---> Inspiziere alle 16 Objekte aus /beta/directory/subscriptions..." -ForegroundColor Cyan
try {
    $allSubs = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/directory/subscriptions" -ErrorAction Stop
    $subList = @()
    if ($allSubs.value) {
        $subList = @($allSubs.value)
    }

    Write-Host "Gesamtzahl Subscriptions: $($subList.Count)" -ForegroundColor Green
    $idx = 1
    foreach ( $s in $subList ) {$pName = ""
        if ($s.commercialRuntimeName) {
            $pName =$s.commercialRuntimeName
        } else {
            $pName =$s.skuPartNumber
        }
        Write-Host "[$idx] Name:$pName | SkuPart: $($s.skuPartNumber) | ServiceType: $($s.serviceType) | Channel: $($s.purchaseChannel) | Status: $($s.status)" -ForegroundColor White
        $idx++
    }
} catch {
    Write-Host "Fehler: $($_.Exception.Message)" -ForegroundColor Red
}