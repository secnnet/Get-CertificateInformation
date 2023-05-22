function Get-CertInfo {
    param (
        [string]$Server,                      # The name or IP address of the Certificate Authority server
        [string]$FilterByOrganization,        # Optional: Filter certificates by organization name
        [string]$FilterByExpiration           # Optional: Filter certificates by expiration date
    )

    try {
        # Establish connection to Certificate server
        $CaView = New-Object -Com CertificateAuthority.View.1
        $CaView.OpenConnection($Server)

        # Define the numbers of columns
        $NumberOfColumns = 8
        $CaView.SetResultColumnCount($NumberOfColumns)
        $Index0 = $CAView.GetColumnIndex($False, "CommonName")
        $Index1 = $CAView.GetColumnIndex($False, "Email")
        $Index2 = $CAView.GetColumnIndex($False, "NotAfter")
        $Index3 = $CAView.GetColumnIndex($False, "Country")
        $Index4 = $CAView.GetColumnIndex($False, "Organization")
        $Index5 = $CAView.GetColumnIndex($False, "OrgUnit")
        $Index6 = $CAView.GetColumnIndex($False, "DistinguishedName")
        $Index7 = $CAView.GetColumnIndex($False, "Disposition")

        $CAView.SetResultColumn($Index0)
        $CAView.SetResultColumn($Index1)
        $CAView.SetResultColumn($Index2)
        $CAView.SetResultColumn($Index3)
        $CAView.SetResultColumn($Index4)
        $CAView.SetResultColumn($Index5)
        $CAView.SetResultColumn($Index6)
        $CAView.SetResultColumn($Index7)

        $RowObj = $CAView.OpenView()
        [void]$RowObj.Next()

        $CertList = @()

        while ($RowObj.Next() -eq 0) {
            $ColObj = $RowObj.EnumCertViewColumn()
            [void]$ColObj.Next()

            # Construct certificate object with attributes
            $CertInfo = [PSCustomObject]@{
                IssuingCA          = $Server
                CommonName         = $ColObj.GetValue(1)
                Email              = $ColObj.GetValue(2)
                NotAfter           = $ColObj.GetValue(3)
                Country            = $ColObj.GetValue(4)
                Organization       = $ColObj.GetValue(5)
                OrgUnit            = $ColObj.GetValue(6)
                DistinguishedName  = $ColObj.GetValue(7)
                Disposition        = $ColObj.GetValue(8)
            }

            # Apply filters if provided
            if ($FilterByOrganization -and $CertInfo.Organization -ne $FilterByOrganization) {
                continue    # Skip certificate if it doesn't match the organization filter
            }
            if ($FilterByExpiration -and $CertInfo.NotAfter -lt (Get-Date)) {
                continue    # Skip certificate if it has expired
            }

            $CertList += $CertInfo

            Clear-Variable ColObj
        }

        # Sort the certificates by NotAfter date
        $CertList = $CertList | Sort-Object -Property NotAfter

        # Output the result
        $CertList

        # Export to CSV if requested
        if ($ExportPath) {
            $CertList | Export-Csv -Path $ExportPath -NoTypeInformation
        }
    }
    catch {
        Write-Error "An error occurred: $_"
    }
    finally {
        if ($CaView) {
            $CaView.CloseConnection()
        }
