# Certificate Information Retrieval Script

This PowerShell script retrieves certificate information from a Certificate Authority (CA) server and provides filtering and exporting options. It connects to the CA server using the `CertificateAuthority.View.1` COM object and retrieves various certificate attributes such as CommonName, Email, NotAfter, Country, Organization, OrgUnit, DistinguishedName, and Disposition.

## Prerequisites

- Windows operating system with PowerShell installed.
- Access to a Certificate Authority (CA) server with the required COM object (`CertificateAuthority.View.1`).

## Usage

1. Clone or download the script from this GitHub repository.

2. Open a PowerShell environment (e.g., PowerShell Console or PowerShell ISE).

3. Navigate to the directory where the script is saved.

4. Run the script using the following command:

    ```powershell
    .\Get-CertInfo.ps1 -Server "CA-SERVER-NAME" [-FilterByOrganization "ORGANIZATION-NAME"] [-FilterByExpiration "EXPIRATION-DATE"] [-ExportPath "OUTPUT-FILE-PATH"]
    ```

   Replace the placeholders with the appropriate values:

   - `CA-SERVER-NAME`: The name or IP address of the Certificate Authority server.
   - `ORGANIZATION-NAME` (optional): Filter certificates by organization name. Specify this parameter if you want to filter by organization.
   - `EXPIRATION-DATE` (optional): Filter certificates by expiration date. Specify this parameter if you want to filter by expiration date.
   - `OUTPUT-FILE-PATH` (optional): Export the certificate information to a CSV file at the specified location.

   Examples:

   - Retrieve all certificates from the CA server:
     ```powershell
     .\Get-CertInfo.ps1 -Server "CA-SERVER-NAME"
     ```

   - Retrieve certificates for a specific organization:
     ```powershell
     .\Get-CertInfo.ps1 -Server "CA-SERVER-NAME" -FilterByOrganization "EXAMPLE-ORG"
     ```

   - Retrieve certificates expiring before a certain date and export the result to a CSV file:
     ```powershell
     .\Get-CertInfo.ps1 -Server "CA-SERVER-NAME" -FilterByExpiration "2023-12-31" -ExportPath "C:\Output\Certificates.csv"
     ```

5. The script will establish a connection to the specified CA server, retrieve the certificate information based on the provided filters, and display the result in the PowerShell console. If you specified an export path using the `-ExportPath` parameter, it will also export the result to a CSV file at the specified location.

## Notes

- Make sure you have the necessary permissions and dependencies in place for the script to execute successfully.
- The script assumes the availability of the `CertificateAuthority.View.1` COM object. If you encounter any issues related to this COM object, ensure it is properly installed and registered on your system.
- For any errors or issues, please raise an issue on this GitHub repository.

