# License Validator for XrmToolBox

[![NuGet](https://img.shields.io/nuget/v/Consultics.XrmToolBox.LicenseValidator.svg)](https://www.nuget.org/packages/Consultics.XrmToolBox.LicenseValidator)
[![NuGet Downloads](https://img.shields.io/nuget/dt/Consultics.XrmToolBox.LicenseValidator.svg)](https://www.nuget.org/packages/Consultics.XrmToolBox.LicenseValidator)

An [XrmToolBox](https://www.xrmtoolbox.com/) plugin that validates Dynamics 365 / Dataverse user license assignments against actual security role rights, usage data and Microsoft Graph API information.

## Features

- **License Audit** – Compares assigned licenses with actual security role requirements per user
- **Usage Analysis** – Scans entity-level usage across Dataverse tables to identify inactive users
- **Graph API Integration** – Retrieves license and sign-in data from Microsoft Entra ID via Microsoft Graph
- **Optimization Recommendations** – Flags underlicensed, overlicensed and users requiring review
- **Excel Export** – Exports full audit results to Excel for further analysis and reporting
- **Configurable** – Adjustable thresholds for inactivity periods, usage scanning depth and more

## Screenshots

*(coming soon)*

## Installation

### Via XrmToolBox (recommended)
1. Open XrmToolBox → **Tool Library**
2. Search for **License Validator**
3. Click **Install**

### Via NuGet
```
nuget install Consultics.XrmToolBox.LicenseValidator
```

## Getting Started

1. Connect to your Dynamics 365 / Dataverse environment
2. Configure Graph API credentials (Tenant ID, Client ID, Client Secret)
3. Select an audit mode (Rights only, Usage only, or Rights & Usage)
4. Run the audit
5. Export results to Excel

## Requirements

- XrmToolBox (latest version recommended)
- Dynamics 365 / Dataverse environment
- Microsoft Entra ID App Registration with the following Graph API permissions:
  - `User.Read.All`
  - `AuditLog.Read.All`
  - `Organization.Read.All`

## Built With

- .NET Framework 4.8
- [XrmToolBox SDK](https://github.com/MscrmTools/XrmToolBox)
- [ClosedXML](https://github.com/ClosedXML/ClosedXML) – Excel export
- Microsoft Graph API

## Author

**Martin Jäger** – [Consultics AG](https://consultics.ch)

## License

This project is licensed under the MIT License – see the [LICENSE](LICENSE) file for details.

## Changelog

### 1.6.0
- Renamed plugin classes for consistency (`LicenseValidatorPlugin`, `LicenseValidatorControl`)
- Added NuGet package for XrmToolBox Plugin Store
- Added MIT license
- Initial public release
