# License Validator for XrmToolBox

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

## Getting Started

1. Install [XrmToolBox](https://www.xrmtoolbox.com/)
2. Open the **Tool Library** inside XrmToolBox
3. Search for **License Validator** and install it
4. Connect to your Dynamics 365 / Dataverse environment
5. Configure Graph API credentials (Tenant ID, Client ID, Client Secret)
6. Run the audit

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
