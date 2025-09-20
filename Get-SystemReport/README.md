# 🖥️ System Inventory Report Tool

This PowerShell script collects detailed system information across multiple servers and supports modular switches for targeted data collection. It also supports application comparison across servers in matrix format.

---

## 🚀 Usage

```powershell
.\Get-SystemReport.ps1 -ComputerName Server01,Server02 -Hotfix -OS -BIOS -Hardware -Apps -CompareApps -ExportCsv


🛠️ Supported Switches
Switch	Description
-Hotfix	Adds hotfix count to report
-HotfixDetail	Exports full hotfix list per machine
-OS	Collects OS version, build, install date, last boot time
-BIOS	Collects BIOS version, manufacturer, release date
-Hardware	Collects CPU, RAM, model, and system manufacturer
-Apps	Collects installed applications per machine and exports list
-CompareApps	Builds matrix comparing app versions across servers
-ExportCsv	Saves report and comparison to timestamped CSV files
📁 Output
Depending on the switches used, the script can generate:

✅ Console matrix view of system data

📄 CSV export of system report (SystemReport_YYYYMMDD_HHMMSS.csv)

📄 CSV export of app comparison matrix (AppComparison_YYYYMMDD_HHMMSS.csv)

📄 Per-server CSVs for hotfixes and installed applications (if -HotfixDetail or -Apps is used)

📦 Requirements
PowerShell 5.1 or later

Remote WMI access enabled on target machines

Administrator privileges (recommended for full data access)

Network connectivity to all target servers

🔐 Notes
App discovery uses registry-based queries to avoid the risks of Win32_Product.

All WMI queries are wrapped in try/catch blocks for fault tolerance.

The script is modular and can be extended with additional switches or export formats.

📬 Optional Enhancements
You can extend this tool with:

HTML email reporting

Parallel execution for faster scans

Filters for app mismatches or outdated hotfixes

Integration with config.json for centralized settings

🧠 Author Notes
This tool is designed for sysadmins, auditors, and IT teams who need fast, flexible visibility across environments. Contributions and improvements are welcome!
