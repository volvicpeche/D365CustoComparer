
# CrmConnectionTool

CrmConnectionTool is a small console utility used to compare customizations between two Microsoft Dynamics 365 environments. It connects to two organizations using connection strings and outputs differences in forms or view definitions.

## Prerequisites

- **.NET Framework 4.8** – required to build and run the tool.
- **Visual Studio 2022** (or later) or the `msbuild` command line with .NET 4.8 installed.
- Dynamics 365 SDK assemblies are restored via NuGet (see `packages.config`).

## Building

1. Restore NuGet packages.
2. Open `CustoComparer/CrmConnectionTool.sln` in Visual Studio and build, or run:
   ```
   msbuild CustoComparer/CrmConnectionTool.sln /t:Restore,Build /p:Configuration=Release
   ```

The compiled executable will be in `CustoComparer/CrmConnectionTool/bin/Release`.

## Configuration

Connection strings and credentials are defined in `Program.cs`. Update the `conn1` and `conn2` strings with the URLs, usernames and domains for the source and target organizations. The password is read from `Settings.settings` (`CrmConnectionTool.Settings`). You can set the value once by editing the `Settings.settings` file or via the Visual Studio project settings.

Logs are written using log4net to the path configured in `log4net.config` (default is `C:\Logs\logs.csv`). Ensure this directory exists or adjust the path as needed.

## Usage

After building, run the executable from a command prompt. The tool will connect to both organizations and compare the specified forms in `Program.cs`.


