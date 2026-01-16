# Project Context & Instructions for Claude

## 1. Project Architecture
* **Type:** Visual Studio Add-in (.NET Framework).
* **Project System:** "Old-style" `.csproj` structure (Non-SDK style).
    * **Crucial:** Unlike .NET Core, this project does *not* automatically include new files.
    * **Rule:** When creating a new `.cs` class, you **MUST** explicitly add it to the `.csproj` file using an `<Compile Include="..." />` entry, or the build will fail.

## 2. Terminal & Shell Discipline
* **Primary Shell:** Always prefer **PowerShell** over Bash.
* **Bash Restrictions:**
    * If you must use the `Bash` tool, **DO NOT** use Windows-style flags.
    * **BAD:** `dir /b` (Causes "Exit code 2" in Bash).
    * **GOOD:** `ls -1` or `find . -maxdepth 1`.
* **Pathing:** Use quotes around paths to handle spaces in `C:\Users\wally\...`.

## 3. Build Instructions
* **Forbidden Command:** Do **NOT** use `dotnet build`.
    * *Reason:* It fails with error `MSB4019` because it cannot find `Microsoft.VisualStudio.Tools.Office.targets`.
* **Correct Command:** You must use the full path to `MSBuild.exe` in PowerShell.

### Standard Build Command
Run this in PowerShell to rebuild:

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "FastPMHelperAddin.sln" /t:Rebuild /p:Configuration=Debug

## 5. Project Map (Key Locations)
* **Entry Point (VSTO):** `Outlook\ThisAddIn.cs`
    * *Start here for Add-in lifecycle events (Startup/Shutdown).*
* **Main UI Panel:** `UI\ProjectActionPane.xaml`
    * *This is the primary Task Pane visible to the user.*
* **Business Logic:** `Services\`
    * *Contains all functional logic (Graph API, Google Sheets, AI, etc.).*
* **Data Models:** `Models\`
    * *POCOs for ActionItems, Rules, and data transfer.*
* **Configuration:** `Configuration\` & `Properties\`
    * *User settings and environment config.*