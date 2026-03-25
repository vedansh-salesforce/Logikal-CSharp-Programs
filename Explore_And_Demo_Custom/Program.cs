// ============================================================
// Logikal Headless ERP Export
// Reads configuration from appsettings.json instead of prompting the user.
//
// TWO MODES:
//   "Discover" - lists all recent projects with their GUIDs so you can
//                find what to put in appsettings.json. Use this first.
//   "Export"   - exports the ERP SQLite for the configured project GUID.
//
// HOW TO USE:
//   1. Edit appsettings.json and set Mode = "Discover"
//   2. Run the program - it will print all projects and their GUIDs
//   3. Copy the GUID of the project you want to export
//   4. Edit appsettings.json: paste the GUID and set Mode = "Export"
//   5. Run again - it will save erp_export.sqlite3 to OutputDirectory
// ============================================================

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;

// Core (non-UI) API namespace - everything we need for headless operation
using Ofcas.Lk.Api.Client.Core;
using Ofcas.Lk.Api.Shared;

namespace LogikalHeadlessExport
{
    // ── Configuration model ──────────────────────────────────────────────────
    // This maps 1-to-1 with the fields in appsettings.json.
    public class AppConfig
    {
        /// <summary>"Discover" to list projects, "Export" to run the export.</summary>
        public string Mode { get; set; } = "Discover";

        /// <summary>Full path to LogikalStarter.exe on this machine.</summary>
        public string LauncherPath { get; set; } = "";

        /// <summary>Login program mode. "erp" is the one used for ERP exports.</summary>
        public string ProgramMode { get; set; } = "erp";

        /// <summary>
        /// GUID of the project to export (only needed in Export mode).
        /// You get this by running in Discover mode first.
        /// </summary>
        public string ProjectGuid { get; set; } = "";

        /// <summary>Folder where the SQLite export file will be saved.</summary>
        public string OutputDirectory { get; set; } = "C:\\LogikalExports";
    }

    public static class Program
    {
        public static async Task<int> Main(string[] args)
        {
            Log("Logikal Headless ERP Export starting...");

            // ── Step 1: Load config ──────────────────────────────────────────
            AppConfig config = LoadConfig();
            if (config == null) return 1;

            Log($"Mode: {config.Mode}");
            Log($"LauncherPath: {config.LauncherPath}");
            Log($"ProgramMode: {config.ProgramMode}");
            if (config.Mode == "Export")
                Log($"ProjectGuid: {config.ProjectGuid}");

            // ── Step 2: Validate config ──────────────────────────────────────
            if (string.IsNullOrWhiteSpace(config.LauncherPath))
            {
                LogError("LauncherPath is empty in appsettings.json. Set it to the full path of LogikalStarter.exe.");
                return 1;
            }

            if (!File.Exists(config.LauncherPath))
            {
                LogError($"LauncherPath not found: {config.LauncherPath}");
                LogError("Make sure you're running this program on the same machine as Logikal.");
                return 1;
            }

            if (config.Mode == "Export")
            {
                if (string.IsNullOrWhiteSpace(config.ProjectGuid))
                {
                    LogError("ProjectGuid is empty. Run in Discover mode first to find your project's GUID.");
                    return 1;
                }

                Directory.CreateDirectory(config.OutputDirectory); // ensure output folder exists
            }

            // ── Step 3: Connect to Logikal ───────────────────────────────────
            // ServiceProxyFactory is the NON-UI version (no dialogs).
            // It creates a background service process that talks to Logikal.
            Log("Creating service proxy (connecting to Logikal)...");

            using (IServiceProxyResult serviceProxyResult = ServiceProxyFactory.CreateServiceProxy(config.LauncherPath))
            {
                IServiceProxy serviceProxy = serviceProxyResult.ServiceProxy;

                // Start the Logikal service host process
                using (IResult startResult = serviceProxy.Start())
                {
                    if (startResult.OperationCode != OperationCode.Accepted)
                    {
                        LogError($"Failed to start Logikal service. OperationCode: {startResult.OperationCode}");
                        return 1;
                    }
                }

                Log("Service started successfully.");

                try
                {
                    // ── Step 4: Login ────────────────────────────────────────
                    // Note: No ApplicationHandle needed because we're not showing any dialogs.
                    var loginParameters = new Dictionary<string, object>
                    {
                        { WellKnownParameterKey.Login.ProgramMode, config.ProgramMode },
                        { WellKnownParameterKey.Login.EnableEventSynchronization, false }
                    };

                    // Check if login is possible before attempting it
                    IOperationInfo canLogin = serviceProxy.CanLogin(loginParameters);
                    if (!canLogin.CanExecute)
                    {
                        LogError($"Cannot login: {canLogin}");
                        return 1;
                    }

                    Log("Logging in...");

                    // Login returns ILoginScope (non-UI) - gives full API access without dialogs
                    using (ICoreObjectResult<ILoginScope> loginResult = serviceProxy.Login(loginParameters))
                    {
                        if (loginResult.OperationCode != OperationCode.Accepted)
                        {
                            LogError($"Login failed. OperationCode: {loginResult.OperationCode}");
                            return 1;
                        }

                        ILoginScope loginScope = loginResult.CoreObject;
                        Log("Logged in successfully.");

                        // ── Step 5: Run the selected mode ────────────────────
                        int exitCode;

                        if (config.Mode == "Discover")
                        {
                            exitCode = RunDiscover(loginScope);
                        }
                        else if (config.Mode == "Export")
                        {
                            exitCode = await RunExport(loginScope, config);
                        }
                        else
                        {
                            LogError($"Unknown Mode '{config.Mode}' in appsettings.json. Use 'Discover' or 'Export'.");
                            exitCode = 1;
                        }

                        // ── Step 6: Logout ───────────────────────────────────
                        IOperationInfo canLogout = loginScope.CanLogout();
                        if (canLogout.CanExecute)
                        {
                            using (IResult logoutResult = loginScope.Logout())
                            {
                                Log(logoutResult.OperationCode == OperationCode.Accepted
                                    ? "Logged out."
                                    : $"Logout returned: {logoutResult.OperationCode}");
                            }
                        }

                        return exitCode;
                    }
                }
                finally
                {
                    // Always stop the service, even if something goes wrong above
                    Log("Stopping Logikal service...");
                    using (IResult stopResult = serviceProxy.Stop())
                    {
                        Log(stopResult.OperationCode == OperationCode.Accepted
                            ? "Service stopped."
                            : $"Stop returned: {stopResult.OperationCode}");
                    }
                }
            }
        }

        // ── DISCOVER MODE ────────────────────────────────────────────────────
        // Lists all recent projects with their GUIDs. Run this first to find
        // the GUID you need for Export mode.
        private static int RunDiscover(ILoginScope loginScope)
        {
            Log("");
            Log("======= DISCOVER MODE =======");
            Log("Listing recent projects...");
            Log("");

            // WellKnownProjectType.Project = 0 (regular offer projects)
            // WellKnownProjectType.FabricationLot = 1 (production lots)
            // Try both types so you see everything
            int[] projectTypes = { WellKnownProjectType.Project, WellKnownProjectType.FabricationLot };
            string[] typeNames = { "Project (Offer)", "FabricationLot (Production)" };

            bool foundAny = false;

            for (int i = 0; i < projectTypes.Length; i++)
            {
                IOperationInfo canGet = loginScope.CanGetRecentProjects(projectTypes[i]);
                if (!canGet.CanExecute)
                {
                    Log($"[{typeNames[i]}] CanGetRecentProjects returned false: {canGet}");
                    continue;
                }

                using (ICoreInfoListResult<IBaseProjectInfo> recentResult = loginScope.GetRecentProjects(projectTypes[i]))
                {
                    if (recentResult.OperationCode != OperationCode.Accepted)
                    {
                        Log($"[{typeNames[i]}] GetRecentProjects failed: {recentResult.OperationCode}");
                        continue;
                    }

                    IList<IBaseProjectInfo> projects = recentResult.CoreInfos;

                    if (projects.Count == 0)
                    {
                        Log($"[{typeNames[i]}] No recent projects found.");
                        continue;
                    }

                    Log($"[{typeNames[i]}] Found {projects.Count} recent project(s):");
                    Log(new string('-', 60));

                    foreach (IBaseProjectInfo proj in projects)
                    {
                        Log($"  Name            : {proj.Name}");
                        Log($"  GUID            : {proj.Guid}");   // <-- copy this into appsettings.json
                        Log($"  PersonInCharge  : {proj.PersonInCharge}");
                        Log($"  Path            : {proj.Path}");
                        Log($"  LastChanged     : {proj.LastChangedDateTime}");
                        Log("");
                        foundAny = true;
                    }
                }
            }

            if (!foundAny)
            {
                Log("No projects found. Things to check:");
                Log("  1. Is the ProgramMode correct? (check with the customer)");
                Log("  2. Has any project been opened in Logikal recently?");
                Log("  3. Is this program running on the same machine as Logikal?");
            }
            else
            {
                Log("======= ACTION REQUIRED =======");
                Log("Copy the GUID of the project you want to export.");
                Log("Then edit appsettings.json:");
                Log("  - Paste the GUID into 'ProjectGuid'");
                Log("  - Change 'Mode' to 'Export'");
                Log("  - Run the program again");
            }

            return 0;
        }

        // ── EXPORT MODE ──────────────────────────────────────────────────────
        // Opens the project by GUID and runs the ERP export, saving the SQLite file.
        private static async Task<int> RunExport(ILoginScope loginScope, AppConfig config)
        {
            Log("");
            Log("======= EXPORT MODE =======");

            // Parse the GUID from config
            if (!Guid.TryParse(config.ProjectGuid, out Guid projectGuid))
            {
                LogError($"ProjectGuid '{config.ProjectGuid}' is not a valid GUID format.");
                LogError("Expected format: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx");
                return 1;
            }

            // ── Open the project by GUID (no UI dialog needed) ───────────────
            Log($"Opening project {projectGuid}...");

            IOperationInfo canGetProject = loginScope.CanGetProjectByGuid(projectGuid);
            if (!canGetProject.CanExecute)
            {
                LogError($"Cannot open project: {canGetProject}");
                return 1;
            }

            using (ICoreObjectHierarchyResult<IProject> projectResult = loginScope.GetProjectByGuid(projectGuid))
            {
                if (projectResult.OperationCode != OperationCode.Accepted)
                {
                    LogError($"GetProjectByGuid failed. OperationCode: {projectResult.OperationCode}");
                    return 1;
                }

                IProject project = projectResult.CoreObject;
                Log("Project opened successfully.");

                // ── Get available reports ────────────────────────────────────
                IOperationInfo canGetReports = project.CanGetReports();
                if (!canGetReports.CanExecute)
                {
                    LogError($"Cannot get reports for this project: {canGetReports}");
                    return 1;
                }

                using (IReportItemsResult reportsResult = project.GetReports())
                {
                    if (reportsResult.OperationCode != OperationCode.Accepted)
                    {
                        LogError($"GetReports failed: {reportsResult.OperationCode}");
                        return 1;
                    }

                    IList<IReportItem> reports = reportsResult.ReportItems;

                    // Log all available reports - useful for debugging
                    Log($"Available reports ({reports.Count}):");
                    foreach (IReportItem r in reports)
                    {
                        Log($"  [{r.Category?.Id}] {r.Id} - {r.Name}");
                    }

                    // ── Find the ERP Export report ───────────────────────────
                    IReportItem erpReport = reports.FirstOrDefault(r =>
                        r.Id == WellKnownReports.Delivery.ErpExport &&
                        r.Category?.Id == WellKnownReports.Delivery.CategoryId);

                    if (erpReport == null)
                    {
                        LogError("ERP Export report not found. Possible reasons:");
                        LogError("  - The project type doesn't support ERP export");
                        LogError("  - The ProgramMode doesn't have ERP export enabled");
                        LogError("  - The Logikal license doesn't include ERP export");
                        return 1;
                    }

                    Log($"Found ERP Export report: {erpReport.Name}");

                    // ── Run the export ───────────────────────────────────────
                    var exportParameters = new Dictionary<string, object>
                    {
                        // SQLite is the only supported format for ERP export
                        { WellKnownParameterKey.Project.Report.ExportFormat, "SQLite" }
                    };

                    // Empty list = export all elevations in the project
                    var allElevations = new List<ICoreInfoReportable>(0);

                    IOperationInfo canRun = project.CanGetReport(erpReport, allElevations, exportParameters);
                    if (!canRun.CanExecute)
                    {
                        LogError($"Cannot run ERP export: {canRun}");
                        return 1;
                    }

                    Log("Running ERP export (this may take a moment)...");

                    // IProject.GetReport() is synchronous - runs in background thread
                    // to avoid blocking the main thread
                    using (IStreamResult streamResult = await Task.Run(() =>
                        project.GetReport(erpReport, allElevations, exportParameters)))
                    {
                        if (streamResult.OperationCode != OperationCode.Accepted)
                        {
                            LogError($"GetReport failed: {streamResult.OperationCode}");
                            return 1;
                        }

                        // ── Save the SQLite file ─────────────────────────────
                        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                        string fileName = $"erp_export_{timestamp}.sqlite3";
                        string outputPath = Path.Combine(config.OutputDirectory, fileName);

                        Log($"Saving export to: {outputPath}");

                        using (FileStream fileStream = File.Create(outputPath))
                        {
                            streamResult.Stream.CopyTo(fileStream);
                        }

                        long fileSizeKb = new FileInfo(outputPath).Length / 1024;
                        Log($"Export saved successfully! File size: {fileSizeKb} KB");
                        Log("");
                        Log("Next steps:");
                        Log("  The SQLite file contains the ERP data.");
                        Log($"  To inspect it, open '{outputPath}' with DB Browser for SQLite.");
                        Log("  (Download free from: https://sqlitebrowser.org/)");
                        Log("  Once you know the table structure, we can add SQLite -> JSON conversion.");

                        return 0;
                    }
                }
            }
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private static AppConfig LoadConfig()
        {
            const string configFile = "appsettings.json";

            if (!File.Exists(configFile))
            {
                LogError($"Config file '{configFile}' not found next to the executable.");
                LogError("Create it with content like:");
                LogError(@"{
  ""Mode"": ""Discover"",
  ""LauncherPath"": ""C:\\Logikal\\LogikalStarter.exe"",
  ""ProgramMode"": ""erp"",
  ""ProjectGuid"": """",
  ""OutputDirectory"": ""C:\\LogikalExports""
}");
                return null;
            }

            try
            {
                string json = File.ReadAllText(configFile);
                var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                AppConfig config = JsonSerializer.Deserialize<AppConfig>(json, options);

                if (config == null)
                {
                    LogError("appsettings.json is empty or invalid JSON.");
                    return null;
                }

                return config;
            }
            catch (JsonException ex)
            {
                LogError($"Failed to parse appsettings.json: {ex.Message}");
                return null;
            }
        }

        private static void Log(string message)
        {
            string line = $"[{DateTime.Now:HH:mm:ss}] {message}";
            Console.WriteLine(line);
        }

        private static void LogError(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ERROR: {message}");
            Console.ResetColor();
        }
    }
}
