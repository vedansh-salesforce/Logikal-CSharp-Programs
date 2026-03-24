using Ofcas.Lk.Api.Client.Core;
using Ofcas.Lk.Api.Client.Ui;
using Ofcas.Lk.Api.Shared;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace Ofcas.Lk.Api.Sample.ProjectERPExport
{
    public static class Programm
    {
        // Windows helper method to acquire required window handle
        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();

        public static async Task Main()
        {
            // To display UI dialogs via the API a window handle of the client application is necessary
            // This enables the application to know where to display the UI dialogs
            IntPtr applicationHandle = GetConsoleWindow();

            Console.WriteLine("Please enter the path to the launcher executable, e.g. \"C:\\Logikal\\LogikalStarter.exe\"");
            Console.Write("Launcher Path: ");
            string launcherPath = Console.ReadLine();
            launcherPath = launcherPath.Replace("\"", "");

            Console.WriteLine("Please enter the progmode to run as, e.g. \"erp\"");
            Console.Write("Program Mode: ");
            string programMode = Console.ReadLine();

            // Initialize API connection by providing the path to the starter executable (launcher)
            using (IServiceProxyUiResult serviceProxyUiResult = ServiceProxyUiFactory.CreateServiceProxy(launcherPath))
            {
                // The lifetime of API objects is tied to the returned result objects
                // Using the 'using' statement ensures automatic disposal of result objects upon exiting the code block,
                // while alternatively, the result object can be manually freed by invoking the Dispose() method of the result object
                // This ensures that memory leaks are prevented and all API objects are cleaned up correctly
                // See: https://learn.microsoft.com/en-us/dotnet/csharp/language-reference/statements/using

                // Get service proxy and start a new api service host process
                IServiceProxyUi serviceProxy = serviceProxyUiResult.ServiceProxyUi;
                using (IResult result = serviceProxy.Start())
                {
                    // Check if method call was successfully processed by the API
                    // and whether user accepted potentially shown dialogs
                    if (result.OperationCode != OperationCode.Accepted)
                    {
                        return;
                    }
                }

                try
                {
                    // Login to API V3 with minimal required parameters
                    var parameters = new Dictionary<string, object>
                    {
                        { WellKnownParameterKey.Login.ProgramMode, programMode },
                        { WellKnownParameterKey.Login.ApplicationHandle, applicationHandle },
                        { WellKnownParameterKey.Login.EnableEventSynchronization, false }
                    };

                    // Check if login is possible with given parameters
                    IOperationInfo operationInfo = serviceProxy.CanLogin(parameters);
                    if (!operationInfo.CanExecute)
                    {
                        Console.WriteLine(operationInfo.ToString());
                        return;
                    }

                    // Login with given parameters
                    using (ICoreObjectResult<ILoginScopeUi> loginScopeResult = serviceProxy.Login(parameters))
                    {
                        if (loginScopeResult.OperationCode != OperationCode.Accepted)
                        {
                            return;
                        }

                        ILoginScopeUi loginScope = loginScopeResult.CoreObject;

                        // Check if project selection is allowed
                        operationInfo = loginScope.CanSelectProject();
                        if (!operationInfo.CanExecute)
                        {
                            Console.WriteLine(operationInfo.ToString());
                            return;
                        }

                        // Select a project for erp export by opening the UI project selection
                        using (ICoreObjectHierarchyResult<IProjectUi> projectResult = loginScope.SelectProject())
                        {
                            // The SelectProject method returns a hierarchy result object
                            // This ensures that all API objects the project depends on are
                            // provided and also freed when the result object is freed

                            if (projectResult.OperationCode != OperationCode.Accepted)
                            {
                                return;
                            }

                            IProjectUi project = projectResult.CoreObject;

                            // Check if it's possible to query reports 
                            operationInfo = project.CanGetReports();
                            if (!operationInfo.CanExecute)
                            {
                                Console.WriteLine(operationInfo.ToString());
                                return;
                            }

                            // Query available reports
                            using (IReportItemsResult reportItemsResult = project.GetReports())
                            {
                                if (reportItemsResult.OperationCode != OperationCode.Accepted)
                                {
                                    return;
                                }

                                IList<IReportItem> reportItems = reportItemsResult.ReportItems;

                                // Filter available reports for the erp export report item
                                IReportItem reportItem = reportItems.First(rep =>
                                    (rep.Id == WellKnownReports.Delivery.ErpExport) &&
                                    (rep.Category.Id == WellKnownReports.Delivery.CategoryId));

                                // Create parameters for erp export, export format is required, but always sqlite
                                var exportParameters = new Dictionary<string, object>
                                {
                                    { WellKnownParameterKey.Project.Report.ExportFormat, "SQLite" },
                                };

                                // An empty list is equal to export all valid elevations
                                // Alternatively a list of IElevationInfo can be provided to select specific elevations
                                var reportAllElevations = new List<ICoreInfoReportable>(0);

                                // Check if report can be exported for the given parameters
                                operationInfo = project.CanGetReport(reportItem, reportAllElevations, exportParameters);
                                if (!operationInfo.CanExecute)
                                {
                                    Console.WriteLine(operationInfo.ToString());
                                    return;
                                }

                                // Run report creation asynchronously - begin method starts the operation in background task
                                using (ISynchronizedOperationResult synchronizedOperationResult =
                                    project.BeginGetReport(reportItem, reportAllElevations, exportParameters))
                                {
                                    // End method waits for the background operation to complete in separate task
                                    using (IStreamResult streamResult = await Task.Run(() =>
                                        project.EndGetReport(synchronizedOperationResult.SynchronizedOperation)))
                                    {
                                        if (streamResult.OperationCode != OperationCode.Accepted)
                                        {
                                            return;
                                        }

                                        Stream exportStream = streamResult.Stream;

                                        // Save the file via the api save file dialog
                                        IFileHandlerUi fileHandler = loginScope.FileHandler;

                                        // Select file to export into
                                        var saveFileDialogParameters = new Dictionary<string, object>
                                        {
                                            { WellKnownParameterKey.LoginScope.FileDialog.Filter, "Sqlite3 Datenbank|*.sqlite3" },
                                            { WellKnownParameterKey.LoginScope.FileDialog.InitialFileName, "erp_export" },
                                            { WellKnownParameterKey.LoginScope.FileDialog.ShowMyComputer, true },
                                            { WellKnownParameterKey.LoginScope.FileDialog.ShowDesktop, true },
                                            { WellKnownParameterKey.LoginScope.FileDialog.ShowDocuments, true },
                                        };

                                        // Check if file save dialog can be displayed for the given parameters
                                        operationInfo = fileHandler.CanShowSaveFileDialog(saveFileDialogParameters);
                                        if (!operationInfo.CanExecute)
                                        {
                                            Console.WriteLine(operationInfo.ToString());
                                            return;
                                        }

                                        // Show file save dialog for the given parameters
                                        using (IFileDataResult fileDataResult = fileHandler.ShowSaveFileDialog(saveFileDialogParameters))
                                        {
                                            if (fileDataResult.OperationCode != OperationCode.Accepted)
                                            {
                                                return;
                                            }

                                            IFileData fileData = fileDataResult.FileData;

                                            // Write erp export data into file
                                            using (FileStream fileStream = File.Create(fileData.Path))
                                            {
                                                exportStream.CopyTo(fileStream);
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // Check if all pending operations are finished and logout is possible savely
                        operationInfo = loginScope.CanLogout();
                        if (!operationInfo.CanExecute)
                        {
                            Console.WriteLine(operationInfo.ToString());
                            return;
                        }

                        // Logout user from API
                        using (IResult logoutResult = loginScope.Logout())
                        {
                            if (logoutResult.OperationCode != OperationCode.Accepted)
                            {
                                // Handle logout failures
                            }
                        }
                    }
                }
                finally
                {
                    // Shutdown api service host
                    using (IResult stopServiceResult = serviceProxy.Stop())
                    {
                        if (stopServiceResult.OperationCode != OperationCode.Accepted)
                        {
                            // Handle shutdown failures
                        }
                    }
                }
            }
        }
    }
}