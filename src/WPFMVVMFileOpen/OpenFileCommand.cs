using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace WPFMVVMFileOpen
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class OpenFileCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("a88a1f5e-0457-4012-8bd7-8ca7418a78b5");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenFileCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private OpenFileCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static OpenFileCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in OpenFileCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new OpenFileCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var csprojFilePath = FindCsProjectPath();
            FileInfo fileInfo = new FileInfo(csprojFilePath);
            var projectDirectory = fileInfo.Directory.FullName;

            if (!string.IsNullOrEmpty(projectDirectory))
            {
                var activeFileName = GetActivateFileName();

                string targetFile = string.Empty;

                if (activeFileName.EndsWith(".xaml.cs") || activeFileName.EndsWith(".xaml"))
                {
                    activeFileName = activeFileName.Replace(".xaml.cs", string.Empty).Replace(".xaml", string.Empty);

                    activeFileName = activeFileName.EndsWith("View") ? $"{activeFileName}Model" : $"{activeFileName}ViewModel";

                    targetFile = FindRelateViewModelFile(projectDirectory, activeFileName);
                }
                else if (activeFileName.EndsWith(".cs"))
                {
                    activeFileName = activeFileName.Replace(".cs", string.Empty);

                    if (activeFileName.EndsWith("ViewModel"))
                    {
                        activeFileName = activeFileName.Replace("Model", string.Empty);
                        targetFile = FindRelateViewFile(projectDirectory, activeFileName);
                    }
                } 
                
                if (targetFile != string.Empty) 
                { 
                    var context = Package.GetGlobalService(typeof(DTE)) as DTE; 
                    context.ItemOperations.OpenFile(targetFile); 
                } 
                else 
                { 
                    // Show a message box to prove we were here
                    VsShellUtilities.ShowMessageBox( 
                        this.package, 
                        "File not found!", 
                        "Error", 
                        OLEMSGICON.OLEMSGICON_INFO, 
                        OLEMSGBUTTON.OLEMSGBUTTON_OK, 
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                    }
            }
        }

        private string FindRelateViewFile(string directoryPath, string targetFileName)
        {
             string targetFilePath = string.Empty;

            foreach (string dir in Directory.GetFileSystemEntries(directoryPath))
            {
                if (File.Exists(dir))
                {
                    FileInfo fileInfo = new FileInfo(dir);

                    if (fileInfo.Extension == ".xaml")
                    {
                        var currentFileName = fileInfo.Name.Replace(fileInfo.Extension, string.Empty);

                        if (currentFileName == targetFileName)
                        {
                            targetFilePath = dir;
                            break;
                        }
                    }
                }
                else
                {
                    FindRelateViewFile(dir, targetFileName);
                }
            }

            return targetFilePath;
        }

        private string FindRelateViewModelFile(string directoryPath, string targetFileName)
        {
            string targetFilePath = string.Empty;

            foreach (string dir in Directory.GetFileSystemEntries(directoryPath))
            {
                if (File.Exists(dir))
                {
                    FileInfo fileInfo = new FileInfo(dir);

                    if (fileInfo.Extension == ".cs")
                    {
                        var currentFileName = fileInfo.Name.Replace(fileInfo.Extension, string.Empty);

                        if (currentFileName == targetFileName)
                        {
                            targetFilePath = dir;
                            break;
                        }
                    }
                }
                else
                {
                    FindRelateViewModelFile(dir, targetFileName);
                }
            }

            return targetFilePath;
        }

        private string FindCsProjectPath()
        {
            var context = Package.GetGlobalService(typeof(DTE)) as DTE;

            var projectPath = new System.Collections.Generic.List<string>();
            foreach (Project proj in context.Solution.Projects)
            {
                projectPath.Add(proj.FullName);
            }

            var targetFilePath = context.DTE.ActiveDocument.Path;
            var targetProjectPath = string.Empty;

            foreach (var proj in projectPath)
            {
                if (proj.Contains(targetFilePath))
                {
                    targetProjectPath = proj;
                    break;
                }
            }

            return targetProjectPath;
        }

        private string GetActivateFileName()
        {
            var context = Package.GetGlobalService(typeof(DTE)) as DTE;
            var filePath = context.ActiveDocument.FullName;
            FileInfo file = new FileInfo(filePath);
            // return file.Name.Replace(file.Extension, string.Empty);
            return file.Name;
        }
    }
}
