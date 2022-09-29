using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.IO;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;

namespace AngularVSExtension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class TypeScriptHtmlSwitchCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("fde67d78-b48b-4d07-9643-af7a2e1cab14");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="TypeScriptHtmlSwitchCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private TypeScriptHtmlSwitchCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new OleMenuCommand(this.Execute, (s, arg) => { }, MenuItem_BeforeQueryStatus, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static string ActiveDocumentFullPath { get; private set; } = string.Empty;

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static TypeScriptHtmlSwitchCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.AsyncPackage ServiceProvider
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
            // Switch to the main thread - the call to AddCommand in TypeScriptHtmlSwitchCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            var dte = package.GetServiceAsync(typeof(EnvDTE.DTE)).Result as EnvDTE.DTE;
            dte.Events.WindowEvents.WindowActivated += WindowEvents_WindowActivated; ;

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new TypeScriptHtmlSwitchCommand(package, commandService);
        }

        private static void WindowEvents_WindowActivated(EnvDTE.Window gotFocus, EnvDTE.Window lostFocus)
        {
            try
            {
                if (gotFocus.Kind == "Document")
                {
                    Document textDoc = gotFocus.Document as Document;
                    ActiveDocumentFullPath = textDoc.FullName;
                }
                else
                {
                    ActiveDocumentFullPath = string.Empty;
                }
            }
            catch (Exception ex)
            {
                ActivityLog.LogError($"[Angular Html TS Switcher]{nameof(TypeScriptHtmlSwitchCommand)}", ex.ToString());
            }
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private async void Execute(object sender, EventArgs e)
        {
            try
            {

                var selectedFile = string.IsNullOrEmpty(ActiveDocumentFullPath) ? ExtensionHelper.GetSelectedFile(ServiceProvider) : ActiveDocumentFullPath;
                var redirecExtension = Path.GetExtension(selectedFile);
                var redirecExtensionPath = Path.GetDirectoryName(selectedFile) + "\\" + Path.GetFileNameWithoutExtension(selectedFile) +
                    (Path.GetExtension(selectedFile).ToLower() == ".ts" ? ".html" : ".ts");

                if (File.Exists(redirecExtensionPath))
                    VsShellUtilities.OpenDocument(this.package, redirecExtensionPath);
            }
            catch (Exception ex)
            {
                ActivityLog.LogError($"[Angular Html TS Switcher]{nameof(TypeScriptHtmlSwitchCommand)}", ex.ToString());
                MessageBox.Show($"Some error in Angular Html TS Switcher extensiton.\n Please take screenshot and create issue on github with this error\n{ex}", $"[Angular Html TS Switcher]:{nameof(TypeScriptHtmlSwitchCommand)} Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void MenuItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            try
            {
                var command = sender as OleMenuCommand;
                command.Visible = false;
                var selectedFile = string.IsNullOrEmpty(ActiveDocumentFullPath) ? ExtensionHelper.GetSelectedFile(ServiceProvider) : ActiveDocumentFullPath;
                command.Visible = command.Enabled = string.IsNullOrEmpty(selectedFile) == false &&
                       selectedFile.EndsWith(".ts") || selectedFile.EndsWith(".html");
            }
            catch (Exception ex)
            {
                ActivityLog.LogError($"[Angular Html TS Switcher]{nameof(TypeScriptHtmlSwitchCommand)}", ex.ToString());
            }
        }
    }
}