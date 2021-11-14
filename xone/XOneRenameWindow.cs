using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using Task = System.Threading.Tasks.Task;

namespace xone
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed partial class XOneRenameWindow
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        const string Command_NewScopedWindow = "ProjectandSolutionContextMenus.Project.SolutionExplorer.NewScopedWindow";


        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("8d754e1a-d1eb-4169-ae24-5dc7ec9fed98");

        /// <summary>
        /// Explorer Object Kind Item
        /// </summary>
        private static readonly string ExplorerObjectKind = "{3AE79031-E1BC-11D0-8F78-00A0C9110057}";

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="XOneRenameWindow"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private XOneRenameWindow(AsyncPackage package, OleMenuCommandService commandService)
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
        public static XOneRenameWindow Instance
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
            // Switch to the main thread - the call to AddCommand in XOneRenameWindow's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new XOneRenameWindow(package, commandService);
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

            var dte = xonePackage.Application;

            dte.ExecuteCommand(Command_NewScopedWindow);

            var lastChar = Char.ConvertFromUtf32(32);

            var windows = dte.Windows.OfType<EnvDTE.Window>().Where(w => w.ObjectKind == ExplorerObjectKind);

            foreach (var explorerWindow in windows)
            {
                try
                {
                    if (explorerWindow.Caption.EndsWith(lastChar)) continue;

                    var itemUi = explorerWindow.Object as EnvDTE.UIHierarchy;

                    var rootItem = itemUi.UIHierarchyItems.GetEnumerator(); rootItem.MoveNext();

                    var caption = String.Join("|", (rootItem.Current as EnvDTE.UIHierarchyItem)?.Name.Split('.').Reverse());

                    explorerWindow.Caption = String.Concat(caption, lastChar);

                    var dockView = explorerWindow.GetType()
                            .GetProperty("DockViewElement", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                            .GetValue(explorerWindow);

                    if (dockView != null)
                    {

                        var dockViewTitle = dockView.GetType()
                                    .GetProperty("Title")
                                    .GetValue(dockView);

                        var dockViewTitleType = dockViewTitle?.GetType();

                        if (dockViewTitleType != null)
                        {
                            foreach (var name in new string[] { "ShortTitle", "Title", "ToolTip" })
                            {
                                dockViewTitleType.GetProperty(name).SetValue(dockViewTitle, caption);
                            }
                        }
                    }

                }
                catch { }

            }


        }
    }
}
