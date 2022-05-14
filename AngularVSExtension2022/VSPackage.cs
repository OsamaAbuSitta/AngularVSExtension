using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace AngularVSExtension
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", Vsix.Version, IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(PackageGuids.guidTypeScriptHtmlSwitchCommandPackageString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideService(typeof(Microsoft.VisualStudio.Shell.Interop.SDTE), IsAsyncQueryable = true)]
    //[ProvideAutoLoad("f1536ef8-92ec-443c-9ed7-fdadf150da82")]
    //[SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class VSPackage : AsyncPackage
    {
        //public const string PackageGuidString = "14261379-f70d-4d2a-b69f-653605fb158f";

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            await TypeScriptHtmlSwitchCommand.InitializeAsync(this);
            await NgServeCommand.InitializeAsync(this);
            await OpenInCmdCommand.InitializeAsync(this);
        }
    }
}