using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace VisualStudioScreenshot
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(VisualStudioScreenshotPackage.PackageGuidString)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class VisualStudioScreenshotPackage : AsyncPackage
    {
        public const string PackageGuidString = "a512387e-6f49-44c9-afe6-bb6db4b32c05";

        public static readonly Guid CommandSet = new Guid("f2f7beed-cbd2-40e2-956a-e533083b9dbc");

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            await CaptureMainWindowCommand.InitializeAsync(this);
            await CaptureActiveWindowCommand.InitializeAsync(this);
        }
    }
}
