using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Platform.WindowManagement;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using System;
using System.ComponentModel.Design;
using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Task = System.Threading.Tasks.Task;

namespace VisualStudioScreenshot
{
    internal sealed class CaptureActiveWindowCommand
    {
        public const int CommandId = 0x0200;

        private readonly AsyncPackage AsyncPackage;

        private CaptureActiveWindowCommand(AsyncPackage asyncPackage, OleMenuCommandService oleMenuCommandService)
        {
            this.AsyncPackage = asyncPackage ?? throw new ArgumentNullException(nameof(asyncPackage));

            oleMenuCommandService = oleMenuCommandService ?? throw new ArgumentNullException(nameof(oleMenuCommandService));

            CommandID commandID = new CommandID(VisualStudioScreenshotPackage.CommandSet, CommandId);

            MenuCommand menuCommand = new MenuCommand(this.Execute, commandID);

            oleMenuCommandService.AddCommand(menuCommand);
        }

        public static CaptureActiveWindowCommand Instance
        {
            get; private set;
        }

        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get { return this.AsyncPackage; }
        }

        public static async Task InitializeAsync(AsyncPackage asyncPackage)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(asyncPackage.DisposalToken);

            OleMenuCommandService commandService = await asyncPackage.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;

            Instance = new CaptureActiveWindowCommand(asyncPackage, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            string windowTitle = "Visual Studio Screenshot";
            string messageText = "";

            ThreadHelper.ThrowIfNotOnUIThread();

            Microsoft.VisualStudio.Shell.ThreadHelper.JoinableTaskFactory.Run(async delegate
            {
                await Microsoft.VisualStudio.Shell.ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                IVsMonitorSelection selection = Package.GetGlobalService(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;

                object frameObj = null;

                ErrorHandler.ThrowOnFailure(selection.GetCurrentElementValue((uint)VSConstants.VSSELELEMID.SEID_WindowFrame, out frameObj));

                if (frameObj == null)
                {
                    messageText = "Captrue failure, No active window found.";

                    VsShellUtilities.ShowMessageBox(
                        this.AsyncPackage,
                        messageText,
                        windowTitle,
                        OLEMSGICON.OLEMSGICON_WARNING,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                }

                IVsWindowFrame vsWindowFrame = frameObj as IVsWindowFrame;                

                Window window = VsShellUtilities.GetWindowObject(vsWindowFrame);

                WindowFrame windowFrame = vsWindowFrame as WindowFrame;

                System.Windows.FrameworkElement frameworkElement = windowFrame.FrameView.Content as System.Windows.FrameworkElement;

                RenderTargetBitmap renderTargetBitmap = new RenderTargetBitmap((int)frameworkElement.ActualWidth, (int)frameworkElement.ActualHeight, 96, 96, PixelFormats.Pbgra32);

                renderTargetBitmap.Render(frameworkElement);

                PngBitmapEncoder pngImage = new PngBitmapEncoder();

                pngImage.Frames.Add(BitmapFrame.Create(renderTargetBitmap));

                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    DefaultExt = ".png",
                    Filter = "PNG|*.png",
                    FileName = "CaptureActiveWindow"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    using (Stream fileStream = File.Create(saveFileDialog.FileName))
                    {
                        pngImage.Save(fileStream);
                    }

                    messageText = "Captrue image file was saved.";

                    VsShellUtilities.ShowMessageBox(
                        this.AsyncPackage,
                        messageText,
                        windowTitle,
                        OLEMSGICON.OLEMSGICON_WARNING,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                }
            });
        }
    }
}
