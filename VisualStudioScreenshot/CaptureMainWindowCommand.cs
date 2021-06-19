using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using System;
using System.ComponentModel.Design;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using Task = System.Threading.Tasks.Task;

namespace VisualStudioScreenshot
{
    internal sealed class CaptureMainWindowCommand
    {
        public const int CommandId = 0x0100;

        private readonly AsyncPackage AsyncPackage;

        private CaptureMainWindowCommand(AsyncPackage asyncPackage, OleMenuCommandService oleMenuCommandService)
        {
            this.AsyncPackage = asyncPackage ?? throw new ArgumentNullException(nameof(asyncPackage));

            oleMenuCommandService = oleMenuCommandService ?? throw new ArgumentNullException(nameof(oleMenuCommandService));

            CommandID commandID = new CommandID(VisualStudioScreenshotPackage.CommandSet, CommandId);

            MenuCommand menuCommand = new MenuCommand(this.Execute, commandID);

            oleMenuCommandService.AddCommand(menuCommand);
        }

        public static CaptureMainWindowCommand Instance
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

            Instance = new CaptureMainWindowCommand(asyncPackage, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            string windowTitle = "Visual Studio Screenshot";
            string messageText = "";

            ThreadHelper.ThrowIfNotOnUIThread();

            Microsoft.VisualStudio.Shell.ThreadHelper.JoinableTaskFactory.Run(async delegate
            {
                await Microsoft.VisualStudio.Shell.ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                EnvDTE.DTE dte = Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(EnvDTE.DTE)) as EnvDTE.DTE;

                Window window = dte.MainWindow;

                if (window.HWnd == 0)
                {
                    messageText = "Unable to get HWnd, please try again.";

                    VsShellUtilities.ShowMessageBox(
                        this.AsyncPackage,
                        messageText,
                        windowTitle,
                        OLEMSGICON.OLEMSGICON_WARNING,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

                    return;
                }

                Rectangle rectangle = default;

                GetWindowRect(new IntPtr(window.HWnd), ref rectangle);

                Size sizeRectangle = new Size(rectangle.Width - rectangle.X, rectangle.Height - rectangle.Y);

                Bitmap bitmap = new Bitmap(sizeRectangle.Width, sizeRectangle.Height, PixelFormat.Format32bppArgb);

                Graphics graphics = Graphics.FromImage(bitmap);

                graphics.CopyFromScreen(rectangle.X, rectangle.Y, 0, 0, sizeRectangle, CopyPixelOperation.SourceCopy);

                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    DefaultExt = ".png",
                    Filter = "PNG|*.png",
                    FileName = "CaptureMainWindow"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    bitmap.Save(saveFileDialog.FileName, ImageFormat.Png);

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

        [DllImport("user32.dll")]
        public static extern Boolean GetWindowRect(IntPtr hWnd, ref Rectangle bounds);
    }
}
