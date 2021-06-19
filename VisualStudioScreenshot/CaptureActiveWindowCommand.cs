using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Platform.WindowManagement;
using Microsoft.VisualStudio.PlatformUI.Shell;
using Microsoft.VisualStudio.PlatformUI.Shell.Controls;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using System;
using System.ComponentModel.Design;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Controls;
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

                WindowFrame windowFrame = vsWindowFrame as WindowFrame;

                View frameView = windowFrame.FrameView;

                if (frameView == null)
                {
                    frameView = windowFrame.RootView;
                }

                Panel contentHostingPanel = frameView.Content as Panel;

                bool isGenericPaneClientHwndHost = false;

                foreach (System.Windows.UIElement uiElement in contentHostingPanel.Children)
                {
                    if (uiElement.GetType().Name == "GenericPaneClientHwndHost")
                    {
                        isGenericPaneClientHwndHost = true;
                    }
                }

                MemoryStream memoryStream;

                if (!isGenericPaneClientHwndHost)
                {
                    memoryStream = this.GetScreenFromRender(contentHostingPanel);
                }
                else
                {
                    memoryStream = this.GetScreenFromGraphics(contentHostingPanel);
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    DefaultExt = ".png",
                    Filter = "PNG|*.png",
                    FileName = "CaptureActiveWindow"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    using (FileStream fileStream = new FileStream(saveFileDialog.FileName, FileMode.Create, FileAccess.Write))
                    {
                        memoryStream.WriteTo(fileStream);
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

        private MemoryStream GetScreenFromRender(System.Windows.FrameworkElement frameworkElement)
        {
            RenderTargetBitmap renderTargetBitmap = new RenderTargetBitmap((int)frameworkElement.ActualWidth, (int)frameworkElement.ActualHeight, 96, 96, PixelFormats.Pbgra32);

            renderTargetBitmap.Render(frameworkElement);

            PngBitmapEncoder pngImage = new PngBitmapEncoder();

            pngImage.Frames.Add(BitmapFrame.Create(renderTargetBitmap));

            MemoryStream memoryStream = new MemoryStream();

            pngImage.Save(memoryStream);

            return memoryStream;
        }

        private MemoryStream GetScreenFromGraphics(System.Windows.FrameworkElement frameworkElement)
        {
            System.Windows.Rect rect = this.GetAbsolutePlacement(frameworkElement, true);

            double windowsScaling = this.GetWindowsScaling();

            Size sizeRect = new Size((int)(frameworkElement.ActualWidth * windowsScaling), (int)(frameworkElement.ActualHeight * windowsScaling));

            MemoryStream memoryStream = new MemoryStream();

            using (Bitmap bitmap = new Bitmap(sizeRect.Width, sizeRect.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
            {

                Graphics graphics = Graphics.FromImage(bitmap);

                graphics.CopyFromScreen((int)rect.Left, (int)rect.Top, 0, 0, sizeRect, CopyPixelOperation.SourceCopy);

                bitmap.Save(memoryStream, ImageFormat.Png);
            }

            return memoryStream;
        }

        private double GetWindowsScaling()
        {
            return System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width / System.Windows.SystemParameters.PrimaryScreenWidth;
        }

        private System.Windows.Rect GetAbsolutePlacement(System.Windows.FrameworkElement frameworkElement, bool relativeToScreen = false)
        {
            var absolutePos = frameworkElement.PointToScreen(new System.Windows.Point(0, 0));

            if (relativeToScreen)
            {
                return new System.Windows.Rect(absolutePos.X, absolutePos.Y, frameworkElement.ActualWidth, frameworkElement.ActualHeight);
            }

            var posMW = System.Windows.Application.Current.MainWindow.PointToScreen(new System.Windows.Point(0, 0));

            absolutePos = new System.Windows.Point(absolutePos.X - posMW.X, absolutePos.Y - posMW.Y);

            return new System.Windows.Rect(absolutePos.X, absolutePos.Y, frameworkElement.ActualWidth, frameworkElement.ActualHeight);
        }
    }
}
