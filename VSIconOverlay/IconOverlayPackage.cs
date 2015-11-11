using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;
using System.Windows.Shell;

namespace Microsoft.VisualStudio.Samples.IconOverlay
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(IconOverlayPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad(UIContextGuids.DesignMode)]
    public sealed class IconOverlayPackage : Package
    {
        /// <summary>
        /// VSPackage1 GUID string.
        /// </summary>
        public const string PackageGuidString = "ad62938c-530e-4d31-a08a-445b752a21f8";

        /// <summary>
        /// Initializes a new instance of the <see cref="IconOverlayPackage"/> class.
        /// </summary>
        public IconOverlayPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();

            object value = null;
            var shell = (IVsShell)GetService(typeof(SVsShell));


            int version = 0;
            if (shell.GetProperty((int)__VSSPROPID5.VSSPROPID_ReleaseVersion, out value) == 0)
            {
                version = int.Parse(value.ToString().Split('.')[0]);
            }

            // convert internal version numbers to external product names
            string overlay;
            switch (version)
            {
                case 10:
                    overlay = "2010";
                    break;
                case 11:
                    overlay = "2012";
                    break;
                case 12:
                    overlay = "2013";
                    break;
                case 14:
                    overlay = "2015";
                    break;
                default:
                    overlay = "??";
                    break;
            }

            AddOverlay(overlay, Brushes.Black);
        }

        private void AddOverlay(string text, Brush color)
        {
            var font = new FontFamily("Segoe UI");
            FormattedText formattedText = new FormattedText(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
                  new Typeface(font, FontStyles.Normal, FontWeights.Bold, FontStretches.Normal), 14, color);
            
            Drawing textDrawing = new GeometryDrawing(color,
                new Pen(color, 1), formattedText.BuildGeometry(new Point(0, 0)));
            Drawing backgroundDrawing = new GeometryDrawing(Brushes.White, new Pen(Brushes.White, 3), new RectangleGeometry(textDrawing.Bounds));
            DrawingGroup overlayDrawing = new DrawingGroup();
            overlayDrawing.Children.Add(backgroundDrawing);
            overlayDrawing.Children.Add(textDrawing);
            
            TaskbarItemInfo taskbarItem = new TaskbarItemInfo();
            taskbarItem.Overlay = new DrawingImage(overlayDrawing);
            Application.Current.MainWindow.TaskbarItemInfo = taskbarItem;
        }

    }
}
