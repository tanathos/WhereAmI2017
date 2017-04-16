using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell.Settings;
using System;
using System.ComponentModel;
using System.ComponentModel.Composition;
using System.Drawing;
using System.Runtime.InteropServices;
using WhereAmI2017.Converters;

namespace WhereAmI2017
{
    /// <summary>
    // Extends the standard dialog functionality for implementing ToolsOptions pages, 
    // with support for the Visual Studio automation model, Windows Forms, and state 
    // persistence through the Visual Studio settings mechanism.
    /// </summary>
    [Guid(GuidStrings.GuidPageGeneral)]
    [Export(typeof(IWhereAmISettings))]
    public class OptionsPageGeneral : Microsoft.VisualStudio.Shell.DialogPage, IWhereAmISettings
    {
        /// <summary>
        /// The real store in which the settings will be saved
        /// </summary>
        readonly WritableSettingsStore writableSettingsStore;

        IWhereAmISettings settings
        {
            get
            {
                var componentModel = (IComponentModel)(Site.GetService(typeof(SComponentModel)));
                IWhereAmISettings s = componentModel.DefaultExportProvider.GetExportedValue<IWhereAmISettings>();

                return s;
            }
        }

        #region Constructors

        [ImportingConstructor]
        public OptionsPageGeneral(SVsServiceProvider vsServiceProvider) : this()
        {
            var shellSettingsManager = new ShellSettingsManager(vsServiceProvider);
            writableSettingsStore = shellSettingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);

            LoadSettings();
        }

        public OptionsPageGeneral()
        {
            FilenameSize = 60;
            FoldersSize = ProjectSize = 52;

            FilenameColor = Color.FromArgb(234, 234, 234);
            FoldersColor = ProjectColor = Color.FromArgb(243, 243, 243);

            ViewFilename = ViewFolders = ViewProject = true;

            Opacity = 1;
        }

        #endregion

        #region Properties

        [Category("File name")]
        [Description("The color of the filename part")]
        [DisplayName("Filename color")]
        public Color FilenameColor { get; set; }

        [Category("File name")]
        [Description("Indicate to show the filename part or not")]
        [DisplayName("Show")]
        public bool ViewFilename { get; set; }

        [Category("Folder")]
        [Description("The color of the folder part")]
        [DisplayName("Folder color")]
        public Color FoldersColor { get; set; }

        [Category("Folder")]
        [Description("Indicate to show the folder part or not")]
        [DisplayName("Show")]
        public bool ViewFolders { get; set; }

        [Category("Project")]
        [Description("The color of the project part")]
        [DisplayName("Project color")]
        public Color ProjectColor { get; set; }

        [Category("Project")]
        [Description("Indicate to show the project part or not")]
        [DisplayName("Show")]
        public bool ViewProject { get; set; }

        [Category("File name")]
        [Description("The size of the filename part")]
        [DisplayName("Filename size")]
        public double FilenameSize { get; set; }

        [Category("Folder")]
        [Description("The size of the folder part")]
        [DisplayName("Folder size")]
        public double FoldersSize { get; set; }

        [Category("Project")]
        [Description("The size of the project part")]
        [DisplayName("Project size")]
        public double ProjectSize { get; set; }

        [Category("Appearance")]
        [DisplayName("Opacity")]
        [Description("Opacity of the text. Insert a value between 0 and 1.")]
        [TypeConverter(typeof(PercentageConverter))]
        public double Opacity { get; set; }

        // TODO
        public AdornmentPositions Position { get; set; }

        #endregion Properties

        #region Event Handlers

        /// <summary>
        /// Handles "activate" messages from the Visual Studio environment.
        /// </summary>
        /// <devdoc>
        /// This method is called when Visual Studio wants to activate this page.  
        /// </devdoc>
        /// <remarks>If this handler sets e.Cancel to true, the activation will not occur.</remarks>
        protected override void OnActivate(CancelEventArgs e)
        {
            int result = VsShellUtilities.ShowMessageBox(Site, "Resources.MessageOnActivateEntered", null /*title*/, OLEMSGICON.OLEMSGICON_QUERY, OLEMSGBUTTON.OLEMSGBUTTON_OKCANCEL, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

            if (result == (int)VSConstants.MessageBoxResult.IDCANCEL)
            {
                e.Cancel = true;
            }

            base.OnActivate(e);

            BindSettings();
        }

        /// <summary>
        /// Handles "close" messages from the Visual Studio environment.
        /// </summary>
        /// <devdoc>
        /// This event is raised when the page is closed.
        /// </devdoc>
        protected override void OnClosed(EventArgs e)
        {
            VsShellUtilities.ShowMessageBox(Site, "Resources.MessageOnClosed", null /*title*/, OLEMSGICON.OLEMSGICON_INFO, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        /// <summary>
        /// Handles "deactivate" messages from the Visual Studio environment.
        /// </summary>
        /// <devdoc>
        /// This method is called when VS wants to deactivate this
        /// page.  If this handler sets e.Cancel, the deactivation will not occur.
        /// </devdoc>
        /// <remarks>
        /// A "deactivate" message is sent when focus changes to a different page in
        /// the dialog.
        /// </remarks>
        protected override void OnDeactivate(CancelEventArgs e)
        {
            int result = VsShellUtilities.ShowMessageBox(Site, "Resources.MessageOnDeactivateEntered", null /*title*/, OLEMSGICON.OLEMSGICON_QUERY, OLEMSGBUTTON.OLEMSGBUTTON_OKCANCEL, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

            if (result == (int)VSConstants.MessageBoxResult.IDCANCEL)
            {
                e.Cancel = true;
            }
        }

        /// <summary>
        /// Handles "apply" messages from the Visual Studio environment.
        /// </summary>
        /// <devdoc>
        /// This method is called when VS wants to save the user's 
        /// changes (for example, when the user clicks OK in the dialog).
        /// </devdoc>
        protected override void OnApply(PageApplyEventArgs e)
        {
            int result = VsShellUtilities.ShowMessageBox(Site, "Resources.MessageOnApplyEntered", null /*title*/, OLEMSGICON.OLEMSGICON_QUERY, OLEMSGBUTTON.OLEMSGBUTTON_OKCANCEL, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

            if (result == (int)VSConstants.MessageBoxResult.IDCANCEL)
            {
                e.ApplyBehavior = ApplyKind.Cancel;
            }
            else
            {
                base.OnApply(e);
            }

            VsShellUtilities.ShowMessageBox(Site, "Resources.MessageOnApply", null /*title*/, OLEMSGICON.OLEMSGICON_INFO, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        public void Store()
        {
            throw new NotImplementedException();
        }

        public void Defaults()
        {
            throw new NotImplementedException();
        }

        #endregion Event Handlers

        private void LoadSettings()
        {
            // Default values
            FilenameSize = 60;
            FoldersSize = ProjectSize = 52;

            FilenameColor = Color.FromArgb(234, 234, 234);
            FoldersColor = ProjectColor = Color.FromArgb(243, 243, 243);

            Position = AdornmentPositions.TopRight;

            Opacity = 1;

            //try
            //{
            //    // Retrieve the Id of the current theme used in VS from user's settings, this is changed a lot in VS2015
            //    string visualStudioThemeId = VSRegistry.RegistryRoot(Microsoft.VisualStudio.Shell.Interop.__VsLocalRegistryType.RegType_UserSettings).OpenSubKey("ApplicationPrivateSettings").OpenSubKey("Microsoft").OpenSubKey("VisualStudio").GetValue("ColorTheme", "de3dbbcd-f642-433c-8353-8f1df4370aba", Microsoft.Win32.RegistryValueOptions.DoNotExpandEnvironmentNames).ToString();

            //    string parsedThemeId = Guid.Parse(visualStudioThemeId.Split('*')[2]).ToString();

            //    switch (parsedThemeId)
            //    {
            //        case "de3dbbcd-f642-433c-8353-8f1df4370aba": // Light
            //        case "a4d6a176-b948-4b29-8c66-53c97a1ed7d0": // Blue
            //        default:
            //            // Just use the defaults
            //            break;

            //        case "1ded0138-47ce-435e-84ef-9ec1f439b749": // Dark
            //            _FilenameColor = Color.FromArgb(48, 48, 48);
            //            _FoldersColor = _ProjectColor = Color.FromArgb(40, 40, 40);
            //            break;
            //    }

            //    // Tries to retrieve the configurations if previously saved
            //    if (writableSettingsStore.PropertyExists(CollectionPath, "FilenameColor"))
            //    {
            //        this.FilenameColor = Color.FromArgb(writableSettingsStore.GetInt32(CollectionPath, "FilenameColor", this.FilenameColor.ToArgb()));
            //    }

            //    if (writableSettingsStore.PropertyExists(CollectionPath, "FoldersColor"))
            //    {
            //        this.FoldersColor = Color.FromArgb(writableSettingsStore.GetInt32(CollectionPath, "FoldersColor", this.FoldersColor.ToArgb()));
            //    }

            //    if (writableSettingsStore.PropertyExists(CollectionPath, "ProjectColor"))
            //    {
            //        this.ProjectColor = Color.FromArgb(writableSettingsStore.GetInt32(CollectionPath, "ProjectColor", this.ProjectColor.ToArgb()));
            //    }

            //    if (writableSettingsStore.PropertyExists(CollectionPath, "ViewFilename"))
            //    {
            //        bool b = this.ViewFilename;
            //        if (Boolean.TryParse(writableSettingsStore.GetString(CollectionPath, "ViewFilename"), out b))
            //            this.ViewFilename = b;
            //    }

            //    if (writableSettingsStore.PropertyExists(CollectionPath, "ViewFolders"))
            //    {
            //        bool b = this.ViewFolders;
            //        if (Boolean.TryParse(writableSettingsStore.GetString(CollectionPath, "ViewFolders"), out b))
            //            this.ViewFolders = b;
            //    }

            //    if (writableSettingsStore.PropertyExists(CollectionPath, "ViewProject"))
            //    {
            //        bool b = this.ViewProject;
            //        if (Boolean.TryParse(writableSettingsStore.GetString(CollectionPath, "ViewProject"), out b))
            //            this.ViewProject = b;
            //    }

            //    if (writableSettingsStore.PropertyExists(CollectionPath, "FilenameSize"))
            //    {
            //        double d = this.FilenameSize;
            //        if (Double.TryParse(writableSettingsStore.GetString(CollectionPath, "FilenameSize"), out d))
            //            this.FilenameSize = d;
            //    }

            //    if (writableSettingsStore.PropertyExists(CollectionPath, "FoldersSize"))
            //    {
            //        double d = this.FoldersSize;
            //        if (Double.TryParse(writableSettingsStore.GetString(CollectionPath, "FoldersSize"), out d))
            //            this.FoldersSize = d;
            //    }

            //    if (writableSettingsStore.PropertyExists(CollectionPath, "ProjectSize"))
            //    {
            //        double d = this.ProjectSize;
            //        if (Double.TryParse(writableSettingsStore.GetString(CollectionPath, "ProjectSize"), out d))
            //            this.ProjectSize = d;
            //    }

            //    if (writableSettingsStore.PropertyExists(CollectionPath, "Position"))
            //    {
            //        AdornmentPositions p = this.Position;
            //        if (Enum.TryParse<AdornmentPositions>(writableSettingsStore.GetString(CollectionPath, "Position"), out p))
            //            this.Position = p;
            //    }

            //    if (writableSettingsStore.PropertyExists(CollectionPath, "Opacity"))
            //    {
            //        double d = this.Opacity;
            //        if (Double.TryParse(writableSettingsStore.GetString(CollectionPath, "Opacity"), out d))
            //            this.Opacity = d;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    Debug.Fail(ex.Message);
            //}
        }

        private void BindSettings()
        {
            FilenameColor = settings.FilenameColor;
            FoldersColor = settings.FoldersColor;
            ProjectColor = settings.ProjectColor;

            FilenameSize = settings.FilenameSize;
            FoldersSize = settings.FoldersSize;
            ProjectSize = settings.ProjectSize;

            ViewFilename = settings.ViewFilename;
            ViewFolders = settings.ViewFolders;
            ViewProject = settings.ViewProject;

            Position = settings.Position;
            Opacity = settings.Opacity;
        }
    }
}
