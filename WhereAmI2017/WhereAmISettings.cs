using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Settings;
using System;
using System.ComponentModel.Composition;
using System.Diagnostics;
using System.Drawing;

namespace WhereAmI2017
{
  [Export(typeof(IWhereAmISettings))]
  public class WhereAmISettings : IWhereAmISettings
  {
	/// <summary>
	/// The real store in which the settings will be saved
	/// </summary>
	readonly WritableSettingsStore writableSettingsStore;

	public WhereAmISettings()
	{

	}

	[ImportingConstructor]
	public WhereAmISettings(SVsServiceProvider vsServiceProvider) : this()
	{
	  var shellSettingsManager = new ShellSettingsManager(vsServiceProvider);
	  writableSettingsStore = shellSettingsManager.GetWritableSettingsStore(SettingsScope.Remote);

	  LoadSettings();
	}

	public Color FilenameColor { get { return _FilenameColor; } set { _FilenameColor = value; } }
	private Color _FilenameColor;

	public Color FoldersColor { get { return _FoldersColor; } set { _FoldersColor = value; } }
	private Color _FoldersColor;

	public Color ProjectColor { get { return _ProjectColor; } set { _ProjectColor = value; } }
	private Color _ProjectColor;

	public bool ViewFilename { get { return _ViewFilename; } set { _ViewFilename = value; } }
	private bool _ViewFilename = true;

	public bool ViewFolders { get { return _ViewFolders; } set { _ViewFolders = value; } }
	private bool _ViewFolders = true;

	public bool ViewProject { get { return _ViewProject; } set { _ViewProject = value; } }
	private bool _ViewProject = true;

	public double FilenameSize { get { return _FilenameSize; } set { _FilenameSize = value; } }
	private double _FilenameSize;

	public double FoldersSize { get { return _FoldersSize; } set { _FoldersSize = value; } }
	private double _FoldersSize;

	public double ProjectSize { get { return _ProjectSize; } set { _ProjectSize = value; } }
	private double _ProjectSize;

	public AdornmentPositions Position { get { return _Position; } set { _Position = value; } }
	private AdornmentPositions _Position;

	public double Opacity { get { return _Opacity; } set { _Opacity = value; } }
	private double _Opacity;

	public Theme Theme { get { return _Theme; } set { _Theme = value; } }
	private Theme _Theme;

	public void Store()
	{
	  try
	  {
		if (!writableSettingsStore.CollectionExists(Constants.SettingsCollectionPath))
		{
		  writableSettingsStore.CreateCollection(Constants.SettingsCollectionPath);
		}

		writableSettingsStore.SetInt32(Constants.SettingsCollectionPath, "FilenameColor", this.FilenameColor.ToArgb());
		writableSettingsStore.SetInt32(Constants.SettingsCollectionPath, "FoldersColor", this.FoldersColor.ToArgb());
		writableSettingsStore.SetInt32(Constants.SettingsCollectionPath, "ProjectColor", this.ProjectColor.ToArgb());

		writableSettingsStore.SetString(Constants.SettingsCollectionPath, "ViewFilename", this.ViewFilename.ToString());
		writableSettingsStore.SetString(Constants.SettingsCollectionPath, "ViewFolders", this.ViewFolders.ToString());
		writableSettingsStore.SetString(Constants.SettingsCollectionPath, "ViewProject", this.ViewProject.ToString());

		writableSettingsStore.SetString(Constants.SettingsCollectionPath, "FilenameSize", this.FilenameSize.ToString());
		writableSettingsStore.SetString(Constants.SettingsCollectionPath, "FoldersSize", this.FoldersSize.ToString());
		writableSettingsStore.SetString(Constants.SettingsCollectionPath, "ProjectSize", this.ProjectSize.ToString());

		writableSettingsStore.SetString(Constants.SettingsCollectionPath, "Position", this.Position.ToString());
		writableSettingsStore.SetString(Constants.SettingsCollectionPath, "Opacity", this.Opacity.ToString());
		writableSettingsStore.SetString(Constants.SettingsCollectionPath, "Theme", this.Theme.ToString());
	  }
	  catch (Exception ex)
	  {
		Debug.Fail(ex.Message);
	  }
	}

	public void Defaults()
	{
	  writableSettingsStore.DeleteCollection(Constants.SettingsCollectionPath);
	  LoadSettings();
	}

	private void LoadSettings()
	{
	  // Default values
	  var lightTheme = WhereAmISettings.LightThemeDefaults();
	  _FilenameSize = lightTheme.FilenameSize;
	  _FoldersSize = _ProjectSize = lightTheme.FoldersSize;

	  _FilenameColor = lightTheme.FilenameColor;
	  _FoldersColor = _ProjectColor = lightTheme.FoldersColor;

	  _Position = lightTheme.Position;

	  _Opacity = lightTheme.Opacity;
	  _ViewFilename = _ViewFolders = _ViewProject = lightTheme.ViewFilename;
	  _Theme = lightTheme.Theme;

	  try
	  {
		// Retrieve the Id of the current theme used in VS from user's settings, this is changed a lot in VS2015
		string visualStudioThemeId = VSRegistry.RegistryRoot(Microsoft.VisualStudio.Shell.Interop.__VsLocalRegistryType.RegType_UserSettings).OpenSubKey("ApplicationPrivateSettings").OpenSubKey("Microsoft").OpenSubKey("VisualStudio").GetValue("ColorTheme", "de3dbbcd-f642-433c-8353-8f1df4370aba", Microsoft.Win32.RegistryValueOptions.DoNotExpandEnvironmentNames).ToString();

		string parsedThemeId = Guid.Parse(visualStudioThemeId.Split('*')[2]).ToString();

		switch (parsedThemeId)
		{
		  case Constants.VisualStudioLightThemeId: // Light
		  case Constants.VisualStudioBlueThemeId: // Blue
		  default:
			// Just use the defaults
			break;

		  case Constants.VisualStudioDarkThemeId: // Dark
			var darkTheme = WhereAmISettings.DarkThemeDefaults();
			_FilenameSize = darkTheme.FilenameSize;
			_FoldersSize = _ProjectSize = darkTheme.FoldersSize;

			_FilenameColor = darkTheme.FilenameColor;
			_FoldersColor = _ProjectColor = darkTheme.FoldersColor;

			_Position = darkTheme.Position;

			_Opacity = darkTheme.Opacity;
			_ViewFilename = _ViewFolders = _ViewProject = darkTheme.ViewFilename;
			_Theme = darkTheme.Theme;
			break;
		}

		// Tries to retrieve the configurations if previously saved
		if (writableSettingsStore.PropertyExists(Constants.SettingsCollectionPath, "FilenameColor"))
		{
		  this.FilenameColor = Color.FromArgb(writableSettingsStore.GetInt32(Constants.SettingsCollectionPath, "FilenameColor", this.FilenameColor.ToArgb()));
		}

		if (writableSettingsStore.PropertyExists(Constants.SettingsCollectionPath, "FoldersColor"))
		{
		  this.FoldersColor = Color.FromArgb(writableSettingsStore.GetInt32(Constants.SettingsCollectionPath, "FoldersColor", this.FoldersColor.ToArgb()));
		}

		if (writableSettingsStore.PropertyExists(Constants.SettingsCollectionPath, "ProjectColor"))
		{
		  this.ProjectColor = Color.FromArgb(writableSettingsStore.GetInt32(Constants.SettingsCollectionPath, "ProjectColor", this.ProjectColor.ToArgb()));
		}

		if (writableSettingsStore.PropertyExists(Constants.SettingsCollectionPath, "ViewFilename"))
		{
		  bool b = this.ViewFilename;
		  if (Boolean.TryParse(writableSettingsStore.GetString(Constants.SettingsCollectionPath, "ViewFilename"), out b))
			this.ViewFilename = b;
		}

		if (writableSettingsStore.PropertyExists(Constants.SettingsCollectionPath, "ViewFolders"))
		{
		  bool b = this.ViewFolders;
		  if (Boolean.TryParse(writableSettingsStore.GetString(Constants.SettingsCollectionPath, "ViewFolders"), out b))
			this.ViewFolders = b;
		}

		if (writableSettingsStore.PropertyExists(Constants.SettingsCollectionPath, "ViewProject"))
		{
		  bool b = this.ViewProject;
		  if (Boolean.TryParse(writableSettingsStore.GetString(Constants.SettingsCollectionPath, "ViewProject"), out b))
			this.ViewProject = b;
		}

		if (writableSettingsStore.PropertyExists(Constants.SettingsCollectionPath, "FilenameSize"))
		{
		  double d = this.FilenameSize;
		  if (Double.TryParse(writableSettingsStore.GetString(Constants.SettingsCollectionPath, "FilenameSize"), out d))
			this.FilenameSize = d;
		}

		if (writableSettingsStore.PropertyExists(Constants.SettingsCollectionPath, "FoldersSize"))
		{
		  double d = this.FoldersSize;
		  if (Double.TryParse(writableSettingsStore.GetString(Constants.SettingsCollectionPath, "FoldersSize"), out d))
			this.FoldersSize = d;
		}

		if (writableSettingsStore.PropertyExists(Constants.SettingsCollectionPath, "ProjectSize"))
		{
		  double d = this.ProjectSize;
		  if (Double.TryParse(writableSettingsStore.GetString(Constants.SettingsCollectionPath, "ProjectSize"), out d))
			this.ProjectSize = d;
		}

		if (writableSettingsStore.PropertyExists(Constants.SettingsCollectionPath, "Position"))
		{
		  AdornmentPositions p = this.Position;
		  if (Enum.TryParse<AdornmentPositions>(writableSettingsStore.GetString(Constants.SettingsCollectionPath, "Position"), out p))
			this.Position = p;
		}

		if (writableSettingsStore.PropertyExists(Constants.SettingsCollectionPath, "Opacity"))
		{
		  double d = this.Opacity;
		  if (Double.TryParse(writableSettingsStore.GetString(Constants.SettingsCollectionPath, "Opacity"), out d))
			this.Opacity = d;
		}
	  }
	  catch (Exception ex)
	  {
		Debug.Fail(ex.Message);
	  }
	}

	public override bool Equals(object obj)
	{
	  WhereAmISettings o = obj as WhereAmISettings;

	  return o.GetHashCode() == this.GetHashCode();
	}

	/// <summary>
	/// Overridden GetHasCode
	/// </summary>
	/// <returns></returns>
	public override int GetHashCode()
	{
	  int hash = 13;
	  hash = (hash * 7) + this.FilenameColor.GetHashCode();
	  hash = (hash * 7) + this.FilenameSize.GetHashCode();
	  hash = (hash * 7) + this.FoldersColor.GetHashCode();
	  hash = (hash * 7) + this.FoldersSize.GetHashCode();
	  hash = (hash * 7) + this.Opacity.GetHashCode();
	  hash = (hash * 7) + this.Position.GetHashCode();
	  hash = (hash * 7) + this.ProjectColor.GetHashCode();
	  hash = (hash * 7) + this.ProjectSize.GetHashCode();
	  hash = (hash * 7) + this.ViewFilename.GetHashCode();
	  hash = (hash * 7) + this.ViewFolders.GetHashCode();
	  hash = (hash * 7) + this.ViewProject.GetHashCode();

	  return hash;
	}

	public static IWhereAmISettings DarkThemeDefaults()
	{
	  WhereAmISettings settings = new WhereAmISettings();
	  settings.FilenameSize = 60;
	  settings.FoldersSize = settings.ProjectSize = 52;

	  settings.FilenameColor = Color.FromArgb(48, 48, 48);
	  settings.FoldersColor = settings.ProjectColor = Color.FromArgb(40, 40, 40);

	  settings.Position = AdornmentPositions.TopRight;

	  settings.Opacity = 1;
	  settings.ViewFilename = true;
	  settings.ViewFolders = true;
	  settings.ViewProject = true;

	  settings.Theme = Theme.Dark;

	  return settings;
	}

	public static IWhereAmISettings LightThemeDefaults()
	{
	  WhereAmISettings settings = new WhereAmISettings();
	  settings.FilenameSize = 60;
	  settings.FoldersSize = settings.ProjectSize = 52;

	  settings.FilenameColor = Color.FromArgb(234, 234, 234);
	  settings.FoldersColor = settings.ProjectColor = Color.FromArgb(243, 243, 243);

	  settings.Position = AdornmentPositions.TopRight;

	  settings.Opacity = 1;
	  settings.ViewFilename = true;
	  settings.ViewFolders = true;
	  settings.ViewProject = true;

	  settings.Theme = Theme.Light;

	  return settings;
	}
  }
}
