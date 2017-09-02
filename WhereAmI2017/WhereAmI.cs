//------------------------------------------------------------------------------
// <copyright file="WhereAmI.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using EnvDTE80;
using EnvDTE;

namespace WhereAmI2017
{
    /// <summary>
    /// Adornment class that draws a square box in the top right hand corner of the viewport
    /// </summary>
    internal sealed class WhereAmI
    {
        private TextBlock _fileName;
        private TextBlock _folderStructure;
        private TextBlock _projectName;

        private IWpfTextView _view;
        private IAdornmentLayer _adornmentLayer;

        /// <summary>
        /// Settings of the extension, injected by the provider
        /// </summary>
        readonly IWhereAmISettings _settings;

        /// <summary>
        /// Initializes a new instance of the <see cref="WhereAmI"/> class.
        /// Creates a square image and attaches an event handler to the layout changed event that
        /// adds the the square in the upper right-hand corner of the TextView via the adornment layer
        /// </summary>
        /// <param name="view">The <see cref="IWpfTextView"/> upon which the adornment will be drawn</param>
        public WhereAmI(IWpfTextView view, IWhereAmISettings settings)
        {
            _view = view;
            _settings = settings;

            _fileName = new TextBlock();
            _folderStructure = new TextBlock();
            _projectName = new TextBlock();

            ITextDocument textDoc;
            object obj;
            if (view.TextBuffer.Properties.TryGetProperty<ITextDocument>(typeof(ITextDocument), out textDoc))
            {
                // Retrieved the ITextDocument from the first level
            }
            else if (view.TextBuffer.Properties.TryGetProperty<object>("IdentityMapping", out obj))
            {
                // Try to get the ITextDocument from the second level (e.g. Razor files)
                if ((obj as ITextBuffer) != null)
                {
                    (obj as ITextBuffer).Properties.TryGetProperty<ITextDocument>(typeof(ITextDocument), out textDoc);
                }
            }

            // If I found an ITextDocument, access to its FilePath prop to retrieve informations about Proj
            if (textDoc != null)
            {
                string fileName = System.IO.Path.GetFileName(textDoc.FilePath);

                Project proj = GetContainingProject(textDoc.FilePath);
                if (proj != null)
                {
                    string projectName = proj.Name;

                    if (_settings.ViewFilename)
                    {
                        _fileName.Text = fileName;

                        Brush fileNameBrush = new SolidColorBrush(Color.FromArgb(_settings.FilenameColor.A, _settings.FilenameColor.R, _settings.FilenameColor.G, _settings.FilenameColor.B));
                        _fileName.FontFamily = new FontFamily("Consolas");
                        _fileName.FontSize = _settings.FilenameSize;
                        _fileName.HorizontalAlignment = System.Windows.HorizontalAlignment.Right;
                        _fileName.TextAlignment = System.Windows.TextAlignment.Right;
                        _fileName.Foreground = fileNameBrush;
                        _fileName.Opacity = _settings.Opacity;
                    }

                    if (_settings.ViewFolders)
                    {
                        _folderStructure.Text = GetFolderDiffs(textDoc.FilePath, proj.FullName);

                        Brush foldersBrush = new SolidColorBrush(Color.FromArgb(_settings.FoldersColor.A, _settings.FoldersColor.R, _settings.FoldersColor.G, _settings.FoldersColor.B));
                        _folderStructure.FontFamily = new FontFamily("Consolas");
                        _folderStructure.FontSize = _settings.FoldersSize;
                        _folderStructure.HorizontalAlignment = System.Windows.HorizontalAlignment.Right;
                        _folderStructure.TextAlignment = System.Windows.TextAlignment.Right;
                        _folderStructure.Foreground = foldersBrush;
                        _folderStructure.Opacity = _settings.Opacity;
                    }

                    if (_settings.ViewProject)
                    {
                        _projectName.Text = projectName;

                        Brush projectNameBrush = new SolidColorBrush(Color.FromArgb(_settings.ProjectColor.A, _settings.ProjectColor.R, _settings.ProjectColor.G, _settings.ProjectColor.B));
                        _projectName.FontFamily = new FontFamily("Consolas");
                        _projectName.FontSize = _settings.ProjectSize;
                        _projectName.HorizontalAlignment = System.Windows.HorizontalAlignment.Right;
                        _projectName.TextAlignment = System.Windows.TextAlignment.Right;
                        _projectName.Foreground = projectNameBrush;
                        _projectName.Opacity = _settings.Opacity;
                    }
                }
            }

            // Force to have an ActualWidth
            System.Windows.Rect finalRect = new System.Windows.Rect();
            _fileName.Arrange(finalRect);
            _folderStructure.Arrange(finalRect);
            _projectName.Arrange(finalRect);

            //Grab a reference to the adornment layer that this adornment should be added to
            _adornmentLayer = view.GetAdornmentLayer(Constants.AdornmentLayerName);

            _view.ViewportHeightChanged += delegate { this.onSizeChange(); };
            _view.ViewportWidthChanged += delegate { this.onSizeChange(); };
        }

        public void onSizeChange()
        {
            _adornmentLayer.RemoveAllAdornments();
            _adornmentLayer.Opacity = 1;

            double lineTopPosition = 0;
            double fromTop = 1;

            switch (_settings.Position)
            {
                case AdornmentPositions.TopRight:
                default:
                    lineTopPosition = _view.ViewportTop + 5;
                    break;

                case AdornmentPositions.BottomRight:
                    lineTopPosition = _view.ViewportBottom - 5;
                    fromTop = -1;
                    break;
            }

            // Place the textes in the layer
            if (_settings.ViewFilename)
            {
                if (fromTop == -1 && lineTopPosition == (_view.ViewportBottom - 5))
                    lineTopPosition += _fileName.ActualHeight * fromTop;

                Canvas.SetLeft(_fileName, _view.ViewportRight - (_fileName.ActualWidth + 15));
                Canvas.SetTop(_fileName, lineTopPosition);

                _adornmentLayer.AddAdornment(AdornmentPositioningBehavior.ViewportRelative, null, null, _fileName, null);

                lineTopPosition += _fileName.ActualHeight * fromTop;
            }

            if (_settings.ViewFolders && !String.IsNullOrEmpty(_folderStructure.Text))
            {
                if (fromTop == -1 && lineTopPosition == (_view.ViewportBottom - 5))
                    lineTopPosition += _folderStructure.ActualHeight * fromTop;

                Canvas.SetLeft(_folderStructure, _view.ViewportRight - (_folderStructure.ActualWidth + 15));
                Canvas.SetTop(_folderStructure, lineTopPosition);

                _adornmentLayer.AddAdornment(AdornmentPositioningBehavior.ViewportRelative, null, null, _folderStructure, null);

                lineTopPosition += _folderStructure.ActualHeight * fromTop;
            }

            if (_settings.ViewProject)
            {
                if (fromTop == -1 && lineTopPosition == (_view.ViewportBottom - 5))
                    lineTopPosition += _projectName.ActualHeight * fromTop;

                Canvas.SetLeft(_projectName, _view.ViewportRight - (_projectName.ActualWidth + 15));
                Canvas.SetTop(_projectName, lineTopPosition);

                _adornmentLayer.AddAdornment(AdornmentPositioningBehavior.ViewportRelative, null, null, _projectName, null);
            }
        }

        /// <summary>
        /// Given a filename, retrieve the Project container
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static Project GetContainingProject(string fileName)
        {
            if (!String.IsNullOrEmpty(fileName))
            {
                var dte2 = (EnvDTE80.DTE2)Package.GetGlobalService(typeof(Microsoft.VisualStudio.Shell.Interop.SDTE));
                if (dte2 != null)
                {
                    var prjItem = dte2.Solution.FindProjectItem(fileName);
                    if (prjItem != null)
                        return prjItem.ContainingProject;
                }
            }
            return null;
        }

        /// <summary>
        /// Given 2 absolute paths, returns the difference in folder structure.
        /// (The first should be nested in the second)
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="folderPath"></param>
        /// <returns></returns>
        private static string GetFolderDiffs(string filePath, string folderPath)
        {
            if (!String.IsNullOrEmpty(folderPath))
                return System.IO.Path.GetDirectoryName(filePath).Replace(System.IO.Path.GetDirectoryName(folderPath), "").Replace("\\", "/").ToLower();

            return "";
        }
    }
}
