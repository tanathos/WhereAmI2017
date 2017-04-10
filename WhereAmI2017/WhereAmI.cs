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
        /// <summary>
        /// The width of the square box.
        /// </summary>
        private const double AdornmentWidth = 30;

        /// <summary>
        /// The height of the square box.
        /// </summary>
        private const double AdornmentHeight = 30;

        /// <summary>
        /// Distance from the viewport top to the top of the square box.
        /// </summary>
        private const double TopMargin = 30;

        /// <summary>
        /// Distance from the viewport right to the right end of the square box.
        /// </summary>
        private const double RightMargin = 30;

        /// <summary>
        /// Text view to add the adornment on.
        /// </summary>
        private readonly IWpfTextView view;

        /// <summary>
        /// Adornment image
        /// </summary>
        private readonly Image image;

        /// <summary>
        /// The layer for the adornment.
        /// </summary>
        private readonly IAdornmentLayer adornmentLayer;

        /// <summary>
        /// Initializes a new instance of the <see cref="WhereAmI"/> class.
        /// Creates a square image and attaches an event handler to the layout changed event that
        /// adds the the square in the upper right-hand corner of the TextView via the adornment layer
        /// </summary>
        /// <param name="view">The <see cref="IWpfTextView"/> upon which the adornment will be drawn</param>
        public WhereAmI(IWpfTextView view)
        {
            if (view == null)
            {
                throw new ArgumentNullException("view");
            }

            this.view = view;

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

                Project proj = GetContainingProject(fileName);
                if (proj != null)
                {
                    string projectName = proj.Name;

                    //if (_settings.ViewFilename)
                    //{
                    //    _fileName.Text = fileName;

                    //    Brush fileNameBrush = new SolidColorBrush(Color.FromArgb(_settings.FilenameColor.A, _settings.FilenameColor.R, _settings.FilenameColor.G, _settings.FilenameColor.B));
                    //    _fileName.FontFamily = new FontFamily("Consolas");
                    //    _fileName.FontSize = _settings.FilenameSize;
                    //    _fileName.HorizontalAlignment = System.Windows.HorizontalAlignment.Right;
                    //    _fileName.TextAlignment = System.Windows.TextAlignment.Right;
                    //    _fileName.Foreground = fileNameBrush;
                    //    _fileName.Opacity = _settings.Opacity;
                    //}

                    //if (_settings.ViewFolders)
                    //{
                    //    _folderStructure.Text = GetFolderDiffs(textDoc.FilePath, proj.FullName);

                    //    Brush foldersBrush = new SolidColorBrush(Color.FromArgb(_settings.FoldersColor.A, _settings.FoldersColor.R, _settings.FoldersColor.G, _settings.FoldersColor.B));
                    //    _folderStructure.FontFamily = new FontFamily("Consolas");
                    //    _folderStructure.FontSize = _settings.FoldersSize;
                    //    _folderStructure.HorizontalAlignment = System.Windows.HorizontalAlignment.Right;
                    //    _folderStructure.TextAlignment = System.Windows.TextAlignment.Right;
                    //    _folderStructure.Foreground = foldersBrush;
                    //    _folderStructure.Opacity = _settings.Opacity;
                    //}

                    //if (_settings.ViewProject)
                    //{
                    //    _projectName.Text = projectName;

                    //    Brush projectNameBrush = new SolidColorBrush(Color.FromArgb(_settings.ProjectColor.A, _settings.ProjectColor.R, _settings.ProjectColor.G, _settings.ProjectColor.B));
                    //    _projectName.FontFamily = new FontFamily("Consolas");
                    //    _projectName.FontSize = _settings.ProjectSize;
                    //    _projectName.HorizontalAlignment = System.Windows.HorizontalAlignment.Right;
                    //    _projectName.TextAlignment = System.Windows.TextAlignment.Right;
                    //    _projectName.Foreground = projectNameBrush;
                    //    _projectName.Opacity = _settings.Opacity;
                    //}
                }
            }

                //var brush = new SolidColorBrush(Colors.BlueViolet);
                //brush.Freeze();
                //var penBrush = new SolidColorBrush(Colors.Red);
                //penBrush.Freeze();
                //var pen = new Pen(penBrush, 0.5);
                //pen.Freeze();

                //// Draw a square with the created brush and pen
                //System.Windows.Rect r = new System.Windows.Rect(0, 0, AdornmentWidth, AdornmentHeight);
                //var geometry = new RectangleGeometry(r);

                //var drawing = new GeometryDrawing(brush, pen, geometry);
                //drawing.Freeze();

                //var drawingImage = new DrawingImage(drawing);
                //drawingImage.Freeze();

                //this.image = new Image
                //{
                //    Source = drawingImage,
                //};

                //this.adornmentLayer = view.GetAdornmentLayer("WhereAmI");

                //this.view.ViewportHeightChanged += this.OnSizeChanged;
                //this.view.ViewportWidthChanged += this.OnSizeChanged;
            }

        /// <summary>
        /// Event handler for viewport height or width changed events. Adds adornment at the top right corner of the viewport.
        /// </summary>
        /// <param name="sender">Event sender</param>
        /// <param name="e">Event arguments</param>
        private void OnSizeChanged(object sender, EventArgs e)
        {
            // Clear the adornment layer of previous adornments
            this.adornmentLayer.RemoveAllAdornments();

            // Place the image in the top right hand corner of the Viewport
            Canvas.SetLeft(this.image, this.view.ViewportRight - RightMargin - AdornmentWidth);
            Canvas.SetTop(this.image, this.view.ViewportTop + TopMargin);

            // Add the image to the adornment layer and make it relative to the viewport
            this.adornmentLayer.AddAdornment(AdornmentPositioningBehavior.ViewportRelative, null, null, this.image, null);
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
    }
}
