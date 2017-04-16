using System.ComponentModel;

namespace WhereAmI2017
{
    /// <summary>
    /// Available positions for the adornment layer
    /// </summary>
    public enum AdornmentPositions
    {
        [Description("Top-right corner")]
        /// <summary>
        /// Top-right corner of the view
        /// </summary>
        TopRight,

        [Description("Bottom-right corner")]
        /// <summary>
        /// Bottom-right corner of the view
        /// </summary>
        BottomRight
    }
}
