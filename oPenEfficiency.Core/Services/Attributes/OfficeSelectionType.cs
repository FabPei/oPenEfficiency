namespace oPenEfficiency.Services.Attributes
{
    /// <summary>
    /// Host-neutral selection type used by <see cref="FeatureMetadataAttribute"/>.
    /// Each Office host maps this to its own native selection enum at the boundary
    /// (e.g. PowerPoint's PpSelectionType, Excel's ranges/shapes).
    /// This keeps feature metadata sharable across PowerPoint, Excel, Word and OneNote.
    /// </summary>
    public enum OfficeSelectionType
    {
        /// <summary>No selection required.</summary>
        None = 0,

        /// <summary>One or more shapes/objects are selected.</summary>
        Shapes = 1,

        /// <summary>Text (inside a shape/cell) is selected / the cursor is in text.</summary>
        Text = 2,

        /// <summary>Slides are selected (PowerPoint) / sheets (Excel).</summary>
        Slides = 3,

        /// <summary>A cell range is selected (Excel).</summary>
        Cells = 4
    }
}
