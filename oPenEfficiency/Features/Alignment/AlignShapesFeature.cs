namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Align shapes in a given direction (Left, Center, Right).
    /// </summary>
    public static class AlignShapesFeature
    {
        public static bool Execute(PowerPointManager manager, string direction)
        {
            return manager != null && !string.IsNullOrWhiteSpace(direction) && AlignmentHelpers.AlignShapes(manager, direction);
        }
    }
}
