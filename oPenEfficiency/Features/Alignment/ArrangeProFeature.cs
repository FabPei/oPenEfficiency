using System;
using System.Windows;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Arrange Pro - advanced shape arrangement dialog.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnArrangePro",
        Name = "Arrange Pro",
        Tooltip = "Arrange pro",
        IconData = "M12,2A10,10 0 0,0 2,12A10,10 0 0,0 12,22A10,10 0 0,0 22,12A10,10 0 0,0 12,2M12,4C13.5,4 14.9,4.4 16.1,5.1L14.9,6.3C14.1,5.8 13.1,5.5 12,5.5C9.2,5.5 7,7.7 7,10.5C7,11.3 7.2,12.1 7.5,12.8L6.1,14C5.4,12.9 5,11.7 5,10.5C5,6.9 8.2,4 12,4M12,20C10.5,20 9.1,19.6 7.9,18.9L9.1,17.7C9.9,18.2 10.9,18.5 12,18.5C14.8,18.5 17,16.3 17,13.5C17,12.7 16.8,11.9 16.5,11.2L17.9,10C18.6,11.1 19,12.3 19,13.5C19,17.1 15.8,20 12,20M12,8C13.9,8 15.5,9.6 15.5,11.5C15.5,13.4 13.9,15 12,15C10.1,15 8.5,13.4 8.5,11.5C8.5,9.6 10.1,8 12,8M12,10C11.4,10 11,10.4 11,11C11,11.6 11.4,12 12,12C12.6,12 13,11.6 13,11C13,10.4 12.6,10 12,10Z",
        Color = "#F43F5E",
        Description = "Opens the Arrange Pro dialog for advanced shape arrangement options.",
        DetailedHelpText = "### Arrange Pro\nAdvanced multi-axis arrangement tool combining alignment, distribution, and gap control in a single configurable operation.",
        Keywords = "advanced arrangement, multi-axis distribute, fine-tune alignment",
        MinSelection = 1)]
    public static class ArrangeProFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - returns false as Arrange Pro requires dialog interaction.
        /// Manual switch in MainSidebar.xaml.cs handles opening the ArrangeProWindow dialog.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            // Requires dialog for arrangement configuration
            // Handled by manual switch in MainSidebar.xaml.cs
            return false;
        }
    }
}
