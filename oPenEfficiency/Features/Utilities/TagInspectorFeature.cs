using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnTagInspector",
        Name = "Tag Inspector",
        Tooltip = "Inspect and manage hidden metadata tags",
        IconData = "M21.41 11.58l-9-9C12.05 2.22 11.55 2 11 2H4c-1.1 0-2 .9-2 2v7c0 .55.22 1.05.59 1.42l9 9c.36.36.86.58 1.41.58.55 0 1.05-.22 1.41-.59l7-7c.37-.36.59-.86.59-1.41 0-.55-.23-1.06-.59-1.42M5.5 7C4.67 7 4 6.33 4 5.5S4.67 4 5.5 4 7 4.67 7 5.5 6.33 7 5.5 7z",
        Color = "#8B5CF6",
        Description = "Allows viewing, adding, and deleting hidden metadata tags on the current slide, the presentation, or selected shapes.",
        Keywords = "tags, metadata, hidden properties, inspect, debug, properties",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class TagInspectorFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                var window = new oPenEfficiency.UI.Dialogs.TagInspectorWindow(manager);
                
                try
                {
                    // Set owner to PowerPoint window to prevent it appearing behind or getting lost
                    var app = manager.GetApplication();
                    var hwnd = app.HWND;
                    var helper = new System.Windows.Interop.WindowInteropHelper(window);
                    helper.Owner = (IntPtr)hwnd;
                }
                catch { }

                window.Show();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TagInspectorFeature.Execute");
                return false;
            }
        }
    }
}
