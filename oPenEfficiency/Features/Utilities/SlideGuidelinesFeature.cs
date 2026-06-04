using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Utilities
{
    /// <summary>
    /// Feature: Removes existing guidelines and creates new standard guidelines for slide layout.
    /// Uses line shapes styled as guidelines (native PowerPoint guidelines are not accessible via interop).
    ///
    /// Horizontal Guides (Top to Bottom) - manage vertical space:
    /// - 8.25 cm Top: Top Margin (~1.25 cm border at top)
    /// - 5.70 cm Top: Title Baseline (separates title from body, ~2.5 cm strip)
    /// - 0.00 cm: Vertical Center
    /// - 8.25 cm Bottom: Bottom Margin (for page numbers, footers)
    ///
    /// Vertical Guides (Left to Right) - create boundaries and column layouts
    /// Multiple layout options available: 3-Column, 4-Column, 5-Column, 1:2 Asymmetrical
    /// </summary>
    [FeatureMetadata(
        Id = "BtnSlideGuidelines",
        Name = "Slide Guidelines",
        Tooltip = "Slide guidelines",
        IconData = "M3,3H21V21H3V3M7,7V19H19V7H7M9,9H17V11H9V9M9,13H17V15H9V13M9,17H17V19H9V17Z",
        Color = "#F43F5E",
        Description = "Creates standard slide layout guidelines for consistent positioning.",
        DetailedHelpText = "### Slide Guidelines\nInjects precision architectural guidelines into the current slide for complex layout alignment.\n\n**Right-Click Options:**\n* **3/4/5 Column Layouts**: Balanced vertical gutters for content distribution.\n* **Rule of Thirds**: 3x3 composition grid for professional image and text placement.\n* **1:2 Ratio**: Asymmetrical grid for sidebar-style slides.",
        Keywords = "guidelines, layout, grids, align, margins, structure",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class SlideGuidelinesFeature
    {
        // Conversion constant: 1 cm = 28.3464567 points
        private const double PtsPerCm = 28.3464567;

        public enum GuidelineLayout
        {
            Default,        // 2-column layout
            ThreeColumn,    // 3-column layout
            FourColumn,     // 4-column layout
            FiveColumn,     // 5-column layout
            OneTwoColumn,   // 1:2 asymmetrical layout
            RuleOfThirds    // 3x3 grid for composition
        }

        /// <summary>
        /// Wrapper for auto-discovery - creates default 2-column layout guidelines.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, GuidelineLayout.Default);
        }

        public static bool Execute(PowerPointManager manager, GuidelineLayout layout = GuidelineLayout.Default)
        {
            if (manager == null) return false;

            try
            {
                System.Diagnostics.Debug.WriteLine($"SlideGuidelinesFeature.Execute: Starting with layout={layout}");

                var app = manager.GetApplication();
                if (app == null || app.ActiveWindow == null || app.ActiveWindow.View == null)
                {
                    return false;
                }

                var slide = (Slide)app.ActiveWindow.View.Slide;
                if (slide == null)
                {
                    return false;
                }

                System.Diagnostics.Debug.WriteLine($"SlideGuidelinesFeature: Processing slide");

                // Step 1: Remove existing guidelines (by name pattern)
                System.Diagnostics.Debug.WriteLine("SlideGuidelinesFeature: Step 1 - Removing existing guidelines");
                RemoveExistingGuidelines(slide);

                // Step 2: Create new guidelines based on layout
                System.Diagnostics.Debug.WriteLine("SlideGuidelinesFeature: Step 2 - Creating new guidelines");
                CreateGuidelines(slide, layout);

                System.Diagnostics.Debug.WriteLine("SlideGuidelinesFeature: Complete");
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SlideGuidelinesFeature.Execute");
                System.Windows.Forms.MessageBox.Show(
                    $"Error creating guidelines: {ex.Message}",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// Removes all existing guidelines by detecting name pattern (Guide_H_ or Guide_V_ prefix).
        /// Line shapes in PowerPoint don't support Tags, so we use name-based detection.
        /// </summary>
        private static void RemoveExistingGuidelines(Slide slide)
        {
            try
            {
                int removedCount = 0;
                int totalShapes = slide.Shapes.Count;

                System.Diagnostics.Debug.WriteLine($"SlideGuidelinesFeature: Checking {totalShapes} shapes for guidelines");

                // Iterate backwards to safely delete while iterating
                for (int i = slide.Shapes.Count; i >= 1; i--)
                {
                    try
                    {
                        var shape = slide.Shapes[i];
                        bool isGuideline = false;
                        string reason = "";

                        // Check by name pattern (Guide_H_ or Guide_V_ prefix)
                        if (shape.Name.StartsWith("Guide_H_") || shape.Name.StartsWith("Guide_V_"))
                        {
                            isGuideline = true;
                            reason = $"name pattern ({shape.Name})";
                        }

                        if (isGuideline)
                        {
                            System.Diagnostics.Debug.WriteLine($"  -> Deleting guideline (detected by {reason})");
                            shape.Delete();
                            removedCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"  -> Error processing shape: {ex.Message}");
                    }
                }

                System.Diagnostics.Debug.WriteLine($"SlideGuidelinesFeature: Removed {removedCount} existing guidelines");
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SlideGuidelinesFeature.RemoveExistingGuidelines");
            }
        }

        /// <summary>
        /// Creates new horizontal and vertical guidelines based on the selected layout.
        /// </summary>
        private static void CreateGuidelines(Slide slide, GuidelineLayout layout)
        {
            try
            {
                // Standard slide dimensions (16:9 aspect ratio)
                var presentation = slide.Parent as Presentation;
                float slideWidthCm = (float)(presentation.PageSetup.SlideWidth / PtsPerCm);
                float slideHeightCm = (float)(presentation.PageSetup.SlideHeight / PtsPerCm);

                // Center points
                float centerX = slideWidthCm / 2f;
                float centerY = slideHeightCm / 2f;

                // Rule of Thirds layout has its own guide set
                if (layout == GuidelineLayout.RuleOfThirds)
                {
                    CreateGuidelines_RuleOfThirds(slide, presentation, slideWidthCm, slideHeightCm);
                    return;
                }

                // Horizontal Guides (same for all other layouts)
                CreateHorizontalGuide(slide, presentation, 8.25f, "Top Margin");       // 8.25 cm from top
                CreateHorizontalGuide(slide, presentation, 5.70f, "Title Baseline");   // 5.70 cm from top
                CreateHorizontalGuide(slide, presentation, centerY, "Vertical Center"); // At center
                CreateHorizontalGuide(slide, presentation, slideHeightCm - 8.25f, "Bottom Margin"); // 8.25 cm from bottom

                // Vertical Guides (varies by layout)
                switch (layout)
                {
                    case GuidelineLayout.ThreeColumn:
                        CreateGuidelines_ThreeColumn(slide, presentation, slideWidthCm);
                        break;
                    case GuidelineLayout.FourColumn:
                        CreateGuidelines_FourColumn(slide, presentation, slideWidthCm);
                        break;
                    case GuidelineLayout.FiveColumn:
                        CreateGuidelines_FiveColumn(slide, presentation, slideWidthCm);
                        break;
                    case GuidelineLayout.OneTwoColumn:
                        CreateGuidelines_OneTwoColumn(slide, presentation, slideWidthCm);
                        break;
                    case GuidelineLayout.Default:
                    default:
                        CreateGuidelines_Default(slide, presentation, slideWidthCm);
                        break;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SlideGuidelinesFeature.CreateGuidelines");
                throw;
            }
        }

        /// <summary>
        /// Default 2-column layout vertical guides.
        /// </summary>
        private static void CreateGuidelines_Default(Slide slide, Presentation presentation, float slideWidthCm)
        {
            CreateVerticalGuide(slide, presentation, 15.70f, "Left Margin");
            CreateVerticalGuide(slide, presentation, 0.60f, "Left Gutter");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 2f, "Horizontal Center");
            CreateVerticalGuide(slide, presentation, slideWidthCm - 0.60f, "Right Gutter");
            CreateVerticalGuide(slide, presentation, slideWidthCm - 15.70f, "Right Margin");
        }

        /// <summary>
        /// 3-Column Layout vertical guides.
        /// </summary>
        private static void CreateGuidelines_ThreeColumn(Slide slide, Presentation presentation, float slideWidthCm)
        {
            CreateVerticalGuide(slide, presentation, 15.70f, "Left Margin");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 2f - 6.03f, "Left Gutter 1");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 2f - 4.83f, "Left Gutter 2");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 2f + 4.83f, "Right Gutter 1");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 2f + 6.03f, "Right Gutter 2");
            CreateVerticalGuide(slide, presentation, slideWidthCm - 15.70f, "Right Margin");
        }

        /// <summary>
        /// 4-Column Layout vertical guides.
        /// </summary>
        private static void CreateGuidelines_FourColumn(Slide slide, Presentation presentation, float slideWidthCm)
        {
            CreateVerticalGuide(slide, presentation, 15.70f, "Left Margin");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 2f - 8.75f, "Column 1 Left");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 2f - 7.55f, "Column 1 Right");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 2f - 0.60f, "Left Gutter");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 2f + 0.60f, "Right Gutter");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 2f + 7.55f, "Column 2 Right");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 2f + 8.75f, "Column 2 Left");
            CreateVerticalGuide(slide, presentation, slideWidthCm - 15.70f, "Right Margin");
        }

        /// <summary>
        /// 5-Column Layout vertical guides.
        /// </summary>
        private static void CreateGuidelines_FiveColumn(Slide slide, Presentation presentation, float slideWidthCm)
        {
            CreateVerticalGuide(slide, presentation, 15.70f, "Left Margin");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 2f - 10.38f, "Column 1 Left");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 2f - 9.18f, "Column 1 Right");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 2f - 3.86f, "Column 2 Left");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 2f - 2.66f, "Column 2 Right");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 2f + 2.66f, "Column 3 Right");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 2f + 3.86f, "Column 3 Left");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 2f + 9.18f, "Column 4 Right");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 2f + 10.38f, "Column 4 Left");
            CreateVerticalGuide(slide, presentation, slideWidthCm - 15.70f, "Right Margin");
        }

        /// <summary>
        /// 1:2 Asymmetrical Column Layout vertical guides.
        /// </summary>
        private static void CreateGuidelines_OneTwoColumn(Slide slide, Presentation presentation, float slideWidthCm)
        {
            CreateVerticalGuide(slide, presentation, 15.70f, "Left Margin");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 3f, "Column 1 Left Boundary");
            CreateVerticalGuide(slide, presentation, slideWidthCm / 3f + 1.2f, "Column 1 Right / Gutter");
            CreateVerticalGuide(slide, presentation, slideWidthCm - 15.70f, "Right Margin");
        }

        /// <summary>
        /// Rule of Thirds (3x3 grid) for professional composition.
        /// </summary>
        private static void CreateGuidelines_RuleOfThirds(Slide slide, Presentation presentation, float slideWidthCm, float slideHeightCm)
        {
            // Horizontal guides at 1/3 and 2/3 of slide height
            CreateHorizontalGuide(slide, presentation, slideHeightCm / 3f, "Third Top");
            CreateHorizontalGuide(slide, presentation, slideHeightCm * 2f / 3f, "Third Bottom");

            // Vertical guides at 1/3 and 2/3 of slide width
            CreateVerticalGuide(slide, presentation, slideWidthCm / 3f, "Third Left");
            CreateVerticalGuide(slide, presentation, slideWidthCm * 2f / 3f, "Third Right");
        }

        /// <summary>
        /// Creates a horizontal guideline at the specified position from top.
        /// </summary>
        private static void CreateHorizontalGuide(Slide slide, Presentation presentation, float positionFromTopCm, string name)
        {
            try
            {
                // Convert cm to points (PowerPoint uses points)
                float positionPts = (float)(positionFromTopCm * PtsPerCm);

                // Create the line
                var line = slide.Shapes.AddLine(0, positionPts, (float)presentation.PageSetup.SlideWidth, positionPts);

                // Style the guideline
                line.Name = $"Guide_H_{name}";
                line.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(100, System.Drawing.Color.Red));
                line.Line.Weight = 1f;
                line.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash;

                // Lock the shape to prevent accidental manipulation
                LockGuideline(line);

                // Send to back
                line.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack);

                System.Diagnostics.Debug.WriteLine($"  -> Added horizontal guide: {name} at {positionFromTopCm}cm ({positionPts}pts)");
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, $"SlideGuidelinesFeature.CreateHorizontalGuide.{name}");
            }
        }

        /// <summary>
        /// Creates a vertical guideline at the specified position from left.
        /// </summary>
        private static void CreateVerticalGuide(Slide slide, Presentation presentation, float positionFromLeftCm, string name)
        {
            try
            {
                // Convert cm to points (PowerPoint uses points)
                float positionPts = (float)(positionFromLeftCm * PtsPerCm);

                // Create the line
                var line = slide.Shapes.AddLine(positionPts, 0, positionPts, (float)presentation.PageSetup.SlideHeight);

                // Style the guideline
                line.Name = $"Guide_V_{name}";
                line.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(100, System.Drawing.Color.Blue));
                line.Line.Weight = 1f;
                line.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash;

                // Lock the shape to prevent accidental manipulation
                LockGuideline(line);

                // Send to back
                line.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack);

                System.Diagnostics.Debug.WriteLine($"  -> Added vertical guide: {name} at {positionFromLeftCm}cm ({positionPts}pts)");
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, $"SlideGuidelinesFeature.CreateVerticalGuide.{name}");
            }
        }

        /// <summary>
        /// Applies locking properties to a guideline shape to prevent accidental manipulation.
        /// </summary>
        /// <remarks>
        /// <para><strong>Tag Prefix Convention:</strong> All custom tags added by oPenEfficiency features
        /// use the <c>oPE_</c> prefix (e.g., <c>"oPE_Locked"</c>, <c>"oPE_Guideline"</c>, <c>"oPE_Widget"</c>).</para>
        /// <para><strong>Why oPE_ prefix?</strong> PowerPoint doesn't persist custom properties reliably
        /// through copy/paste operations. Using a unique prefix avoids conflicts with any user-defined
        /// tags and allows features to identify shapes created by oPenEfficiency.</para>
        /// <para><strong>Detection Strategy:</strong> Features detect oPenEfficiency shapes by checking
        /// <c>AlternativeText</c> (survives copy/paste) and <c>Tags</c> collection (when available).
        /// Line shapes may not support Tags in some PowerPoint versions, so AlternativeText is the
        /// primary detection mechanism.</para>
        /// </remarks>
        private static void LockGuideline(Shape guideline)
        {
            try
            {
                // Lock aspect ratio (prevents proportional resizing)
                guideline.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

                // Mark as guideline in AlternativeText for detection in selection handlers
                guideline.AlternativeText = "oPE_Guideline";

                // Add tag for detection (line shapes may not support Tags in all PowerPoint versions)
                try
                {
                    guideline.Tags.Add("oPE_Locked", "true");
                }
                catch
                {
                    // Line shapes may not support Tags in some versions - AlternativeText is still set
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SlideGuidelinesFeature.LockGuideline");
            }
        }
    }
}
