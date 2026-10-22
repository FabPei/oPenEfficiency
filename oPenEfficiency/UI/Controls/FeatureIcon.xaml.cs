using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI.Controls
{
    public partial class FeatureIcon : UserControl
    {
        public static readonly DependencyProperty FeatureIdProperty =
            DependencyProperty.Register("FeatureId", typeof(string), typeof(FeatureIcon), 
                new PropertyMetadata(string.Empty, OnFeatureIdChanged));

        public string FeatureId
        {
            get => (string)GetValue(FeatureIdProperty);
            set => SetValue(FeatureIdProperty, value);
        }

        public static readonly DependencyProperty IconColorProperty =
            DependencyProperty.Register("IconColor", typeof(Brush), typeof(FeatureIcon), 
                new PropertyMetadata(Brushes.Gray, OnIconColorChanged));

        public Brush IconColor
        {
            get => (Brush)GetValue(IconColorProperty);
            set => SetValue(IconColorProperty, value);
        }

        public static readonly DependencyProperty ColorModeProperty =
            DependencyProperty.Register(
                nameof(ColorMode),
                typeof(string),
                typeof(FeatureIcon),
                new PropertyMetadata(IconColorModes.PowerPoint, OnRenderingPropertyChanged));

        public string ColorMode
        {
            get => (string)GetValue(ColorModeProperty);
            set => SetValue(ColorModeProperty, value);
        }

        public static readonly DependencyProperty VisualStateProperty =
            DependencyProperty.Register(
                nameof(VisualState),
                typeof(FeatureIconVisualState),
                typeof(FeatureIcon),
                new PropertyMetadata(FeatureIconVisualState.Normal, OnRenderingPropertyChanged));

        public FeatureIconVisualState VisualState
        {
            get => (FeatureIconVisualState)GetValue(VisualStateProperty);
            set => SetValue(VisualStateProperty, value);
        }

        private readonly System.Collections.Generic.List<System.Windows.Shapes.Path> _iconPaths =
            new System.Collections.Generic.List<System.Windows.Shapes.Path>();
        private ParsedIconGeometry _parsedIcon;

        public FeatureIcon()
        {
            InitializeComponent();
            IsEnabledChanged += (sender, args) => ApplyIconState();
        }

        private static void OnFeatureIdChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is FeatureIcon control)
            {
                control.RebuildIcon();
            }
        }

        private static void OnIconColorChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is FeatureIcon control)
            {
                control.ApplyIconState();
            }
        }

        private static void OnRenderingPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is FeatureIcon control)
            {
                control.ApplyIconState();
            }
        }

        /// <summary>
        /// Rebuilds the glyph geometry. Only a FeatureId change may reach this path: rebuilding
        /// replaces the Path children, which must never happen while a mouse event is in dispatch
        /// (PQ-015/PQ-016). All other changes go through <see cref="ApplyIconState"/>.
        /// </summary>
        private void RebuildIcon()
        {
            _iconPaths.Clear();
            IconLayerHost.Children.Clear();
            _parsedIcon = null;
            if (string.IsNullOrEmpty(FeatureId))
            {
                return;
            }

            try
            {
                var parsed = IconService.GetIconGeometryWithRoles(FeatureId);

                if (parsed == null || parsed.Elements.Count == 0)
                {
                    bool isFilled;
                    var geometry = IconService.GetIconGeometryLegacy(FeatureId, out isFilled);
                    if (geometry == null) return;

                    parsed = new ParsedIconGeometry { BaseGeometry = geometry as GeometryGroup };
                    parsed.Elements.Add(new ParsedIconElement
                    {
                        Geometry = geometry,
                        IsAccent = false,
                        AccentRole = FeatureIconAccentRole.None,
                        HasFill = isFilled,
                        HasStroke = !isFilled,
                        StrokeThickness = isFilled ? 0d : 2d,
                        StrokeLineCap = PenLineCap.Round,
                        StrokeLineJoin = PenLineJoin.Round
                    });
                }

                foreach (ParsedIconElement element in parsed.Elements)
                {
                    var path = FeatureIconPathFactory.CreatePath(element, parsed.ViewBox);
                    _iconPaths.Add(path);
                    IconLayerHost.Children.Add(path);
                }

                _parsedIcon = parsed;
                ApplyIconState();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FeatureIcon.RebuildIcon");
            }
        }

        /// <summary>
        /// Re-skins the existing paths for the current state. Never creates or removes visuals,
        /// so it is safe to call from mouse-driven state changes (hover/pressed) mid dispatch.
        /// </summary>
        private void ApplyIconState()
        {
            if (_iconPaths.Count == 0)
            {
                // Nothing built yet (for example a property set before the icon id); build now.
                RebuildIcon();
                return;
            }

            try
            {
                FeatureIconVisualState effectiveState = IsEnabled
                    ? VisualState
                    : FeatureIconVisualState.Disabled;
                FeatureIconPathFactory.ApplyState(
                    _iconPaths,
                    _parsedIcon,
                    ColorMode,
                    effectiveState,
                    IconColor);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FeatureIcon.ApplyIconState");
            }
        }
    }
}
