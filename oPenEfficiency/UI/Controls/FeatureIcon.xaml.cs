using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using oPenEfficiency.Services;

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

        public FeatureIcon()
        {
            InitializeComponent();
        }

        private static void OnFeatureIdChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is FeatureIcon control)
            {
                control.UpdateIcon();
            }
        }

        private static void OnIconColorChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is FeatureIcon control)
            {
                control.UpdateIcon();
            }
        }

        private void UpdateIcon()
        {
            if (string.IsNullOrEmpty(FeatureId))
            {
                IconPath.Data = null;
                return;
            }

            try
            {
                bool isFilled;
                var geometry = IconService.GetIconGeometry(FeatureId, out isFilled);

                if (geometry != null)
                {
                    IconPath.Data = geometry;
                    if (isFilled)
                    {
                        IconPath.Fill = IconColor;
                        IconPath.Stroke = null;
                    }
                    else
                    {
                        IconPath.Fill = Brushes.Transparent;
                        IconPath.Stroke = IconColor;
                        IconPath.StrokeThickness = 2;
                        IconPath.StrokeLineJoin = PenLineJoin.Round;
                        IconPath.StrokeStartLineCap = PenLineCap.Round;
                        IconPath.StrokeEndLineCap = PenLineCap.Round;
                    }
                }
            }
            catch
            {
                // Silently fail or log as needed
            }
        }
    }
}
