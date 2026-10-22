using System;
using System.Collections.Generic;
using System.Linq;

namespace oPenEfficiency.Services
{
    public sealed class ShapeGeometry
    {
        public int Id { get; set; }
        public float Left { get; set; }
        public float Top { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }

        public ShapeGeometry Clone()
        {
            return (ShapeGeometry)MemberwiseClone();
        }
    }

    public static class ShapeGeometryPlanner
    {
        public static IReadOnlyList<ShapeGeometry> AdjustSpacing(
            IReadOnlyList<ShapeGeometry> source,
            float increment,
            bool horizontal,
            bool preserveCenter)
        {
            var result = CloneAndValidate(source);
            var ordered = result
                .OrderBy(shape => horizontal ? shape.Left : shape.Top)
                .ThenBy(shape => shape.Id)
                .ToList();
            float compensation = preserveCenter ? ((ordered.Count - 1) * increment) / 2f : 0f;

            for (int index = 0; index < ordered.Count; index++)
            {
                float offset = index * increment - compensation;
                if (horizontal) ordered[index].Left += offset;
                else ordered[index].Top += offset;
            }

            return result;
        }

        public static IReadOnlyList<ShapeGeometry> RemoveGaps(
            IReadOnlyList<ShapeGeometry> source,
            bool horizontal,
            bool packToCenter,
            int? anchorId)
        {
            var result = CloneAndValidate(source);
            var ordered = result
                .OrderBy(shape => horizontal ? shape.Left : shape.Top)
                .ThenBy(shape => shape.Id)
                .ToList();

            if (packToCenter)
            {
                float minimum = ordered.Min(shape => horizontal ? shape.Left : shape.Top);
                float maximum = ordered.Max(shape => horizontal
                    ? shape.Left + shape.Width
                    : shape.Top + shape.Height);
                float total = ordered.Sum(shape => horizontal ? shape.Width : shape.Height);
                float cursor = (minimum + maximum - total) / 2f;
                foreach (var shape in ordered)
                {
                    if (horizontal)
                    {
                        shape.Left = cursor;
                        cursor += shape.Width;
                    }
                    else
                    {
                        shape.Top = cursor;
                        cursor += shape.Height;
                    }
                }
                return result;
            }

            int anchorIndex = anchorId.HasValue
                ? ordered.FindIndex(shape => shape.Id == anchorId.Value)
                : -1;
            if (anchorIndex < 0)
                throw new ArgumentException("The anchor must be present in the source geometry.", nameof(anchorId));

            for (int index = anchorIndex - 1; index >= 0; index--)
            {
                if (horizontal)
                    ordered[index].Left = ordered[index + 1].Left - ordered[index].Width;
                else
                    ordered[index].Top = ordered[index + 1].Top - ordered[index].Height;
            }
            for (int index = anchorIndex + 1; index < ordered.Count; index++)
            {
                if (horizontal)
                    ordered[index].Left = ordered[index - 1].Left + ordered[index - 1].Width;
                else
                    ordered[index].Top = ordered[index - 1].Top + ordered[index - 1].Height;
            }

            return result;
        }

        public static IReadOnlyList<ShapeGeometry> ArrangeGrid(
            IReadOnlyList<ShapeGeometry> source,
            int rows,
            int columns,
            float spacing)
        {
            var result = CloneAndValidate(source);
            if (rows < 1) throw new ArgumentOutOfRangeException(nameof(rows));
            if (columns < 1) throw new ArgumentOutOfRangeException(nameof(columns));
            if (spacing < 0) throw new ArgumentOutOfRangeException(nameof(spacing));
            if ((long)rows * columns < result.Count)
                throw new ArgumentException("The grid does not contain enough cells.");

            float minimumLeft = result.Min(shape => shape.Left);
            float minimumTop = result.Min(shape => shape.Top);
            var columnWidths = new float[columns];
            var rowHeights = new float[rows];
            for (int index = 0; index < result.Count; index++)
            {
                int row = index / columns;
                int column = index % columns;
                columnWidths[column] = Math.Max(columnWidths[column], result[index].Width);
                rowHeights[row] = Math.Max(rowHeights[row], result[index].Height);
            }

            var columnOffsets = PrefixOffsets(columnWidths, spacing);
            var rowOffsets = PrefixOffsets(rowHeights, spacing);
            for (int index = 0; index < result.Count; index++)
            {
                result[index].Left = minimumLeft + columnOffsets[index % columns];
                result[index].Top = minimumTop + rowOffsets[index / columns];
            }
            return result;
        }

        public static IReadOnlyList<ShapeGeometry> SplitGrid(
            ShapeGeometry bounds,
            int rows,
            int columns,
            float spacing)
        {
            ValidateShape(bounds);
            if (rows < 1) throw new ArgumentOutOfRangeException(nameof(rows));
            if (columns < 1) throw new ArgumentOutOfRangeException(nameof(columns));
            if (spacing < 0) throw new ArgumentOutOfRangeException(nameof(spacing));

            float width = (bounds.Width - spacing * (columns - 1)) / columns;
            float height = (bounds.Height - spacing * (rows - 1)) / rows;
            if (width <= 0 || height <= 0)
                throw new ArgumentException("Spacing leaves no positive cell area.", nameof(spacing));

            var result = new List<ShapeGeometry>(rows * columns);
            for (int row = 0; row < rows; row++)
            {
                for (int column = 0; column < columns; column++)
                {
                    result.Add(new ShapeGeometry
                    {
                        Id = result.Count,
                        Left = bounds.Left + column * (width + spacing),
                        Top = bounds.Top + row * (height + spacing),
                        Width = width,
                        Height = height
                    });
                }
            }
            return result;
        }

        private static List<ShapeGeometry> CloneAndValidate(IReadOnlyList<ShapeGeometry> source)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (source.Count == 0) throw new ArgumentException("At least one shape is required.", nameof(source));
            var result = new List<ShapeGeometry>(source.Count);
            foreach (var shape in source)
            {
                ValidateShape(shape);
                result.Add(shape.Clone());
            }
            if (result.Select(shape => shape.Id).Distinct().Count() != result.Count)
                throw new ArgumentException("Shape identifiers must be unique.", nameof(source));
            return result;
        }

        private static void ValidateShape(ShapeGeometry shape)
        {
            if (shape == null) throw new ArgumentException("Shape geometry cannot be null.");
            if (shape.Width <= 0 || shape.Height <= 0)
                throw new ArgumentException("Shape dimensions must be positive.");
        }

        private static float[] PrefixOffsets(float[] sizes, float spacing)
        {
            var offsets = new float[sizes.Length];
            for (int index = 1; index < offsets.Length; index++)
                offsets[index] = offsets[index - 1] + sizes[index - 1] + spacing;
            return offsets;
        }
    }
}
