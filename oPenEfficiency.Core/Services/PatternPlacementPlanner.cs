using System;
using System.Collections.Generic;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Host-neutral target-position plans for the geometric arrangement commands (BtnArrangeInShape /
    /// BtnArrangeCircle and the ArrangePro presets). All targets are shape-center points; the COM layer
    /// places shapes on them. The math is deterministic and unit-tested; the office-automated leg of the
    /// same features verifies it on a real host.
    /// </summary>
    public static class PatternPlacementPlanner
    {
        public static List<Tuple<float, float>> BuildTargets(
            string shapeType,
            int count,
            float radius,
            float centerX,
            float centerY,
            float startAngleDegree,
            bool fillShape,
            bool clockwise,
            float arcSpanDegrees,
            bool reserveCenter)
        {
            var targets = new List<Tuple<float, float>>(count);

            if (fillShape && shapeType == "square")
            {
                int gridSize = (int)Math.Ceiling(Math.Sqrt(count));
                float spacing = gridSize > 1 ? 2 * radius / (gridSize - 1) : 0;
                int rowCount = (int)Math.Ceiling((double)count / gridSize);
                float startY = centerY - ((rowCount - 1) * spacing) / 2f;
                int placed = 0;
                for (int row = 0; row < rowCount; row++)
                {
                    int itemsInRow = Math.Min(gridSize, count - placed);
                    float startX = centerX - ((itemsInRow - 1) * spacing) / 2f;
                    for (int column = 0; column < itemsInRow; column++)
                    {
                        targets.Add(Tuple.Create(startX + column * spacing, startY + row * spacing));
                        placed++;
                    }
                }
            }
            else if (fillShape && shapeType == "triangle")
            {
                int rows = Math.Max(1, (int)Math.Ceiling((-1 + Math.Sqrt(1 + 8 * count)) / 2));
                float diameter = radius * 2f;
                float spacingY = rows > 1 ? 1.5f * diameter / (rows - 1) : 0;
                float spacingX = rows > 1 ? (float)(diameter * Math.Sqrt(3)) / (rows - 1) : 0;
                float startY = centerY - diameter;
                int[] rowCounts = new int[rows];
                int remaining = count;
                for (int row = rows - 1; row >= 0 && remaining > 0; row--)
                {
                    rowCounts[row] = Math.Min(row + 1, remaining);
                    remaining -= rowCounts[row];
                }

                for (int row = 0; row < rows; row++)
                {
                    int itemsInRow = rowCounts[row];
                    float startX = centerX - ((itemsInRow - 1) * spacingX) / 2f;
                    for (int column = 0; column < itemsInRow; column++)
                        targets.Add(Tuple.Create(startX + column * spacingX, startY + row * spacingY));
                }
            }
            else if (fillShape)
            {
                float goldenAngle = (float)(Math.PI * (3.0 - Math.Sqrt(5.0)));
                for (int i = 0; i < count; i++)
                {
                    double fraction = reserveCenter
                        ? (double)(i + 1) / count
                        : (count > 1 ? (double)i / (count - 1) : 0);
                    float itemRadius = radius * (float)Math.Sqrt(fraction);
                    float theta = i * goldenAngle - (float)Math.PI / 2;
                    targets.Add(Tuple.Create(
                        centerX + itemRadius * (float)Math.Cos(theta),
                        centerY + itemRadius * (float)Math.Sin(theta)));
                }
            }
            else
            {
                for (int i = 0; i < count; i++)
                {
                    int pathIndex = clockwise ? i : (count - i) % count;
                    double angle;
                    float targetX;
                    float targetY;

                    if (shapeType == "arc")
                    {
                        double fraction = count > 1
                            ? (arcSpanDegrees >= 359.999f ? (double)i / count : (double)i / (count - 1))
                            : 0;
                        double direction = clockwise ? 1.0 : -1.0;
                        angle = -Math.PI / 2 + direction * arcSpanDegrees * Math.PI / 180.0 * fraction;
                        targetX = centerX + radius * (float)Math.Cos(angle);
                        targetY = centerY + radius * (float)Math.Sin(angle);
                    }
                    else
                    {
                        angle = 2.0 * Math.PI * pathIndex / count;
                        GetOutlinePoint(shapeType, centerX, centerY, radius, angle, pathIndex, count, out targetX, out targetY);
                    }

                    targets.Add(Tuple.Create(targetX, targetY));
                }
            }

            if (!clockwise && fillShape)
                targets.Reverse();

            if (reserveCenter)
            {
                for (int i = 0; i < targets.Count; i++)
                {
                    float dx = targets[i].Item1 - centerX;
                    float dy = targets[i].Item2 - centerY;
                    if (Math.Abs(dx) < 0.01f && Math.Abs(dy) < 0.01f)
                    {
                        double angle = -Math.PI / 2 + 2.0 * Math.PI * i / Math.Max(1, targets.Count);
                        targets[i] = Tuple.Create(
                            centerX + radius * 0.35f * (float)Math.Cos(angle),
                            centerY + radius * 0.35f * (float)Math.Sin(angle));
                    }
                }
            }

            RotateTargets(targets, centerX, centerY, startAngleDegree);
            return targets;
        }

        public static void GetOutlinePoint(
            string shapeType,
            float centerX,
            float centerY,
            float radius,
            double angle,
            int itemIndex,
            int itemCount,
            out float x,
            out float y)
        {
            switch (shapeType)
            {
                case "heart":
                    x = centerX + radius * (float)(16 * Math.Pow(Math.Sin(angle), 3)) / 16f;
                    y = centerY - radius * (float)(13 * Math.Cos(angle) - 5 * Math.Cos(2 * angle) -
                        2 * Math.Cos(3 * angle) - Math.Cos(4 * angle)) / 16f;
                    return;
                case "triangle":
                    GetRegularPolygonPoint(centerX, centerY, radius, 3, itemIndex, itemCount, false, out x, out y);
                    return;
                case "square":
                    GetSquarePoint(centerX, centerY, radius, itemIndex, itemCount, out x, out y);
                    return;
                case "pentagon":
                    GetRegularPolygonPoint(centerX, centerY, radius, 5, itemIndex, itemCount, false, out x, out y);
                    return;
                case "octagon":
                    GetRegularPolygonPoint(centerX, centerY, radius, 8, itemIndex, itemCount, false, out x, out y);
                    return;
                case "star":
                    GetRegularPolygonPoint(centerX, centerY, radius, 10, itemIndex, itemCount, true, out x, out y);
                    return;
                default:
                    x = centerX + radius * (float)Math.Cos(angle - Math.PI / 2);
                    y = centerY + radius * (float)Math.Sin(angle - Math.PI / 2);
                    return;
            }
        }

        public static void RotateTargets(
            List<Tuple<float, float>> targets,
            float centerX,
            float centerY,
            float angleDegrees)
        {
            double angle = angleDegrees * Math.PI / 180.0;
            if (Math.Abs(angle) < 0.000001) return;

            for (int i = 0; i < targets.Count; i++)
            {
                double dx = targets[i].Item1 - centerX;
                double dy = targets[i].Item2 - centerY;
                targets[i] = Tuple.Create(
                    centerX + (float)(dx * Math.Cos(angle) - dy * Math.Sin(angle)),
                    centerY + (float)(dx * Math.Sin(angle) + dy * Math.Cos(angle)));
            }
        }

        public static void GetSquarePoint(
            float centerX,
            float centerY,
            float radius,
            int itemIndex,
            int itemCount,
            out float x,
            out float y)
        {
            double perimeterPosition = (double)itemIndex * 4 / itemCount;
            int firstCorner = (int)Math.Floor(perimeterPosition) % 4;
            int secondCorner = (firstCorner + 1) % 4;
            double fraction = perimeterPosition - Math.Floor(perimeterPosition);
            float[,] corners = { { 1, -1 }, { 1, 1 }, { -1, 1 }, { -1, -1 } };
            x = centerX + radius * (float)(corners[firstCorner, 0] * (1 - fraction) +
                                           corners[secondCorner, 0] * fraction);
            y = centerY + radius * (float)(corners[firstCorner, 1] * (1 - fraction) +
                                           corners[secondCorner, 1] * fraction);
        }

        public static void GetRegularPolygonPoint(float centerX, float centerY, float radius, int vertexCount, int itemIndex, int itemCount, bool alternateRadius, out float x, out float y)
        {
            double perimeterPosition = (double)itemIndex * vertexCount / itemCount;
            int firstVertex = (int)Math.Floor(perimeterPosition) % vertexCount;
            int secondVertex = (firstVertex + 1) % vertexCount;
            double fraction = perimeterPosition - Math.Floor(perimeterPosition);

            Func<int, double> vertexRadius = index => alternateRadius && index % 2 == 1 ? radius * 0.45 : radius;
            double firstAngle = -Math.PI / 2 + 2 * Math.PI * firstVertex / vertexCount;
            double secondAngle = -Math.PI / 2 + 2 * Math.PI * secondVertex / vertexCount;
            double firstRadius = vertexRadius(firstVertex);
            double secondRadius = vertexRadius(secondVertex);
            double x1 = centerX + firstRadius * Math.Cos(firstAngle);
            double y1 = centerY + firstRadius * Math.Sin(firstAngle);
            double x2 = centerX + secondRadius * Math.Cos(secondAngle);
            double y2 = centerY + secondRadius * Math.Sin(secondAngle);
            x = (float)(x1 + (x2 - x1) * fraction);
            y = (float)(y1 + (y2 - y1) * fraction);
        }
    }
}
