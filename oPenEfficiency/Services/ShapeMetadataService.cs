using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.Services
{
    public static class ShapeMetadataService
    {
        public static void SetTag(Shape shape, string key, string value)
        {
            if (shape == null || string.IsNullOrEmpty(key)) return;
            try
            {
                shape.Tags.Add(key, value);
            }
            catch (Exception) { }
        }

        public static string GetTag(Shape shape, string key)
        {
            if (shape == null || string.IsNullOrEmpty(key)) return null;
            try
            {
                for (int i = 1; i <= shape.Tags.Count; i++)
                {
                    if (shape.Tags.Name(i).Equals(key, StringComparison.OrdinalIgnoreCase))
                    {
                        return shape.Tags.Value(i);
                    }
                }
            }
            catch (Exception) { }
            return null;
        }

        public static bool HasTag(Shape shape, string key)
        {
            return GetTag(shape, key) != null;
        }

        public static void RemoveTag(Shape shape, string key)
        {
            if (shape == null || string.IsNullOrEmpty(key)) return;
            try
            {
                for (int i = 1; i <= shape.Tags.Count; i++)
                {
                    if (shape.Tags.Name(i).Equals(key, StringComparison.OrdinalIgnoreCase))
                    {
                        shape.Tags.Delete(shape.Tags.Name(i));
                        return;
                    }
                }
            }
            catch (Exception) { }
        }

        public static bool HasOpeTags(Shape shape)
        {
            if (shape == null) return false;
            try
            {
                for (int i = 1; i <= shape.Tags.Count; i++)
                {
                    if (shape.Tags.Name(i).StartsWith("oPE_", StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }
            catch (Exception) { }
            return false;
        }

        public static bool HasThirdPartyTags(Shape shape)
        {
            return HasTag(shape, "EE4P_SMART_ELEMENT") || HasTag(shape, "THINKCELLSHAPEDONOTDELETE");
        }
    }
}
