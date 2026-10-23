using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Xml.Linq;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.UI;
using oPenEfficiency.Utils;

internal static class IconComposerExportTests
{
    private static int tests;
    private static string root;
    [STAThread]
    private static int Main(string[] args)
    {
        root = args[0];
        try
        {
            if (args.Length > 1 && args[1] == "--production-files")
            {
                ParseProductionFiles(args[2]);
                Console.WriteLine("ICON_BUNDLE_PRODUCTION_PARSE_PASS files=" + Directory.GetFiles(args[2], "*.svg").Length + " tests=" + tests);
                return 0;
            }
            Units();
            if (args.Length > 1) throw new ArgumentException("Use --production-files with a production SVG directory.");
            Console.WriteLine("ICON_EXPORT_TESTS_PASS tests=" + tests);
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }

    private static void ParseProductionFiles(string directory)
    {
        string[] files = Directory.GetFiles(directory, "*.svg");
        Check(files.Length > 0, "bundle production SVG files exist");
        foreach (string file in files)
        {
            string svg = File.ReadAllText(file);
            Check(!svg.Contains("<mask") && !svg.Contains("mask="), "bundle output has no masks: " + Path.GetFileName(file));
            ParsedIconGeometry parsed = Parse(svg);
            Check(parsed != null && parsed.Elements.Count > 0, "production parser accepts bundle output: " + Path.GetFileName(file));
            foreach (ParsedIconElement element in parsed.Elements)
            {
                Check(element.Geometry != null && !element.Geometry.Bounds.IsEmpty, "non-empty WPF geometry: " + Path.GetFileName(file));
            }
            Check(parsed.Elements.Any(element => element.IsAccent && element.AccentRole == FeatureIconAccentRole.CreateGreen),
                "semantic CreateGreen accent survives production parse: " + Path.GetFileName(file));
        }
    }

    private static void Check(bool value, string name)
    {
        tests++;
        if (!value) throw new Exception("FAIL " + name);
    }

    private static string Wrap(string body)
    { return "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 24 24\">" + body + "</svg>"; }

    private static void Reject(string body, string name)
    {
        bool rejected = false;
        try { IconComposerSvgExporter.Export(body); } catch { rejected = true; }
        Check(rejected, name);
    }

    private static void Units()
    {
        string mask = "<defs><mask id=\"hole\" maskUnits=\"userSpaceOnUse\"><rect width=\"24\" height=\"24\" fill=\"white\"/><circle cx=\"12\" cy=\"12\" r=\"3\" fill=\"black\"/></mask></defs>";
        string source = Wrap(mask + "<rect width=\"24\" height=\"24\" mask=\"url(#hole)\"/>");
        string output = IconComposerSvgExporter.Export(source);
        Check(output == IconComposerSvgExporter.Export(source), "deterministic bytes");
        Check(!output.Contains("<mask") && !output.Contains("mask="), "mask-free output");
        ParsedIconGeometry parsed = Parse(output);
        Check(!parsed.BaseGeometry.FillContains(new Point(12, 12)), "transparent mask hole");
        Check(parsed.BaseGeometry.FillContains(new Point(2, 2)), "positive control survives");
        string moved = IconComposerSvgExporter.Export(Wrap(mask + "<g transform=\"translate(2 1)\"><rect x=\"1\" y=\"1\" width=\"18\" height=\"18\" mask=\"url(#hole)\"/></g>"));
        Check(!Parse(moved).BaseGeometry.FillContains(new Point(14, 13)), "mask transforms with subject");
        string stroked = IconComposerSvgExporter.Export(Wrap("<path fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" d=\"M2 2H20\"/>"));
        Check(Parse(stroked).BaseGeometry.FillContains(new Point(10, 2.5)), "stroke expansion");
        Check(!Parse(stroked).BaseGeometry.FillContains(new Point(10, 4)), "stroke positive boundary");
        string nonzero = IconComposerSvgExporter.Export(Wrap("<path fill-rule=\"nonzero\" d=\"M2 2H14V14H2Z M8 8H20V20H8Z\"/>"));
        Check(Parse(nonzero).BaseGeometry.FillContains(new Point(10, 10)), "nonzero overlap survives runtime EvenOdd default");
        string ring = IconComposerSvgExporter.Export(Wrap("<path fill-rule=\"nonzero\" d=\"M2 2H20V20H2Z M8 8V14H14V8Z\"/>"));
        Check(!Parse(ring).BaseGeometry.FillContains(new Point(10, 10)), "nonzero ring hole survives runtime EvenOdd default");
        string roles = IconComposerSvgExporter.Export(Wrap("<rect x=\"1\" y=\"1\" width=\"2\" height=\"2\"/><g data-oe-role=\"accent\" data-oe-accent-role=\"SuccessGreen\"><circle cx=\"12\" cy=\"12\" r=\"3\"/></g>"));
        Check(Parse(roles).Elements.Any(e => e.IsAccent && e.AccentRole == FeatureIconAccentRole.SuccessGreen), "semantic accent retained");
        string reset = IconComposerSvgExporter.Export(Wrap("<g data-oe-role=\"accent\" data-oe-accent-role=\"SuccessGreen\"><rect x=\"1\" y=\"1\" width=\"2\" height=\"2\" data-oe-role=\"base\"/><circle cx=\"12\" cy=\"12\" r=\"3\"/></g>"));
        Check(Parse(reset).BaseGeometry.FillContains(new Point(2, 2)), "nested explicit base reset retained");
        Check(Parse(reset).Elements.Any(e => e.IsAccent && e.AccentRole == FeatureIconAccentRole.SuccessGreen), "base reset keeps sibling accent");
        string content = IconComposerSvgExporter.Export(Wrap("<rect width=\"2\" height=\"2\"/><rect x=\"5\" width=\"3\" height=\"3\" fill=\"#12ab34\" data-oe-role=\"accent\" data-oe-accent-role=\"ContentColor\"/>"));
        Check(Parse(content).Elements.Any(e => e.AccentRole == FeatureIconAccentRole.ContentColor && e.FillPaint.Equals("#12ab34", StringComparison.OrdinalIgnoreCase)), "ContentColor sample retained");
        CheckContentColorStates(Parse(content));
        Reject(Wrap(""), "empty export rejected");
        Reject(Wrap("<rect width=\"0\" height=\"2\"/>"), "zero-size export rejected");
        Reject(Wrap("<rect x=\"30\" width=\"2\" height=\"2\"/>"), "clipped-empty export rejected");
        Reject(Wrap("<rect width=\"2\" height=\"2\" data-oe-role=\"accent\" data-oe-accent-role=\"ActionBlue\"/>"), "accent-only export rejected");
        Reject(Wrap("<image href=\"remote.png\"/>"), "unsupported image rejected");
        Reject(Wrap("<rect width=\"2\" height=\"2\" filter=\"url(#blur)\"/>"), "filter rejected");
        Reject(Wrap("<rect width=\"2\" height=\"2\" clip-path=\"url(#clip)\"/>"), "clip rejected");
        Reject(Wrap("<rect width=\"2\" height=\"2\" opacity=\".5\"/>"), "partial opacity rejected");
        Reject(Wrap("<rect width=\"1\" height=\"1\"/><g fill=\"#123456\" data-oe-role=\"accent\" data-oe-accent-role=\"ContentColor\"><rect x=\"3\" width=\"2\" height=\"2\" fill=\"#ABCDEF\" style=\"fill:inherit\"/></g>"), "ambiguous CSS inherit rejected");
        Reject(Wrap("<rect width=\"2\" height=\"2\" mask=\"url(#missing)\"/>"), "unresolved mask rejected");
        Reject(Wrap("<rect width=\"2\" height=\"2\"/>").Replace("viewBox=", "preserveAspectRatio=\"none\" viewBox="), "unsupported root aspect ratio rejected");
        Reject(Wrap("<rect width=\"2\" height=\"2\"/>").Replace("viewBox=", "width=\"48\" height=\"24\" viewBox="), "root viewport distortion rejected");
        Reject(Wrap("<rect width=\"2\" height=\"2\"/>").Replace("viewBox=", "width=\"-1\" viewBox="), "invalid root viewport rejected");
        Reject(Wrap("<svg x=\"4\" y=\"4\" width=\"4\" height=\"4\" overflow=\"visible\"><rect width=\"10\" height=\"10\"/></svg>"), "unclipped nested viewport rejected");
        Reject("<!DOCTYPE svg [<!ENTITY x SYSTEM 'file:///missing'>]>" + Wrap("&x;"), "DTD rejected");
        foreach (string relative in new[] { "tools/icon-composer/kit/objects/freeform-nodes.svg", "tools/icon-composer/kit/objects/number-format.svg", "tools/icon-composer/tests/fixtures/pause-resume-A.svg" })
        {
            string actual = IconComposerSvgExporter.Export(File.ReadAllText(Path.Combine(root, relative)));
            Check(!actual.Contains("<mask") && actual.Contains("<path"), "approved public source exports: " + relative);
        }
    }

    private static void CheckContentColorStates(ParsedIconGeometry parsed)
    {
        var paths = parsed.Elements.Select(e => FeatureIconPathFactory.CreatePath(e, parsed.ViewBox)).ToList();
        int sample = parsed.Elements.ToList().FindIndex(e => e.AccentRole == FeatureIconAccentRole.ContentColor);
        foreach (string theme in new[] { "Light", "Dark" })
        {
            var resources = Theme(theme);
            foreach (var path in paths)
                foreach (var resource in resources) path.Resources[resource.Key] = resource.Value;
            var featureBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#A855F7"));
            FeatureIconPathFactory.ApplyState(paths, parsed, IconColorModes.PowerPoint, FeatureIconVisualState.Normal, featureBrush);
            Check(((SolidColorBrush)paths[sample].Fill).Color == (Color)ColorConverter.ConvertFromString("#12ab34"), "ContentColor normal " + theme);
            FeatureIconPathFactory.ApplyState(paths, parsed, IconColorModes.PowerPoint, FeatureIconVisualState.Disabled, featureBrush);
            Check(((SolidColorBrush)paths[sample].Fill).Color == ((SolidColorBrush)resources["OfficeIconDisabledBrush"]).Color, "ContentColor disabled " + theme);
            FeatureIconPathFactory.ApplyState(paths, parsed, IconColorModes.FeatureColors, FeatureIconVisualState.Normal, featureBrush);
            Check(((SolidColorBrush)paths[sample].Fill).Color == featureBrush.Color, "ContentColor feature color " + theme);
        }
    }

    private static ParsedIconGeometry Parse(string svg)
    {
        string file = Path.GetTempFileName();
        try
        {
            File.WriteAllText(file, svg);
            var parsed = SvgParser.ParseSvgWithRoles(file);
            if (parsed == null) throw new Exception("Production parser rejected exported SVG");
            return parsed;
        }
        finally { File.Delete(file); }
    }

    private static Dictionary<string, Brush> Theme(string name)
    {
        var result = new Dictionary<string, Brush>();
        XNamespace x = "http://schemas.microsoft.com/winfx/2006/xaml";
        var document = XDocument.Load(Path.Combine(root, "oPenEfficiency/UI/Styles/Themes", name + ".xaml"));
        foreach (var element in document.Descendants().Where(e => e.Name.LocalName == "Color"))
        {
            string key = (string)element.Attribute(x + "Key");
            if (key != null) result[key.Replace("Color", "Brush")] = new SolidColorBrush((Color)ColorConverter.ConvertFromString(element.Value));
        }
        return result;
    }

}
