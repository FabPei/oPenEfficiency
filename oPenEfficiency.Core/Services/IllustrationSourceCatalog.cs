using System;
using System.Collections.Generic;
using System.Linq;
using oPenEfficiency.Models;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// The evaluated illustration sources behind the Illustration Library window. Every entry was
    /// checked against the source's own license text and against a downloaded sample; the evaluation
    /// itself is written down in docs/product/feat-011-illustration-library-todo.md.
    ///
    /// Only the entries marked <see cref="IllustrationSourceStatus.Bundled"/> travel inside the
    /// add-in. Everything else is a link the user follows on their own responsibility: the add-in
    /// downloads nothing automatically and ships nothing from those sources.
    /// </summary>
    public static class IllustrationSourceCatalog
    {
        private const string Cc0Url = "https://creativecommons.org/publicdomain/zero/1.0/";

        /// <summary>
        /// Shown in the window and repeated in documentation: the add-in bundles CC0 artwork only and
        /// every other source stays the user's own licensing decision.
        /// </summary>
        public const string ResponsibilityNotice =
            "Only the collections marked \"Included in the add-in\" travel inside this add-in; they are CC0 and may be " +
            "used, changed and redistributed without credit. Every other source below is a link: nothing is downloaded " +
            "automatically and nothing is shipped with the add-in. When you import artwork from those sources you are " +
            "responsible for the license yourself - read the source's terms, keep any credit the license requires in " +
            "your presentation, and do not redistribute material whose terms forbid it. oPenEfficiency gives no legal " +
            "advice and no warranty for imported artwork.";

        private static readonly IllustrationSource[] Sources =
        {
            new IllustrationSource
            {
                Id = "open-peeps",
                Name = "Open Peeps",
                HomePageUrl = "https://www.openpeeps.com/",
                LicenseName = "CC0 1.0 Universal",
                LicenseUrl = Cc0Url,
                Status = IllustrationSourceStatus.Bundled,
                AttributionRequired = false,
                BundledFigureCount = 93,
                AcquisitionNote = "Already part of the add-in. Originals are published on the Open Peeps site and in the Figma community.",
                TechnicalNote = "Mix-and-match people with named parts (body, head, face, facial hair, accessories) - the only bundled pack that can be split into parts and recoloured per part."
            },
            new IllustrationSource
            {
                Id = "kitbitz",
                Name = "Kitbitz",
                HomePageUrl = "https://kitbitz.art/",
                LicenseName = "CC0 1.0 Universal",
                LicenseUrl = "https://kitbitz.art/docs/license",
                Status = IllustrationSourceStatus.Bundled,
                AttributionRequired = false,
                BundledFigureCount = 325,
                AcquisitionNote = "Three kits (space, food, winter) are part of the add-in. The full 15-kit library is at github.com/CaptExcellent/kits-library-assets.",
                TechnicalNote = "Single props with inline hex colours and no named groups - recolour, do not split."
            },
            new IllustrationSource
            {
                Id = "open-doodles",
                Name = "Open Doodles",
                HomePageUrl = "https://www.opendoodles.com/",
                LicenseName = "CC0 1.0 Universal",
                LicenseUrl = Cc0Url,
                Status = IllustrationSourceStatus.Bundled,
                AttributionRequired = false,
                BundledFigureCount = 3,
                AcquisitionNote = "Part of the add-in; the complete set is on the Open Doodles site.",
                TechnicalNote = "One ink group plus an accent colour, no named parts - recolour, do not split."
            },
            new IllustrationSource
            {
                Id = "humaaans",
                Name = "Humaaans",
                HomePageUrl = "https://www.humaaans.com/",
                LicenseName = "CC0 1.0 Universal (origin)",
                LicenseUrl = Cc0Url,
                Status = IllustrationSourceStatus.BundlingPermitted,
                AttributionRequired = false,
                AcquisitionNote = "Free download from the origin: Gumroad \"$0\" checkout with the e-mail form (manual step). A public mirror exists, but it labels itself CC BY 4.0 - use the origin files.",
                TechnicalNote = "Named parts (Head, Body, Bottom, Accessories, Pose) plus ready-made single pieces - the best remaining candidate for a split-and-recolour pack."
            },
            new IllustrationSource
            {
                Id = "ctrlv.design",
                Name = "ctrlv.design",
                HomePageUrl = "https://ctrlv.design/",
                LicenseName = "CC0 1.0 Universal",
                LicenseUrl = "https://ctrlv.design/LICENSE",
                Status = IllustrationSourceStatus.BundlingPermitted,
                AttributionRequired = false,
                AcquisitionNote = "Download from the site's gallery; the site's own export step writes hex colours into the file.",
                TechnicalNote = "Raw files colour through CSS variables and carry no named groups - after export they recolour, but they cannot be split."
            },
            new IllustrationSource
            {
                Id = "cocomaterial",
                Name = "CocoMaterial",
                HomePageUrl = "https://cocomaterial.com/",
                LicenseName = "CC0 1.0 Universal",
                LicenseUrl = Cc0Url,
                Status = IllustrationSourceStatus.BundlingPermitted,
                AttributionRequired = false,
                AcquisitionNote = "Download individually on the site or through its public vector API.",
                TechnicalNote = "Two-tone line figures with short hex values, no named parts - recolour, do not split."
            },
            new IllustrationSource
            {
                Id = "scale-flexiple",
                Name = "Scale (Flexiple)",
                HomePageUrl = "https://scale.flexiple.com/",
                LicenseName = "Custom - free to download, tweak and share",
                LicenseUrl = "https://scale.flexiple.com/",
                Status = IllustrationSourceStatus.BundlingPermitted,
                AttributionRequired = false,
                AcquisitionNote = "Download directly on scale.flexiple.com.",
                TechnicalNote = "Inline hex colours, no named groups. The terms ask you not to sell the artwork as an illustration pack."
            },
            new IllustrationSource
            {
                Id = "ira-design",
                Name = "IRA Design",
                HomePageUrl = "https://iradesign.io/",
                LicenseName = "MIT",
                LicenseUrl = "https://iradesign.io/",
                Status = IllustrationSourceStatus.BundlingPermitted,
                AttributionRequired = true,
                AcquisitionNote = "Download the packs and carry the MIT notice with them.",
                TechnicalNote = "Characters, objects and backgrounds ship as separate files, so parts come from the file layout; colours live in CSS classes, so recolouring needs a class resolution step first."
            },
            new IllustrationSource
            {
                Id = "lukaszadam",
                Name = "Lukasz Adam Illustrations",
                HomePageUrl = "https://lukaszadam.com/illustrations",
                LicenseName = "CC0 as stated on the site (labelled \"CC0 license (MIT license)\")",
                LicenseUrl = "https://lukaszadam.com/illustrations",
                Status = IllustrationSourceStatus.BundlingPermitted,
                AttributionRequired = false,
                AcquisitionNote = "Download the bundle (Gumroad, $0).",
                TechnicalNote = "Colours live in CSS classes and gradients, no named parts - recolouring needs class resolution."
            },
            new IllustrationSource
            {
                Id = "highlights",
                Name = "Highlights (Outdraw Design)",
                HomePageUrl = "https://outdrawdesign.gumroad.com/l/highlights",
                LicenseName = "CC0 1.0 Universal",
                LicenseUrl = "https://outdrawdesign.gumroad.com/l/highlights",
                Status = IllustrationSourceStatus.BundlingPermitted,
                AttributionRequired = false,
                AcquisitionNote = "Free download through Gumroad or the Figma community.",
                TechnicalNote = "Hand-drawn highlighter shapes (arrows, loops, blobs) rather than figures, and only black strokes - useful as decorations, not as characters."
            },
            new IllustrationSource
            {
                Id = "illlustrations-co",
                Name = "illlustrations.co",
                HomePageUrl = "https://illlustrations.co/",
                LicenseName = "Split licensing - CC0 for legacy packs, MIT for the repository, premium packs restricted",
                LicenseUrl = "https://illlustrations.co/",
                Status = IllustrationSourceStatus.RedistributionRestricted,
                AttributionRequired = false,
                AcquisitionNote = "Download packs on the site and check each pack's own license before passing files on.",
                TechnicalNote = "Colour handling mixes CSS classes with hex, and at least one sample embeds a raster image - verify every file you import."
            },
            new IllustrationSource
            {
                Id = "storyset",
                Name = "Storyset (Freepik)",
                HomePageUrl = "https://storyset.com/",
                LicenseName = "Storyset/Freepik terms - attribution required, inclusion in packs or libraries forbidden",
                LicenseUrl = "https://storyset.com/",
                Status = IllustrationSourceStatus.RedistributionRestricted,
                AttributionRequired = true,
                AcquisitionNote = "Customise and download single illustrations on the site.",
                TechnicalNote = "Named groups and inline hex colours, so it edits well - but the terms allow single use with credit only, never redistribution inside a library."
            },
            new IllustrationSource
            {
                Id = "undraw",
                Name = "unDraw",
                HomePageUrl = "https://undraw.co/",
                LicenseName = "unDraw license - no attribution, redistribution in packs forbidden",
                LicenseUrl = "https://undraw.co/license",
                Status = IllustrationSourceStatus.RedistributionRestricted,
                AttributionRequired = false,
                AcquisitionNote = "Customise the accent colour and download single illustrations on the site.",
                TechnicalNote = "No named parts, and the license forbids distributing the assets in packs or fetching them automatically."
            },
            new IllustrationSource
            {
                Id = "manypixels",
                Name = "ManyPixels",
                HomePageUrl = "https://www.manypixels.co/gallery",
                LicenseName = "ManyPixels license - redistribution in packs and integrations forbidden",
                LicenseUrl = "https://www.manypixels.co/gallery",
                Status = IllustrationSourceStatus.RedistributionRestricted,
                AttributionRequired = false,
                AcquisitionNote = "Download single illustrations from the gallery.",
                TechnicalNote = "Recolouring works, but the terms forbid passing the artwork on in packs or offering it as part of an integration."
            },
            new IllustrationSource
            {
                Id = "drawkit",
                Name = "DrawKit",
                HomePageUrl = "https://www.drawkit.com/",
                LicenseName = "DrawKit license - no stand-alone or compilation redistribution",
                LicenseUrl = "https://www.drawkit.com/license",
                Status = IllustrationSourceStatus.RedistributionRestricted,
                AttributionRequired = false,
                AcquisitionNote = "Create an account and download the free packs on the site.",
                TechnicalNote = "Recolouring works; redistribution as a compilation is prohibited."
            },
            new IllustrationSource
            {
                Id = "pixeltrue",
                Name = "PixelTrue Free",
                HomePageUrl = "https://www.pixeltrue.com/free-illustrations",
                LicenseName = "PixelTrue free license - no offering of the downloaded files to third parties",
                LicenseUrl = "https://www.pixeltrue.com/free-illustrations",
                Status = IllustrationSourceStatus.RedistributionRestricted,
                AttributionRequired = false,
                AcquisitionNote = "Download the free packs on the site.",
                TechnicalNote = "Use inside your own presentation is fine; handing the files to others is not."
            },
            new IllustrationSource
            {
                Id = "sketchvalley",
                Name = "SketchValley",
                HomePageUrl = "https://sketchvalley.com/",
                LicenseName = "SketchValley license - no redistribution grant, stand-alone resale forbidden",
                LicenseUrl = "https://sketchvalley.com/",
                Status = IllustrationSourceStatus.RedistributionRestricted,
                AttributionRequired = false,
                AcquisitionNote = "Download on the site (the site blocks automated requests).",
                TechnicalNote = "Colours are class or rgb() based; no named groups."
            },
            new IllustrationSource
            {
                Id = "wdceee",
                Name = "WDCEEE",
                HomePageUrl = "https://wdceee.com/free-svg-illustrations/",
                LicenseName = "Unclear - the site mixes \"MIT licensed\" with \"under the C00 license\"",
                LicenseUrl = "https://wdceee.com/free-svg-illustrations/",
                Status = IllustrationSourceStatus.RedistributionRestricted,
                AttributionRequired = false,
                AcquisitionNote = "Download on the site; ask the author before passing files on.",
                TechnicalNote = "Technically usable, but the license wording contradicts itself, so redistribution cannot be justified from the source text."
            },
            new IllustrationSource
            {
                Id = "desi-log",
                Name = "Desi Log",
                HomePageUrl = "https://desilog.com/",
                LicenseName = "Custom - credit and link required wherever redistributed, resale forbidden",
                LicenseUrl = "https://desilog.com/",
                Status = IllustrationSourceStatus.RedistributionRestricted,
                AttributionRequired = true,
                AcquisitionNote = "Shared through the author's site; the free tier ships no SVG files.",
                TechnicalNote = "SVG availability depends on the author; credit must survive wherever the artwork appears."
            },
            new IllustrationSource
            {
                Id = "abstract-lab",
                Name = "Abstract Lab",
                HomePageUrl = "https://www.abstractlab.co/",
                LicenseName = "CC0 per the site FAQ, but only for the free PNG tier",
                LicenseUrl = "https://www.abstractlab.co/",
                Status = IllustrationSourceStatus.RedistributionRestricted,
                AttributionRequired = false,
                AcquisitionNote = "Free downloads are PNG; the SVG packs are paid products.",
                TechnicalNote = "No acquirable vector artwork, so nothing can be imported as an editable illustration."
            },
            new IllustrationSource
            {
                Id = "open-illustrations",
                Name = "Open Illustrations",
                HomePageUrl = "https://www.figma.com/community/plugin/1374259595504988210/open-illustrations",
                LicenseName = "CC0 claimed by the plugin, not documented elsewhere",
                LicenseUrl = "https://www.figma.com/community/plugin/1374259595504988210/open-illustrations",
                Status = IllustrationSourceStatus.RedistributionRestricted,
                AttributionRequired = false,
                AcquisitionNote = "Only reachable through the Figma plugin (Figma account required); there is no download page.",
                TechnicalNote = "Provenance of the CC0 claim is undocumented, so treat exported files as unverified before you rely on them."
            }
        };

        /// <summary>Every evaluated source, bundled ones first, then alphabetically within a status group.</summary>
        public static IReadOnlyList<IllustrationSource> All
        {
            get
            {
                return Sources
                    .OrderBy(source => source.Status)
                    .ThenBy(source => source.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
        }

        /// <summary>The collections that ship inside the add-in.</summary>
        public static IReadOnlyList<IllustrationSource> Bundled
        {
            get { return All.Where(source => source.IsBundled).ToList(); }
        }

        /// <summary>Sources whose own terms permit redistribution, even though nothing is shipped yet.</summary>
        public static IReadOnlyList<IllustrationSource> BundlingPermitted
        {
            get { return All.Where(source => source.Status == IllustrationSourceStatus.BundlingPermitted).ToList(); }
        }

        public static IllustrationSource Find(string id)
        {
            if (string.IsNullOrWhiteSpace(id)) return null;

            return All.FirstOrDefault(source => string.Equals(source.Id, id, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>Total figures bundled inside the add-in across all sources.</summary>
        public static int BundledFigureTotal
        {
            get { return All.Sum(source => source.BundledFigureCount); }
        }
    }
}
