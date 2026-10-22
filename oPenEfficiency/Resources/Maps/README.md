# Bundled map data

`de_adm1_naturalearth_5_1_1.json` contains presentation-optimized geometry for
Germany's first-level administrative regions. It was derived from Natural Earth
Admin 1 – States, Provinces version 5.1.1.

`countries_admin0_naturalearth_5_1_1.json` contains the detailed Natural Earth
1:10m Admin 0 country layer used for individual country maps.
`countries_overview_admin0_naturalearth_5_1_1.json` contains the Natural Earth
1:50m layer with 242 catalog entries, stronger slide-scale simplification, and at
most sixteen polygon parts per entity. It keeps world and continent insertion
responsive while preserving independently editable country shapes.
Presentation-sized markers preserve otherwise invisible microstates. Continent
and entity-type metadata let the catalog filter both assets without a GIS runtime.
Detailed rings are capped at 475 nodes per native freeform to avoid known Office
instability while retaining the additional 1:10m boundary detail.

`states_provinces_admin1_naturalearth_5_1_1.json.gz` contains 4,596 first-level
administrative regions across 251 countries and territories from Natural Earth's
1:10m Admin 1 – States, Provinces layer. It uses the same 475-node freeform limit
as detailed countries and is gzip-compressed for embedding. The accompanying TSV
catalog lets the wizard enable the detail option without loading the full geometry.
The remote-region renderer derives its mainland reference from the bundled ADM0
geometry. At insertion time it can keep remote ADM1 regions in place, reposition
them as compact insets, or omit them from the inserted map. Neither option duplicates
geometry or increases the embedded data size.

`de_adm2_government_districts_geoboundaries_2021.json.gz` contains Germany's 38
second-level government districts. It is derived from geoBoundaries' 2021 Germany
ADM2 release, whose stated source is the Federal Agency for Cartography and
Geodesy, under Data licence Germany – Attribution – Version 2.0. The wizard can
insert all districts or only those belonging to a selected federal state.

`de_adm3_districts_geoboundaries_2021.json.gz` contains Germany's 428 ADM3
districts and independent cities. It uses the same geoBoundaries/BKG 2021 source
and licence as the ADM2 asset, and is intentionally inserted only for one selected
federal state at a time.

`administrative_regions_adm2_geoboundaries_current.json.gz` contains 48,234
second-level regions across 180 countries from geoBoundaries gbOpen. It is prepared
for the global ADM2 Map Wizard rollout; countries without an available source dataset
remain unavailable at this detail level.

geoBoundaries gbOpen is licensed under Creative Commons Attribution 4.0
International (CC BY 4.0). The bundled files are modified derivatives: their
geometry is projected, precision-reduced, simplified, normalized, stripped of
unused source attributes, linked to parent regions, converted to the application
schema, and compressed. Required attribution, source links, change notices, and
the suggested citation are recorded in the repository-level
`THIRD_PARTY_NOTICES.md`, which is included in release payloads.

The global ADM2 and ADM3 assets attach lower-level regions to the bundled first-level
catalog where a spatial parent match is available. This lets the wizard insert a
single first-level region's children instead of the whole country. Run
`assign_global_parent_groups.py` after regenerating either global asset to rebuild
those links.

`administrative_regions_adm3_geoboundaries_current.json.gz` contains 105,035
third-level regions across 81 countries from geoBoundaries gbOpen. As with global
ADM2, the wizard enables this option only for covered countries and releases its
decompressed geometry after five minutes of inactivity.

`map_overviews_mercator_naturalearth_5_1_1.pptx` and
`map_overviews_naturalearth_5_1_1.pptx` contain the world and seven continent
maps as prebuilt native PowerPoint shape groups. Mercator is the insertion default;
the natural/equirectangular asset preserves the optional previous appearance. At
runtime the requested group is copied into the active slide, which avoids rebuilding
thousands of freeform points. The overview JSON remains the automatic fallback.
Run `generate_overview_pptx.cs` against a blank presentation and pass `Mercator`
or `Natural` to regenerate the corresponding embedded asset.

Natural Earth data is in the public domain:
https://www.naturalearthdata.com/about/terms-of-use/

The geometry is projected at insertion time, precision-simplified, normalized,
and stripped of GIS attributes that are not needed by the PowerPoint renderer.
Run `generate_country_assets.py` to reproduce the country and global ADM1 assets
from the tagged Natural Earth source data. The older Germany-only ADM1 asset remains
embedded solely as a compatibility fallback.

Current bundled sizes:

- Detailed country JSON: 2,425,913 bytes (90,161 boundary points)
- Compressed states/provinces JSON: 1,464,572 bytes (7,292,179 bytes raw; 241,685 boundary points)
- States/provinces availability catalog: 1,912 bytes
- Overview JSON fallback: 1,049,182 bytes (38,560 boundary points)
- Prebuilt Mercator world and continent PowerPoint asset: 1,472,909 bytes
- Prebuilt Natural world and continent PowerPoint asset: 1,443,568 bytes
- Reserved Germany ADM1 JSON: 57,626 bytes
