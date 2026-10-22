"""Generate presentation-optimized country map assets from Natural Earth 5.1.1."""

import json
import gzip
import math
import pathlib
import urllib.request


VERSION = "5.1.1"
BASE_URL = (
    "https://raw.githubusercontent.com/nvkelso/natural-earth-vector/"
    f"v{VERSION}/geojson"
)
OUTPUT_DIR = pathlib.Path(__file__).resolve().parent


def download(scale, layer="admin_0_countries"):
    url = f"{BASE_URL}/ne_{scale}_{layer}.geojson"
    with urllib.request.urlopen(url) as response:
        return json.load(response)


def outer_rings(geometry):
    if geometry["type"] == "Polygon":
        polygons = [geometry["coordinates"]]
    elif geometry["type"] == "MultiPolygon":
        polygons = geometry["coordinates"]
    else:
        return []
    return [polygon[0] for polygon in polygons if polygon and len(polygon[0]) >= 4]


def clip_ring(ring, bounds):
    min_x, min_y, max_x, max_y = bounds

    def clip(points, inside, intersection):
        if not points:
            return []
        output = []
        previous = points[-1]
        previous_inside = inside(previous)
        for current in points:
            current_inside = inside(current)
            if current_inside:
                if not previous_inside:
                    output.append(intersection(previous, current))
                output.append(current)
            elif previous_inside:
                output.append(intersection(previous, current))
            previous = current
            previous_inside = current_inside
        return output

    def vertical(a, b, x):
        delta = b[0] - a[0]
        t = 0 if abs(delta) < 1e-12 else (x - a[0]) / delta
        return [x, a[1] + t * (b[1] - a[1])]

    def horizontal(a, b, y):
        delta = b[1] - a[1]
        t = 0 if abs(delta) < 1e-12 else (y - a[1]) / delta
        return [a[0] + t * (b[0] - a[0]), y]

    points = [list(point[:2]) for point in ring]
    points = clip(points, lambda p: p[0] >= min_x, lambda a, b: vertical(a, b, min_x))
    points = clip(points, lambda p: p[0] <= max_x, lambda a, b: vertical(a, b, max_x))
    points = clip(points, lambda p: p[1] >= min_y, lambda a, b: horizontal(a, b, min_y))
    points = clip(points, lambda p: p[1] <= max_y, lambda a, b: horizontal(a, b, max_y))
    if len(points) >= 3 and points[0] != points[-1]:
        points.append(points[0])
    return points


def perpendicular_distance(point, start, end):
    dx = end[0] - start[0]
    dy = end[1] - start[1]
    if abs(dx) + abs(dy) < 1e-12:
        return math.hypot(point[0] - start[0], point[1] - start[1])
    return abs(dy * point[0] - dx * point[1] + end[0] * start[1] - end[1] * start[0]) / math.hypot(dx, dy)


def simplify(points, tolerance):
    if len(points) <= 4:
        return points
    open_points = points[:-1] if points[0] == points[-1] else points

    def rdp(segment):
        if len(segment) <= 2:
            return segment
        distance, index = max(
            (perpendicular_distance(point, segment[0], segment[-1]), index)
            for index, point in enumerate(segment[1:-1], 1)
        )
        if distance <= tolerance:
            return [segment[0], segment[-1]]
        return rdp(segment[: index + 1])[:-1] + rdp(segment[index:])

    reduced = rdp(open_points)
    if len(reduced) < 3:
        reduced = open_points[:3]
    return reduced + [reduced[0]]


def simplify_for_powerpoint(points, tolerance, max_points=475):
    """Keep a single native PowerPoint freeform below Office's stability limit."""
    reduced = simplify(points, tolerance)
    if len(reduced) - 1 <= max_points:
        return reduced

    lower = tolerance
    upper = tolerance
    while len(reduced) - 1 > max_points:
        lower = upper
        upper *= 1.5
        reduced = simplify(points, upper)

    best = reduced
    for _ in range(10):
        candidate_tolerance = (lower + upper) / 2
        candidate = simplify(points, candidate_tolerance)
        if len(candidate) - 1 <= max_points:
            upper = candidate_tolerance
            best = candidate
        else:
            lower = candidate_tolerance
    return best


def valid_aliases(properties):
    fields = (
        "ADMIN", "NAME", "NAME_LONG", "NAME_EN", "NAME_DE", "ABBREV",
        "ISO_A2", "ISO_A2_EH", "ISO_A3", "ISO_A3_EH", "ADM0_A3",
        "WB_A2", "WB_A3", "POSTAL",
    )
    aliases = []
    for field in fields:
        value = properties.get(field)
        if value and value != "-99" and value not in aliases:
            aliases.append(value)
    return aliases


def valid_admin1_aliases(properties):
    fields = (
        "name", "name_en", "name_de", "name_local", "name_alt", "abbrev",
        "postal", "iso_3166_2", "adm1_code", "code_hasc", "code_local",
        "gn_a1_code", "fips", "woe_name",
    )
    aliases = []
    for field in fields:
        value = properties.get(field)
        if value and value != "-99" and value not in aliases:
            aliases.append(str(value))
    return aliases


def project(point, bounds, aspect):
    min_lon, min_lat, max_lon, max_lat = bounds
    x = ((point[0] - min_lon) / (max_lon - min_lon)) * aspect
    y = (max_lat - point[1]) / (max_lat - min_lat)
    return [round(x, 5), round(y, 5)]


def build_asset(data, output_name, map_id, display_name, bounds, aspect, predicate, tolerance, enlarge_tiny=False, max_parts=None, max_points_per_part=None):
    regions = []
    for feature in data["features"]:
        properties = feature["properties"]
        if not predicate(properties):
            continue
        parts = []
        for ring in outer_rings(feature["geometry"]):
            clipped = clip_ring(ring, bounds)
            if len(clipped) < 4:
                continue
            projected = [project(point, bounds, aspect) for point in clipped]
            projected = (
                simplify_for_powerpoint(projected, tolerance, max_points_per_part)
                if max_points_per_part
                else simplify(projected, tolerance)
            )
            if len(projected) >= 4:
                parts.append({"Points": [{"X": p[0], "Y": p[1]} for p in projected[:-1]]})
        if not parts:
            continue

        if max_parts and len(parts) > max_parts:
            def part_area(part):
                points = part["Points"]
                return abs(sum(
                    points[index]["X"] * points[(index + 1) % len(points)]["Y"] -
                    points[(index + 1) % len(points)]["X"] * points[index]["Y"]
                    for index in range(len(points))
                ) / 2)
            parts = sorted(parts, key=part_area, reverse=True)[:max_parts]

        if enlarge_tiny:
            all_points = [(p["X"], p["Y"]) for part in parts for p in part["Points"]]
            min_x = min(point[0] for point in all_points)
            max_x = max(point[0] for point in all_points)
            min_y = min(point[1] for point in all_points)
            max_y = max(point[1] for point in all_points)
            if max(max_x - min_x, max_y - min_y) < 0.006:
                center_x = (min_x + max_x) / 2
                center_y = (min_y + max_y) / 2
                radius = 0.003
                parts = [{"Points": [
                    {"X": round(center_x - radius, 5), "Y": round(center_y - radius, 5)},
                    {"X": round(center_x + radius, 5), "Y": round(center_y - radius, 5)},
                    {"X": round(center_x + radius, 5), "Y": round(center_y + radius, 5)},
                    {"X": round(center_x - radius, 5), "Y": round(center_y + radius, 5)},
                ]}]

        region_id = properties.get("ADM0_A3")
        regions.append({
            "Id": region_id,
            "Name": properties.get("NAME_DE") or properties.get("NAME") or region_id,
            "EnglishName": properties.get("NAME_EN") or properties.get("ADMIN") or region_id,
            "Group": properties.get("CONTINENT"),
            "EntityType": properties.get("TYPE"),
            "Aliases": valid_aliases(properties),
            "Parts": parts,
        })

    regions.sort(key=lambda region: region["EnglishName"])
    asset = {
        "MapId": map_id,
        "DisplayName": display_name,
        "Source": "Natural Earth",
        "SourceVersion": VERSION,
        "SourceUrl": "https://www.naturalearthdata.com/",
        "AspectRatio": aspect,
        "CoordinateSpace": "GlobalEquirectangular",
        "MinLatitude": bounds[1],
        "MaxLatitude": bounds[3],
        "Regions": regions,
    }
    output_path = OUTPUT_DIR / output_name
    output_path.write_text(json.dumps(asset, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    point_count = sum(len(part["Points"]) for region in regions for part in region["Parts"])
    print(f"{output_path.name}: {len(regions)} regions, {point_count} points")


def build_admin1_asset(data):
    bounds = (-180.0, -90.0, 180.0, 90.0)
    regions = []
    used_ids = set()
    counts = {}

    for feature in data["features"]:
        properties = feature["properties"]
        country_id = properties.get("adm0_a3")
        if not country_id or country_id == "-99":
            continue

        parts = []
        for ring in outer_rings(feature["geometry"]):
            clipped = clip_ring(ring, bounds)
            if len(clipped) < 4:
                continue
            projected = [project(point, bounds, 2.0) for point in clipped]
            projected = simplify_for_powerpoint(projected, 0.0001, 475)
            if len(projected) >= 4:
                parts.append({"Points": [{"X": p[0], "Y": p[1]} for p in projected[:-1]]})
        if not parts:
            continue

        candidates = (
            properties.get("iso_3166_2"),
            properties.get("adm1_code"),
            properties.get("code_hasc"),
            properties.get("gn_a1_code"),
            str(properties.get("ne_id") or ""),
        )
        region_id = next((str(value) for value in candidates if value and str(value) != "-99"), None)
        if not region_id:
            continue
        base_id = region_id
        suffix = 2
        while region_id in used_ids:
            region_id = f"{base_id}-{suffix}"
            suffix += 1
        used_ids.add(region_id)

        local_name = properties.get("name_de") or properties.get("name_local") or properties.get("name") or region_id
        english_name = properties.get("name_en") or properties.get("name") or local_name
        regions.append({
            "Id": region_id,
            "Name": local_name,
            "EnglishName": english_name,
            "Group": country_id,
            "EntityType": properties.get("type_en") or properties.get("type"),
            "Aliases": valid_admin1_aliases(properties),
            "Parts": parts,
        })
        counts[country_id] = counts.get(country_id, 0) + 1

    regions.sort(key=lambda region: (region["Group"], region["EnglishName"]))
    asset = {
        "MapId": "ADMIN1-GLOBAL",
        "DisplayName": "States and Provinces",
        "Source": "Natural Earth",
        "SourceVersion": VERSION,
        "SourceUrl": "https://www.naturalearthdata.com/",
        "AspectRatio": 2.0,
        "CoordinateSpace": "GlobalEquirectangular",
        "MinLatitude": -90.0,
        "MaxLatitude": 90.0,
        "Regions": regions,
    }
    payload = json.dumps(asset, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    asset_path = OUTPUT_DIR / "states_provinces_admin1_naturalearth_5_1_1.json.gz"
    asset_path.write_bytes(gzip.compress(payload, compresslevel=9, mtime=0))

    catalog_path = OUTPUT_DIR / "states_provinces_admin1_catalog_naturalearth_5_1_1.tsv"
    catalog_path.write_text(
        "".join(f"{country_id}\t{count}\n" for country_id, count in sorted(counts.items())),
        encoding="utf-8",
    )
    point_count = sum(len(part["Points"]) for region in regions for part in region["Parts"])
    print(
        f"{asset_path.name}: {len(regions)} regions across {len(counts)} countries, "
        f"{point_count} points, {len(payload)} bytes raw, {asset_path.stat().st_size} bytes gzip"
    )


def main():
    overview_countries = download("50m")
    build_asset(
        overview_countries,
        "countries_overview_admin0_naturalearth_5_1_1.json",
        "COUNTRIES-OVERVIEW-ADM0",
        "Countries and Territories – Overview",
        (-180.0, -90.0, 180.0, 90.0),
        2.0,
        lambda properties: True,
        # About 0.027 degrees at the equator. This keeps shared borders and
        # coastlines crisp even when a continent is enlarged on a slide.
        0.00015,
        enlarge_tiny=True,
        max_parts=16,
    )
    detailed_countries = download("10m")
    build_asset(
        detailed_countries,
        "countries_admin0_naturalearth_5_1_1.json",
        "COUNTRIES-ADM0",
        "Countries and Territories",
        (-180.0, -90.0, 180.0, 90.0),
        2.0,
        lambda properties: True,
        # Natural Earth 1:10m plus roughly 0.018 degrees of simplification.
        # This more than doubles Germany's former boundary detail while keeping
        # every PowerPoint freeform below the complexity that destabilizes Office.
        0.0001,
        enlarge_tiny=True,
        max_points_per_part=475,
    )
    admin1 = download("10m", "admin_1_states_provinces")
    build_admin1_asset(admin1)


if __name__ == "__main__":
    main()
