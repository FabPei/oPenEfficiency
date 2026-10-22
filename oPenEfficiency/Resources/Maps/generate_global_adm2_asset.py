"""Build the bundled global ADM2 asset from geoBoundaries gbOpen country files."""

import gzip
import json
import pathlib
import time
import urllib.request
from shapely.geometry import Point, MultiPolygon, Polygon
from shapely.prepared import prep

from generate_germany_adm2_asset import project, rings, simplify_for_powerpoint

CATALOG_URL = "https://www.geoboundaries.org/api/current/gbOpen/ALL/ADM2/"
ADMIN1_ASSET = pathlib.Path(__file__).resolve().parent / "states_provinces_admin1_naturalearth_5_1_1.json.gz"
OUTPUT = pathlib.Path(__file__).resolve().parent / "administrative_regions_adm2_geoboundaries_current.json.gz"

def download_json(url):
    for attempt in range(3):
        try:
            with urllib.request.urlopen(url, timeout=90) as response:
                return json.load(response)
        except Exception:
            if attempt == 2: raise
            time.sleep(2 ** attempt)

def point_in_ring(point, ring):
    inside = False
    for index, current in enumerate(ring):
        previous = ring[index - 1]
        if (current[1] > point[1]) != (previous[1] > point[1]):
            x = (previous[0] - current[0]) * (point[1] - current[1]) / (previous[1] - current[1]) + current[0]
            if point[0] < x: inside = not inside
    return inside

def representative_point(geometry):
    ring = rings(geometry)[0]
    return [sum(point[0] for point in ring) / len(ring), sum(point[1] for point in ring) / len(ring)]

def load_admin1_regions():
    parents = {}
    with gzip.open(ADMIN1_ASSET, "rt", encoding="utf-8") as stream:
        asset = json.load(stream)
    for region in asset["Regions"]:
        polygons = [Polygon([(point["X"], point["Y"]) for point in part["Points"]]) for part in region["Parts"] if len(part["Points"]) >= 3]
        if polygons:
            parent_geometry = MultiPolygon(polygons)
            parents.setdefault(region["Group"], []).append((region["Id"], prep(parent_geometry), parent_geometry.bounds))
    return parents

def find_parent(geometry, candidates):
    if not candidates: return None
    point = Point(project(representative_point(geometry)))
    for region_id, parent_geometry, bounds in candidates:
        if bounds[0] <= point.x <= bounds[2] and bounds[1] <= point.y <= bounds[3] and parent_geometry.covers(point): return region_id
    return None

def main():
    catalog = download_json(CATALOG_URL)
    admin1_regions = load_admin1_regions()
    regions, failures = [], []
    for index, dataset in enumerate(catalog, 1):
        country_id = dataset["boundaryISO"]
        try:
            data = download_json(dataset["simplifiedGeometryGeoJSON"])
            for feature in data["features"]:
                properties = feature["properties"]
                name = properties.get("shapeName") or properties.get("name") or properties.get("shapeID")
                parent_group = find_parent(feature["geometry"], admin1_regions.get(country_id))
                parts = []
                for ring in rings(feature["geometry"]):
                    projected = [project(point) for point in ring]
                    points = simplify_for_powerpoint(projected)
                    if len(points) < 4:
                        min_x, max_x = min(point[0] for point in projected), max(point[0] for point in projected)
                        min_y, max_y = min(point[1] for point in projected), max(point[1] for point in projected)
                        size = 0.0001
                        if max_x - min_x < size: min_x, max_x = min_x - size / 2, max_x + size / 2
                        if max_y - min_y < size: min_y, max_y = min_y - size / 2, max_y + size / 2
                        points = [[min_x, min_y], [max_x, min_y], [max_x, max_y], [min_x, max_y], [min_x, min_y]]
                    parts.append({"Points": [{"X": x, "Y": y} for x, y in points[:-1]]})
                if not parts:
                    continue
                shape_id = str(properties.get("shapeID") or len(regions))
                regions.append({
                    "Id": f"{country_id}-ADM2-{shape_id}", "Name": name, "EnglishName": name,
                    "Group": country_id, "ParentGroup": parent_group, "EntityType": dataset.get("boundaryCanonical") or "Second-level administrative region",
                    "Aliases": [value for value in (name, shape_id) if value], "Parts": parts,
                })
        except Exception as error:
            failures.append(f"{country_id}: {error}")
        print(f"{index}/{len(catalog)} {country_id}: {len(regions)} regions total")

    if failures:
        raise RuntimeError("Failed ADM2 downloads:\n" + "\n".join(failures))
    asset = {
        "MapId": "ADMIN2-GLOBAL", "DisplayName": "Second-level Administrative Regions",
        "Source": "geoBoundaries gbOpen", "SourceVersion": "current", "SourceUrl": "https://www.geoboundaries.org/",
        "AspectRatio": 2.0, "CoordinateSpace": "GlobalEquirectangular", "MinLatitude": -90.0, "MaxLatitude": 90.0,
        "Regions": sorted(regions, key=lambda region: (region["Group"], region["Name"])),
    }
    OUTPUT.write_bytes(gzip.compress(json.dumps(asset, ensure_ascii=False, separators=(",", ":")).encode("utf-8"), compresslevel=9, mtime=0))
    print(f"{OUTPUT.name}: {len(regions)} regions across {len(catalog)} countries")

if __name__ == "__main__": main()
