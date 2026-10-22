"""Build a PowerPoint-safe Germany ADM3 asset and assign every district to a state."""

import gzip
import json
import pathlib
import urllib.request

from generate_germany_adm2_asset import project, rings, simplify_for_powerpoint

OUTPUT = pathlib.Path(__file__).resolve().parent / "de_adm3_districts_geoboundaries_2021.json.gz"
SOURCE = "https://github.com/wmgeolab/geoBoundaries/raw/9469f09/releaseData/gbOpen/DEU/ADM3/geoBoundaries-DEU-ADM3.geojson"
ADMIN1 = "https://raw.githubusercontent.com/nvkelso/natural-earth-vector/v5.1.1/geojson/ne_10m_admin_1_states_provinces.geojson"
FALLBACK_PARENTS = {
    "Konstanz": "DE-BW", "Hamburg": "DE-HH", "Rostock, Kreisfreie Stadt": "DE-MV",
    "Nordfriesland": "DE-SH", "Bremen, Kreisfreie Stadt": "DE-HB", "Bremerhaven, Kreisfreie Stadt": "DE-HB",
    "Landkreis Rostock": "DE-MV", "Vorpommern-Rügen": "DE-MV", "Vorpommern-Greifswald": "DE-MV",
    "Stade": "DE-NI", "Emden, Kreisfreie Stadt": "DE-NI", "Wilhelmshaven, Kreisfreie Stadt": "DE-NI",
    "Aurich": "DE-NI", "Leer": "DE-NI", "Wittmund": "DE-NI", "Flensburg, Kreisfreie Stadt": "DE-SH",
    "Kiel, Kreisfreie Stadt": "DE-SH", "Dithmarschen": "DE-SH", "Steinburg": "DE-SH",
}

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

def state_rings(data):
    states = []
    for feature in data["features"]:
        props = feature["properties"]
        if props.get("adm0_a3") != "DEU": continue
        state_id = props.get("iso_3166_2")
        if state_id and state_id != "-99": states.append((state_id, rings(feature["geometry"])))
    return states

def parent_state(point, states):
    for state_id, state_parts in states:
        if any(point_in_ring(point, ring) for ring in state_parts): return state_id
    return None

def main():
    with urllib.request.urlopen(SOURCE) as response: districts = json.load(response)
    with urllib.request.urlopen(ADMIN1) as response: states = state_rings(json.load(response))
    regions, missing = [], []
    for feature in districts["features"]:
        props, name = feature["properties"], feature["properties"]["shapeName"]
        parent = parent_state(representative_point(feature["geometry"]), states) or FALLBACK_PARENTS.get(name)
        if not parent:
            missing.append(name)
            continue
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
        if parts:
            regions.append({"Id": "DE-KRS-" + props["shapeID"], "Name": name, "EnglishName": name,
                            "Group": "DEU", "ParentGroup": parent, "EntityType": "Independent city or district",
                            "Aliases": [name, props["shapeID"]], "Parts": parts})
    if missing: raise RuntimeError("Could not assign districts: " + ", ".join(missing))
    asset = {"MapId": "DEU-ADM3", "DisplayName": "Germany – Districts and Independent Cities",
             "Source": "geoBoundaries / Federal Agency for Cartography and Geodesy", "SourceVersion": "2021",
             "SourceUrl": "https://www.geoboundaries.org/api/current/gbOpen/DEU/ADM3/", "AspectRatio": 2.0,
             "CoordinateSpace": "GlobalEquirectangular", "MinLatitude": -90.0, "MaxLatitude": 90.0,
             "Regions": sorted(regions, key=lambda region: (region["ParentGroup"], region["Name"]))}
    OUTPUT.write_bytes(gzip.compress(json.dumps(asset, ensure_ascii=False, separators=(",", ":")).encode("utf-8"), compresslevel=9, mtime=0))
    print(f"{OUTPUT.name}: {len(regions)} districts across {len(set(region['ParentGroup'] for region in regions))} states")

if __name__ == "__main__": main()
