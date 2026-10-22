"""Build a PowerPoint-safe Germany ADM2 map from geoBoundaries/BKG data."""

import gzip
import json
import math
import pathlib
import urllib.request

OUTPUT = pathlib.Path(__file__).resolve().parent / "de_adm2_government_districts_geoboundaries_2021.json.gz"
SOURCE = "https://github.com/wmgeolab/geoBoundaries/raw/9469f09/releaseData/gbOpen/DEU/ADM2/geoBoundaries-DEU-ADM2.geojson"
PARENTS = {
    "Stuttgart": "DE-BW", "Karlsruhe": "DE-BW", "Freiburg": "DE-BW", "Tübingen": "DE-BW",
    "Oberbayern": "DE-BY", "Niederbayern": "DE-BY", "Oberpfalz": "DE-BY", "Oberfranken": "DE-BY", "Mittelfranken": "DE-BY", "Unterfranken": "DE-BY", "Schwaben": "DE-BY",
    "Berlin": "DE-BE", "Brandenburg": "DE-BB", "Bremen": "DE-HB", "Hamburg": "DE-HH",
    "Darmstadt": "DE-HE", "Gießen": "DE-HE", "Kassel": "DE-HE", "Mecklenburg-Vorpommern": "DE-MV",
    "Braunschweig": "DE-NI", "Hannover": "DE-NI", "Lüneburg": "DE-NI", "Weser-Ems": "DE-NI",
    "Düsseldorf": "DE-NW", "Köln": "DE-NW", "Münster": "DE-NW", "Detmold": "DE-NW", "Arnsberg": "DE-NW",
    "Koblenz": "DE-RP", "Trier": "DE-RP", "Rheinhessen-Pfalz": "DE-RP", "Saarland": "DE-SL",
    "Dresden": "DE-SN", "Chemnitz": "DE-SN", "Leipzig": "DE-SN", "Sachsen-Anhalt": "DE-ST", "Schleswig-Holstein": "DE-SH", "Thüringen": "DE-TH",
}

def simplify(points, tolerance):
    if len(points) <= 4: return points
    open_points = points[:-1] if points[0] == points[-1] else points
    def distance(point, start, end):
        dx, dy = end[0] - start[0], end[1] - start[1]
        if abs(dx) + abs(dy) < 1e-12: return math.hypot(point[0] - start[0], point[1] - start[1])
        return abs(dy * point[0] - dx * point[1] + end[0] * start[1] - end[1] * start[0]) / math.hypot(dx, dy)
    def rdp(segment):
        if len(segment) <= 2: return segment
        maximum, index = max((distance(point, segment[0], segment[-1]), index) for index, point in enumerate(segment[1:-1], 1))
        return [segment[0], segment[-1]] if maximum <= tolerance else rdp(segment[:index + 1])[:-1] + rdp(segment[index:])
    reduced = rdp(open_points)
    return reduced + [reduced[0]]

def simplify_for_powerpoint(points, tolerance=0.0001, maximum=475):
    reduced, threshold = simplify(points, tolerance), tolerance
    while len(reduced) - 1 > maximum:
        threshold *= 1.5
        reduced = simplify(points, threshold)
    return reduced

def project(point): return [round((point[0] + 180.0) / 180.0, 5), round((90.0 - point[1]) / 180.0, 5)]
def rings(geometry):
    polygons = [geometry["coordinates"]] if geometry["type"] == "Polygon" else geometry["coordinates"]
    return [polygon[0] for polygon in polygons if polygon and len(polygon[0]) >= 4]

def main():
    with urllib.request.urlopen(SOURCE) as response: data = json.load(response)
    regions = []
    for feature in data["features"]:
        properties, name = feature["properties"], feature["properties"]["shapeName"]
        parts = []
        for ring in rings(feature["geometry"]):
            points = simplify_for_powerpoint([project(point) for point in ring])
            if len(points) >= 4: parts.append({"Points": [{"X": x, "Y": y} for x, y in points[:-1]]})
        if not parts: continue
        regions.append({"Id": "DE-RB-" + name.upper().replace(" ", "-"), "Name": name, "EnglishName": name,
                        "Group": "DEU", "ParentGroup": PARENTS[name], "EntityType": "Government district",
                        "Aliases": [name, properties.get("shapeID", "")], "Parts": parts})
    regions.sort(key=lambda region: region["Name"])
    asset = {"MapId": "DEU-ADM2", "DisplayName": "Germany – Government Districts",
             "Source": "geoBoundaries / Federal Agency for Cartography and Geodesy", "SourceVersion": "2021",
             "SourceUrl": "https://www.geoboundaries.org/api/current/gbOpen/DEU/ADM2/", "AspectRatio": 2.0,
             "CoordinateSpace": "GlobalEquirectangular", "MinLatitude": -90.0, "MaxLatitude": 90.0, "Regions": regions}
    OUTPUT.write_bytes(gzip.compress(json.dumps(asset, ensure_ascii=False, separators=(",", ":")).encode("utf-8"), compresslevel=9, mtime=0))
    print(f"{OUTPUT.name}: {len(regions)} government districts")

if __name__ == "__main__": main()
