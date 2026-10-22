"""Attach first-level parent IDs to the generated global ADM2 and ADM3 assets."""

import gzip
import json
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent
ADMIN1 = ROOT / "states_provinces_admin1_naturalearth_5_1_1.json.gz"
TARGETS = [
    ROOT / "administrative_regions_adm2_geoboundaries_current.json.gz",
    ROOT / "administrative_regions_adm3_geoboundaries_current.json.gz",
]

def load(path):
    with gzip.open(path, "rt", encoding="utf-8") as stream: return json.load(stream)

def bounds(parts):
    points = [point for part in parts for point in part["Points"]]
    return min(point["X"] for point in points), min(point["Y"] for point in points), max(point["X"] for point in points), max(point["Y"] for point in points)

def center(parts):
    left, top, right, bottom = bounds(parts)
    return (left + right) / 2, (top + bottom) / 2

def main():
    admin1 = load(ADMIN1)
    parents = {}
    for region in admin1["Regions"]:
        left, top, right, bottom = bounds(region["Parts"])
        parents.setdefault(region["Group"], []).append((region["Id"], left, top, right, bottom))

    for target in TARGETS:
        asset, assigned = load(target), 0
        for region in asset["Regions"]:
            x, y = center(region["Parts"])
            candidates = [parent for parent in parents.get(region["Group"], []) if parent[1] <= x <= parent[3] and parent[2] <= y <= parent[4]]
            if candidates:
                region["ParentGroup"] = min(candidates, key=lambda parent: (parent[3] - parent[1]) * (parent[4] - parent[2]))[0]
                assigned += 1
            else:
                region["ParentGroup"] = None
        target.write_bytes(gzip.compress(json.dumps(asset, ensure_ascii=False, separators=(",", ":")).encode("utf-8"), compresslevel=9, mtime=0))
        print(f"{target.name}: {assigned}/{len(asset['Regions'])} regions assigned")

if __name__ == "__main__": main()
