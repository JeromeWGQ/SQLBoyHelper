import sys
import json
import math

def haversine(lon1, lat1, lon2, lat2):
    # Convert decimal degrees to radians
    lon1, lat1, lon2, lat2 = map(math.radians, [lon1, lat1, lon2, lat2])
    # Haversine formula
    dlon = lon2 - lon1
    dlat = lat2 - lat1
    a = math.sin(dlat/2)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon/2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))
    r = 6371000  # Radius of earth in meters
    return c * r

def interpolate_points(lon1, lat1, lon2, lat2, distance):
    num_segments = int(distance // 5)
    points = [(lon1, lat1)]
    for i in range(1, num_segments):
        fraction = i / num_segments
        lon = lon1 + fraction * (lon2 - lon1)
        lat = lat1 + fraction * (lat2 - lat1)
        points.append((lon, lat))
    points.append((lon2, lat2))
    return points

def process_coordinates(coordinates):
    # Remove the third dimension from each coordinate
    coordinates_2d = [(lon, lat) for lon, lat, _ in coordinates]
    results = []
    for i in range(len(coordinates_2d) - 1):
        lon1, lat1 = coordinates_2d[i]
        lon2, lat2 = coordinates_2d[i + 1]
        distance = haversine(lon1, lat1, lon2, lat2)
        results.extend(interpolate_points(lon1, lat1, lon2, lat2, distance))
    return results

if __name__ == "__main__":
    for line in sys.stdin:
        coord_str = line.strip()
        coordinates = json.loads(coord_str)
        result = process_coordinates(coordinates)
        print(json.dumps(result))
