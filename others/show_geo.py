import requests

def reverse_geocode(lng, lat):
    url = f"https://restapi.amap.com/v3/geocode/regeo?key=YOUR_KEY&location={lng},{lat}"
    return requests.get(url).json()

# 示例调用（取多边形中心点）
print(reverse_geocode(116.75, 39.92))  # 应返回"北京市丰台区"
