import requests, json

API_KEY = "comfyui-dfca9467ae46c03990902717582983c6cc8d5be7e0ccea4fac6ce45051393c8e"
url = "https://cloud.comfy.org/api/object_info"
r = requests.get(url, headers={"X-API-Key": API_KEY})
r.raise_for_status()
data = r.json()

node_names = sorted(list(data.keys()))
print("\n".join(node_names))

# 필요하면 파일 저장
with open("comfy_nodes.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(node_names))