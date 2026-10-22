import os
import django
from django.test import Client

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "ppttoxls_web.settings")
django.setup()

def test_endpoints():
    client = Client()

    print("Testing GET / ...")
    r1 = client.get("/")
    assert r1.status_code == 200, f"Expected 200, got {r1.status_code}"
    print("GET / passed!")

    print("Testing GET /load-demo/ ...")
    r2 = client.get("/load-demo/")
    assert r2.status_code == 200, f"Expected 200, got {r2.status_code}"
    data2 = r2.json()
    assert "mapping_json" in data2, "mapping_json missing from response"
    print("GET /load-demo/ passed!")

    print("Testing POST /convert/ with demo mode standard ...")
    r3 = client.post("/convert/", {"use_demo": "standard"})
    assert r3.status_code == 200, f"Expected 200, got {r3.status_code}: {r3.content}"
    data3 = r3.json()
    assert data3.get("success") is True, f"Conversion failed: {data3}"
    assert "download_url" in data3, "download_url missing"
    assert "extracted_data" in data3, "extracted_data missing"
    file_id = data3["file_id"]
    print(f"POST /convert/ standard passed! File ID: {file_id}")

    print("Testing POST /convert/ with demo mode variant (determinism test) ...")
    r3_var = client.post("/convert/", {"use_demo": "variant"})
    assert r3_var.status_code == 200, f"Expected 200, got {r3_var.status_code}: {r3_var.content}"
    data3_var = r3_var.json()
    assert data3_var.get("success") is True
    # Verify exact match of extracted numbers
    assert data3_var["extracted_data"]["total_budget"] == data3["extracted_data"]["total_budget"]
    print("POST /convert/ variant passed and matched standard deck output!")

    print(f"Testing GET {data3['download_url']} ...")
    r4 = client.get(data3["download_url"])
    assert r4.status_code == 200, f"Expected 200, got {r4.status_code}"
    content_bytes = b"".join(r4.streaming_content)
    assert len(content_bytes) > 1000, "Excel binary content is too small"
    print(f"GET /download/<file_id>/ passed with valid Excel binary stream ({len(content_bytes)} bytes)!")

    print("\n>>> ALL DJANGO ENDPOINTS TESTED & WORKING PERFECTLY! <<<")

if __name__ == "__main__":
    test_endpoints()
