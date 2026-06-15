import os
import urllib.request
import zipfile

URL = "https://download.geonames.org/export/zip/US.zip"


def main():
    os.makedirs("data", exist_ok=True)
    txt_path = os.path.join("data", "US.txt")
    if os.path.exists(txt_path):
        print(f"{txt_path} already exists, skipping download")
        return
    zip_path = os.path.join("data", "US.zip")
    urllib.request.urlretrieve(URL, zip_path)
    with zipfile.ZipFile(zip_path) as zf:
        zf.extract("US.txt", "data")
    os.remove(zip_path)
    print(f"saved {txt_path} ({os.path.getsize(txt_path):,} bytes)")


if __name__ == "__main__":
    main()
