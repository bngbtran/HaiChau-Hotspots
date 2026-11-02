"""Cập nhật dữ liệu cho web bằng một lệnh:

  1. Chạy lại notebook HaiChau_Hotspots.ipynb: đọc data/cleaned/data_cleaned.xlsx -> data/web/areas.json
  2. In tóm tắt kết quả; tải lại trang web (Live Server) hoặc deploy lại để thấy dữ liệu mới

Chạy bằng Python của môi trường conda msa:
  D:\\LIBRARY\\MINICONDA\\Conda\\envs\\msa\\python.exe scripts/update_web.py
"""
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def run(*cmd):
    print("$", " ".join(str(c) for c in cmd), flush=True)
    subprocess.run(cmd, check=True, cwd=ROOT)


def main():
    run(sys.executable, "-m", "jupyter", "nbconvert", "--to", "notebook", "--execute", "--inplace",
        "--ExecutePreprocessor.kernel_name=msa", "--ExecutePreprocessor.timeout=900", "HaiChau_Hotspots.ipynb")

    d = json.loads((ROOT / "data" / "web" / "areas.json").read_text(encoding="utf-8"))
    m = d["meta"]
    print(f"\nĐã cập nhật data/web/areas.json lúc {m['generated_at']}")
    print(f"  Nguồn: {m['source']['stores']}")
    print(f"  {m['method']} k={m['k']} · silhouette {m['metrics']['silhouette']} · ổn định {m['metrics']['stability']}")
    for t in d["types"]:
        print(f"  - {t['name']}: {t['cells']} ô, {t['stores']} quán")
    print("\nTải lại trang web (Live Server) hoặc deploy lại thư mục dự án lên Vercel.")


if __name__ == "__main__":
    main()
