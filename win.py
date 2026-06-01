import cv2
import numpy as np
import json
import zstandard as zstd
import base64
from pathlib import Path

OUTPUT_FILENAME = 'a.txt'

SUPPORTED_IMAGE_SUFFIXES = {'.png', '.jpg', '.jpeg', '.webp'}
def find_single_image_file_in_cwd():
    image_files = [path for path in Path.cwd().iterdir() if path.is_file() and path.suffix.lower() in SUPPORTED_IMAGE_SUFFIXES]
    return image_files[0] if len(image_files) == 1 else None

def process_image_to_compressed_txt(image_path):
    img_array = np.fromfile(image_path, dtype=np.uint8)
    img_bgr = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
    
    if img_bgr is None:
        print(f"エラー: 画像ファイル '{image_path}' が見つかりません。")
        return

    h_orig, w_orig, _ = img_bgr.shape
    n_w = w_orig
    n_h = h_orig

    print(f"画像サイズ: {n_w}x{n_h}")
    
    print("reshaping image data...")
    pixels = img_bgr.reshape(-1, 3)

    print("finding unique colors...")
    # RGBを一つの24bit整数にパックする (R << 16 + G << 8 + B)
    # 符号なし32bit整数(uint32)に変換してからシフトします
    pixels_u32 = pixels.astype(np.uint32)
    packed_pixels = (pixels_u32[:, 2] << 16) | (pixels_u32[:, 1] << 8) | pixels_u32[:, 0] # ここでBGRからRGBに変換しつつパック
    # 1次元配列として unique を実行（圧倒的に高速）
    unique_packed, inverse_indices = np.unique(packed_pixels, return_inverse=True)
    # 元のRGB形式（3列）に戻す
    unique_colors = np.zeros((len(unique_packed), 3), dtype=np.uint8)
    unique_colors[:, 0] = (unique_packed >> 16) & 0xFF
    unique_colors[:, 1] = (unique_packed >> 8) & 0xFF
    unique_colors[:, 2] = unique_packed & 0xFF

    print("making list...")
    list_data = [{"id": int(i), "r": int(c[0]), "g": int(c[1]), "b": int(c[2])} for i, c in enumerate(unique_colors)]
    img_data = inverse_indices.reshape(n_h, n_w).tolist()

    print("making json...")    
    combined_data = {"list": list_data, "img": img_data}
    json_str = json.dumps(combined_data, ensure_ascii=False)
    
    print("compressing...")
    params = zstd.ZstdCompressionParameters.from_level(6, threads=-1)
    cctx = zstd.ZstdCompressor(compression_params=params)
    compressed_bytes = cctx.compress(json_str.encode('utf-8'))

    print("encoding to Base64...")
    compressed_base64 = base64.b64encode(compressed_bytes).decode('utf-8')
    
    print("writing to file...")
    with open(OUTPUT_FILENAME, 'w', encoding='utf-8') as f:
        f.write(compressed_base64)
    
    print(f"\n[処理成功]")
    print(f"使われている色の数: {len(list_data)} 色")
    print(f"生成されたファイル: {OUTPUT_FILENAME}")
    print("このファイルをGoogleドライブの mZXjn09dp2zurWN49atK フォルダにアップロードしてください。")

if __name__ == '__main__':
    auto_image = find_single_image_file_in_cwd()
    if auto_image is not None:
        image_file = str(auto_image)
        print(f"画像ファイルが1つだけ見つかったため、自動で読み込みます: {auto_image.name}")
    else:
        image_file = input("画像ファイル名（拡張子含む）を入力してください: ")
    try:
        process_image_to_compressed_txt(image_file)
    except ValueError:
        print("エラー: 正しい値を入力してください。")
