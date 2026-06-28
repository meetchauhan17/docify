import cv2
import easyocr
import sys
import numpy as np

print("Initializing EasyOCR...")
reader = easyocr.Reader(['en'], gpu=False)

def log_results(title, results):
    print("=" * 50)
    print(f"--- {title} ---")
    if not results:
        print(" -> [NO TEXT DETECTED]")
    for (bbox, text, prob) in results:
        print(f"[{prob:.2f}] {text}")
    print("=" * 50 + "\n")

def test_easyocr(image_path):
    print(f"Loading image: {image_path}")
    img = cv2.imread(image_path)
    if img is None:
        print("Failed to load image!")
        return

    # 1. Test Raw Image
    res_raw = reader.readtext(image_path)
    log_results("RAW IMAGE (No Preprocessing)", res_raw)

    # 2. Test Raw Image with EasyOCR's built-in parameters (mag_ratio, adjust_contrast)
    res_enhanced = reader.readtext(image_path, adjust_contrast=0.5, mag_ratio=1.5, text_threshold=0.5, low_text=0.3)
    log_results("RAW + EASYOCR PARAMS (adjust_contrast=0.5, mag_ratio=1.5)", res_enhanced)

    # 3. Test Current OpenCV Pipeline
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray_blur = cv2.GaussianBlur(gray, (5, 5), 0)
    _, thresh = cv2.threshold(gray_blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    kernel = np.ones((2,2), np.uint8)
    thresh = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel, iterations=1)
    thresh = cv2.resize(thresh, None, fx=2, fy=2)
    
    cv2.imwrite("outputs/test_current_pipeline.png", thresh)
    res_pipe = reader.readtext("outputs/test_current_pipeline.png")
    log_results("CURRENT PIPELINE (saved to outputs/test_current_pipeline.png)", res_pipe)

    # 4. Test Current Pipeline + EasyOCR Params
    res_pipe_enhanced = reader.readtext("outputs/test_current_pipeline.png", adjust_contrast=0.5, mag_ratio=1.5, text_threshold=0.5, low_text=0.3)
    log_results("CURRENT PIPELINE + EASYOCR PARAMS", res_pipe_enhanced)

if __name__ == "__main__":
    test_img = "uploads/006759f7-56d0-483d-b4f7-719bfc6b62d6_Screenshot 2026-02-27 191932.png"
    test_easyocr(test_img)
