import cv2
from transformers import TrOCRProcessor, VisionEncoderDecoderModel
from PIL import Image
import torch

print("Loading TrOCR model for Handwriting...")
# Using the base handwriting model
processor = TrOCRProcessor.from_pretrained("microsoft/trocr-base-handwritten")
model = VisionEncoderDecoderModel.from_pretrained("microsoft/trocr-base-handwritten")

device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
model.to(device)

def test_trocr(image_path):
    print(f"Loading image: {image_path}")
    image = Image.open(image_path).convert("RGB")
    
    # Preprocess
    pixel_values = processor(image, return_tensors="pt").pixel_values.to(device)
    
    # Generate
    generated_ids = model.generate(pixel_values, max_new_tokens=256)
    generated_text = processor.batch_decode(generated_ids, skip_special_tokens=True)[0]
    
    print("====================================")
    print("TrOCR DETECTED TEXT:")
    print(generated_text)
    print("====================================")

if __name__ == "__main__":
    test_img = "uploads/fb6830c9-bec4-4a37-8dfd-afaaf5a0e540_Screenshot 2026-02-27 191932.png"
    test_trocr(test_img)
