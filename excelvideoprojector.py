import cv2
import numpy as np
import os

print("--- Excel Video Projector: Optimized ---")
video_path = input("Drag and drop your .mp4 file here: ").strip().strip('"').strip("'")

if not os.path.exists(video_path):
    print("Error: File not found.")
    exit()

try:
    width = int(input("Width in cells (Press Enter for 100): ") or 100)
    height = int(input("Height in cells (Press Enter for 75): ") or 75)
except ValueError:
    width, height = 100, 75

directory = os.path.dirname(video_path)
filename = os.path.splitext(os.path.basename(video_path))[0]
output_bin = os.path.join(directory, f"{filename}_optimized.bin")

cap = cv2.VideoCapture(video_path)
frame_count = 0

print(f"\nProcessing {width}x{height} video...")

with open(output_bin, 'wb') as f:
    while True:
        ret, frame = cap.read()
        if not ret: break
        
        # Resize and convert to RGB
        frame = cv2.resize(frame, (width, height))
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        
        # CRUCIAL FIX FOR ERROR 1004: Quantize colors
        # Rounds all RGB values to multiples of 16 to keep the color palette small
        frame = (frame // 16) * 16
        
        # Write the entire frame array instantly
        f.write(frame.tobytes())
        
        frame_count += 1
        if frame_count % 100 == 0:
            print(f"Processed {frame_count} frames...")

cap.release()
print(f"Done! Saved to: {output_bin}")