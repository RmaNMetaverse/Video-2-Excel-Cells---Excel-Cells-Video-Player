import cv2
import win32com.client
import pywintypes
import time
import os

def get_user_inputs():
    print("--- 📑 Excel High-Speed Flipbook Generator (Vibrant Edition) ---")
    while True:
        video_path = input("Enter video file path (or drag-and-drop): ").strip().strip('"').strip("'")
        if os.path.exists(video_path): break
        print("❌ Error: File not found.")
        
    while True:
        try:
            width = int(input("Enter Excel resolution width (Recommended 40-60): "))
            if width > 0: break
        except ValueError: pass
            
    while True:
        try:
            max_frames = int(input("Enter max frames to extract (Recommended 30-100): "))
            if max_frames > 0: break
        except ValueError: pass
            
    return video_path, width, max_frames

def get_column_letter(n):
    result = ""
    while n > 0:
        n -= 1
        result = chr(n % 26 + 65) + result
        n //= 26
    return result

VIDEO_FILE, WIDTH, MAX_FRAMES = get_user_inputs()

print("\nStarting Excel...")
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = True 
excel.Interactive = False 

try:
    wb = excel.Workbooks.Add()
    ws = wb.ActiveSheet
    ws.Name = "Frame_1"

    print("Formatting grid...")
    ws.Columns("A:ZZ").ColumnWidth = 2.0 
    ws.Rows("1:200").RowHeight = 15.0
    excel.ActiveWindow.Zoom = 30

    cap = cv2.VideoCapture(VIDEO_FILE)
    success, first_frame = cap.read()
    if not success:
        print("Error: Could not read video.")
        exit()

    height, width_vid, _ = first_frame.shape
    ratio = WIDTH / width_vid
    new_height = int(height * ratio)

    print("Pre-calculating cell matrices...")
    cell_names = []
    for y in range(new_height):
        row = []
        for x in range(WIDTH):
            row.append(f"{get_column_letter(x+1)}{y+1}")
        cell_names.append(row)

    cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
    print(f"▶️ Extracting {MAX_FRAMES} frames to separate tabs...")

    prev_grid = [[None for _ in range(WIDTH)] for _ in range(new_height)]
    start_time = time.time()
    frames_processed = 0

    excel.ScreenUpdating = False 

    for frame_num in range(MAX_FRAMES):
        success, frame = cap.read()
        if not success: break 
            
        if frame_num > 0:
            last_sheet = wb.Sheets(wb.Sheets.Count)
            ws.Copy(None, last_sheet) 
            ws = wb.Sheets(wb.Sheets.Count)
            ws.Name = f"Frame_{frame_num + 1}"

        resized_frame = cv2.resize(frame, (WIDTH, new_height))
        
        # --- THE COLOR UPGRADE ---
        # 1. Boost contrast (alpha) and brightness (beta) so the colors "pop"
        enhanced_frame = cv2.convertScaleAbs(resized_frame, alpha=1.35, beta=10)
        
        # 2. Use a gentler quantization (24 instead of 32) for better color depth
        quantized = (enhanced_frame // 24) * 24
        
        color_map = {}
        
        for y in range(new_height):
            for x in range(WIDTH):
                b, g, r = quantized[y, x]
                color_int = int(r) + (int(g) * 256) + (int(b) * 65536)
                
                if prev_grid[y][x] != color_int:
                    prev_grid[y][x] = color_int
                    if color_int not in color_map:
                        color_map[color_int] = []
                    color_map[color_int].append(cell_names[y][x])
        
        for color_int, cells in color_map.items():
            for i in range(0, len(cells), 40):
                range_str = ",".join(cells[i:i+40])
                try:
                    ws.Range(range_str).Interior.Color = color_int
                except pywintypes.com_error:
                    pass 
        
        frames_processed += 1
        print(f"Tab created: Frame_{frames_processed}/{MAX_FRAMES}")

    save_path = os.path.join(os.getcwd(), "Excel_Flipbook.xlsx")
    
    excel.DisplayAlerts = False 
    wb.SaveAs(save_path)
    excel.DisplayAlerts = True

    excel.ScreenUpdating = True 
    total_time = time.time() - start_time
    
    print(f"\n✅ Finished! Processed {frames_processed} frames in {total_time:.1f} seconds.")
    print(f"💾 File saved successfully to: {save_path}")

except KeyboardInterrupt:
    print("\n⏹️ Stopped by user.")
except Exception as e:
    print(f"\n❌ An error occurred: {e}")
finally:
    try:
        excel.Interactive = True
        excel.ScreenUpdating = True
    except: pass