# Excel Video Projector (Video → .bin for Excel)

https://github.com/user-attachments/assets/487292d6-bb29-4877-801e-8b208686890f



A small Python tool that converts a video (.mp4) into a compact binary (.bin) frame stream which an Excel VBA macro can rapidly paint into worksheet cells to "play" the video inside Excel.

## How it works
- The Python script `excelvideoprojector.py` reads a video, resizes each frame to a user-specified resolution (width × height in cells), quantizes colors to reduce palette size, and writes raw RGB bytes per frame into a .bin file.
- The provided VBA macros in `ExcelVBAScript.VBA` load the .bin file and paint each pixel into a cell via `Interior.Color = RGB(r,g,b)`.

## Requirements
- Python 3.8+ (recommended)
- Excel (Windows) with VBA support

Python libraries (install with pip):

```powershell
python -m pip install --upgrade pip
pip install opencv-python numpy
```

## Generate the .bin file
1. Run the Python script:

```powershell
py .\excelvideoprojector.py
```

2. When prompted, drag-and-drop or paste the path to your `.mp4` file.
3. Enter the desired `Width in cells` and `Height in cells` (press Enter to use defaults). The script will produce a file named `<video_basename>_optimized.bin` next to your video.

Notes:
- The script defaults are `width=100` and `height=75` if you press Enter. It resizes frames and writes raw RGB bytes (3 bytes per pixel) per frame.
- The code rounds RGB values to multiples of 16 to keep the color palette small and reduce visual variance in Excel (`frame = (frame // 16) * 16`).

## Play the .bin file in Excel
1. Open the VBA module (ALT + F11, from the menus above click on Insert > module), pase the modified code in`ExcelVBAScript.VBA`, and then close the window.

## Important: Resolution must match
- The Python script resizes video frames to `(width × height)` and writes exactly `width * height * 3` bytes per frame. The VBA macro reads that many bytes per frame and populates a `width × height` region in the worksheet.
- If these numbers don't match, frames will be misread or the macro will fail. Always set the same width/height in both the Python input and the VBA variables.

## Recommended settings & tips
- Start with `100×75` or `100×100` for a balance between visual quality and Excel performance.
- Smaller resolutions run much faster in Excel; larger ones will be slow and may freeze Excel.
- If Excel errors with `Error 1004` when painting colors, ensure the `.bin` file was created by this script (RGB, 3 bytes per pixel). The Python script includes color quantization to avoid palette explosion.

2. ALT + F8 to run the Macros you just added.
3. Run `SetupCanvasMinimal` to prepare the worksheet (it sizes columns/rows and paints the canvas black).
4. Run `PlayVideoStreamMaxSpeed` and open the `.bin` file created by the Python script. The macro reads frames and paints cells rapidly.

File references:
- Python script: [excelvideoprojector.py](excelvideoprojector.py)
- VBA macros: [ExcelVBAScript.VBA](ExcelVBAScript.VBA)



## Troubleshooting
- If the Python script reports "File not found", ensure the path you paste is correct and accessible from your user account.
- If Excel is very slow, reduce the resolution used when generating the .bin file.

## License & Notes
This is a small utility intended for experimentation and creative Excel demos. Feel free to adapt the code for your needs.
