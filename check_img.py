import os
from PIL import Image
import numpy as np

desktop = 'C:/Users/lenovo/Desktop/'
files = [f for f in os.listdir(desktop) if f.endswith('.jpg')]
print("All jpg files:", files)

# Find the result file
result_file = None
for f in files:
    if '结果' in f or 'result' in f.lower():
        result_file = f
        break

if result_file:
    full = os.path.join(desktop, result_file)
    print(f"Opening: {full}")
    img = Image.open(full)
    print(f"Size: {img.size}")
    print(f"Mode: {img.mode}")

    # Save as png so we can view it
    img.save('result_converted.png')
    print("Saved as result_converted.png")
else:
    print("No result file found")
    # List all files with Chinese chars
    for f in files:
        print(f"  {f!r}")
