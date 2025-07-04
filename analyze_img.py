import argparse
import numpy as np
import matplotlib.pyplot as plt
from tkinter import Tk, filedialog

# Constants based on U8167.exe settings
HEIGHT = 639
WIDTH = 479
DTYPE = np.uint16
HEADER_BYTES = 11084

def load_raw_image(path: str) -> np.ndarray:
    """Load raw .img file and return as (HEIGHT, WIDTH) numpy array."""
    with open(path, 'rb') as f:
        f.seek(HEADER_BYTES)
        data = np.fromfile(f, dtype=DTYPE, count=HEIGHT * WIDTH)
    if data.size != HEIGHT * WIDTH:
        raise ValueError('Unexpected file size')
    return data.reshape((HEIGHT, WIDTH))

def main():
    parser = argparse.ArgumentParser(description='Analyze raw IMG file')
    parser.add_argument('--show', action='store_true', help='Display image')
    parser.add_argument('--output', help='Save plotted image to file')
    args = parser.parse_args()

    # Use GUI dialog to choose the .img file
    root = Tk()
    root.withdraw()  # hide tkinter main window
    img_path = filedialog.askopenfilename(
        title='Select .img file',
        filetypes=[('Raw Image', '*.img'), ('All files', '*.*')]
    )
    if not img_path:
        print('No file selected.')
        return

    arr = load_raw_image(img_path)
    print(f"Loaded: {img_path}")
    print(f'Shape: {arr.shape}, dtype: {arr.dtype}')
    print(f'Min: {arr.min()}, Max: {arr.max()}, Mean: {arr.mean():.2f}')

    if args.show or args.output:
        plt.imshow(arr, cmap='gray', vmin=arr.min(), vmax=arr.max())
        plt.colorbar(label='Intensity')
        plt.title(img_path)
        if args.output:
            plt.savefig(args.output, dpi=300)
        if args.show:
            plt.show()

if __name__ == '__main__':
    main()
