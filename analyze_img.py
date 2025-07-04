import argparse
import numpy as np
import matplotlib.pyplot as plt

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
    parser.add_argument('img_path', help='Path to .img file')
    parser.add_argument('--show', action='store_true', help='Display image')
    parser.add_argument('--output', help='Save plotted image to file')
    args = parser.parse_args()

    arr = load_raw_image(args.img_path)
    print(f'Shape: {arr.shape}, dtype: {arr.dtype}')
    print(f'Min: {arr.min()}, Max: {arr.max()}, Mean: {arr.mean():.2f}')

    if args.show or args.output:
        plt.imshow(arr, cmap='gray', vmin=arr.min(), vmax=arr.max())
        plt.colorbar(label='Intensity')
        plt.title(args.img_path)
        if args.output:
            plt.savefig(args.output, dpi=300)
        if args.show:
            plt.show()

if __name__ == '__main__':
    main()
