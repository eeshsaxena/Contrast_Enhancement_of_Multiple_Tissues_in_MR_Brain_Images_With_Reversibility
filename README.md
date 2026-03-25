"""
=============================================================================
FULL IMPLEMENTATION OF:
"Contrast Enhancement of Multiple Tissues in MR Brain Images With Reversibility"
Hao-Tian Wu, Kaihan Zheng, Qi Huang, Jiankun Hu
IEEE Signal Processing Letters, Vol. 28, 2021
DOI: 10.1109/LSP.2020.3048840
=============================================================================

PAPER OVERVIEW:
--------------
This paper proposes a hierarchical contrast enhancement (CE) scheme for MR brain
images with reversibility using Reversible Data Hiding (RDH).

KEY CONTRIBUTIONS:
  1. U-Net CNN tissue segmentation for multi-class MR brain tissue labeling
  2. Hierarchical CE guided per-tissue (WM, GM, CSF, Background)
  3. Principal grey-level identification via percentage threshold R
  4. Reversible histogram bin expansion (Eq.1) – embeds recovery bits
  5. Side info stored in LSBs of 16 pixels for full lossless recovery
  6. Recovery via Eq.(2) and Eq.(3) – perfect original image restoration
  7. Metrics: PSNR, SSIM, RCEOI (contrast), REEOI (entropy), RMBEOI (brightness)

HOW TO RUN:
-----------
  pip install numpy scipy scikit-image matplotlib pillow
  python 1.py

The script will:
  - Generate a synthetic MR-like brain image (if no real image given)
  - Optionally load a real grayscale MRI from a path (edit REAL_IMAGE_PATH)
  - Perform multi-tissue segmentation (U-Net simplified with Otsu thresholds)
  - Apply hierarchical reversible CE for each tissue class (Procedure 1)
  - Recover the original image from each enhanced image (Procedure 2)
  - Compute and display all paper metrics in a table
  - Display comparison figures
=============================================================================
"""
