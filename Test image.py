from rembg import remove
from PIL import Image
import numpy as np
import skimage.color
import cv2

def get_mask_rgb(image_path, out_shape=None):
    img = Image.open(image_path).convert('RGBA')
    if out_shape is not None:
        img = img.resize(out_shape, resample=Image.BILINEAR)
    arr = np.array(img)
    out = remove(img)
    out = out.resize(img.size, resample=Image.BILINEAR)
    out_arr = np.array(out)
    mask = out_arr[:,:,3] > 0
    rgb = arr[:,:,:3]
    return rgb, mask

def compare_products_ignore_light(
    file1, file2,
    label1="Sản phẩm A", label2="Sản phẩm B",
    deltae_threshold=4
):
    img1 = Image.open(file1)
    size = img1.size
    rgb1, mask1 = get_mask_rgb(file1, out_shape=size)
    rgb2, mask2 = get_mask_rgb(file2, out_shape=size)
    min_h = min(rgb1.shape[0], rgb2.shape[0])
    min_w = min(rgb1.shape[1], rgb2.shape[1])
    rgb1, mask1 = rgb1[:min_h, :min_w], mask1[:min_h, :min_w]
    rgb2, mask2 = rgb2[:min_h, :min_w], mask2[:min_h, :min_w]
    mask_core = cv2.erode((mask1 & mask2).astype(np.uint8), np.ones((7,7), np.uint8), iterations=1).astype(bool)
    pixels1 = rgb1[mask_core]
    pixels2 = rgb2[mask_core]
    mean1 = np.mean(pixels1, axis=0)
    mean2 = np.mean(pixels2, axis=0)
    # Chuyển sang Lab
    lab1 = skimage.color.rgb2lab(mean1[np.newaxis, np.newaxis, :] / 255.0)[0,0]
    lab2 = skimage.color.rgb2lab(mean2[np.newaxis, np.newaxis, :] / 255.0)[0,0]
    # Chỉ lấy a* và b* (bỏ qua L*)
    ab1 = lab1[1:3]
    ab2 = lab2[1:3]
    # Tính khoảng cách màu (ab space)
    ab_dist = np.linalg.norm(ab1 - ab2)
    print(f"{label1} ab*: {ab1}")
    print(f"{label2} ab*: {ab2}")
    print(f"Khoảng cách màu tổng thể (bỏ qua ánh sáng): {ab_dist:.2f}")
    # Đánh giá
    if ab_dist < deltae_threshold:
        comment = f"{label1} và {label2} được coi là CÙNG MÀU tổng thể nếu bỏ qua tác động ánh sáng (|Δab*| = {ab_dist:.2f} < {deltae_threshold})"
    else:
        comment = f"{label1} và {label2} CÓ SỰ KHÁC BIỆT MÀU tổng thể (|Δab*| = {ab_dist:.2f} ≥ {deltae_threshold})"
    print("=> Kết luận:", comment)
    return comment, ab_dist, ab1, ab2

# Ví dụ sử dụng
file1 = "IMG_7182.jpg"
file2 = "IMG_7183.jpg"
label1 = "Mẫu 1"
label2 = "Mẫu 2"

comment, ab_dist, ab1, ab2 = compare_products_ignore_light(file1, file2, label1, label2)
