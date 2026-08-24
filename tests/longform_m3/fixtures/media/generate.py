"""Rebuild the deterministic tiny media fixtures used by M3 tests."""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageCms


ROOT = Path(__file__).parent


def _rgb(size: tuple[int, int] = (4, 3)) -> Image.Image:
    image = Image.new("RGB", size)
    image.putdata(
        [
            ((x * 61) % 256, (y * 83) % 256, ((x + y) * 47) % 256)
            for y in range(size[1])
            for x in range(size[0])
        ]
    )
    return image


def main() -> None:
    base = _rgb()
    base.save(ROOT / "png-wrong.jpg", format="PNG", dpi=(300, 300))
    base.save(ROOT / "jpeg-wrong.png", format="JPEG", quality=95, subsampling=0)
    base.save(ROOT / "tiff-wrong.dat", format="TIFF", compression="raw")
    base.save(ROOT / "bmp-wrong.gif", format="BMP")

    second = Image.new("RGB", base.size, (240, 30, 80))
    base.save(ROOT / "gif-wrong.bmp", format="GIF")
    base.save(
        ROOT / "animated.gif",
        format="GIF",
        save_all=True,
        append_images=[second],
        duration=(20, 20),
        loop=0,
    )
    base.save(
        ROOT / "multipage.tiff",
        format="TIFF",
        save_all=True,
        append_images=[second],
        compression="raw",
    )

    icc = ImageCms.ImageCmsProfile(ImageCms.createProfile("sRGB")).tobytes()
    oriented = _rgb((2, 3))
    exif = Image.Exif()
    exif[274] = 6
    oriented.save(
        ROOT / "oriented.jpg",
        format="JPEG",
        quality=100,
        subsampling=0,
        exif=exif,
        icc_profile=icc,
    )

    base.save(ROOT / "webp-renamed.png", format="WEBP", lossless=True)
    (ROOT / "static.svg").write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20" '
        'viewBox="0 0 40 20"><rect width="40" height="20" fill="#369"/></svg>',
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
