import os
import sys
import tempfile
from pathlib import Path
from unittest.mock import patch

from PIL import Image, ImageDraw, ImageFont

from web_app import parse_screenshot_tsv, read_screenshot_metrics


def check_packaged_ocr(backend):
    internal = Path(backend).resolve() / "_internal"
    engine = internal / "ocr" / "tesseract"
    tessdata = internal / "ocr" / "tessdata"
    assert engine.is_file()
    assert all((tessdata / name).is_file() for name in ("chi_sim.traineddata", "eng.traineddata"))
    env = {**os.environ, "PATH": "/usr/bin:/bin", "TESSDATA_PREFIX": str(tessdata)}

    def check_positions(text):
        records = parse_screenshot_tsv(text)
        assert "12345" in text, "Bundled OCR did not read the synthetic test image"
        return records

    with tempfile.TemporaryDirectory() as directory:
        image = Image.new("RGB", (2000, 400), "white")
        ImageDraw.Draw(image).text((80, 80), "12345", font=ImageFont.load_default(size=80), fill="black")
        path = Path(directory) / "ocr.png"
        image.save(path)
        with patch("web_app.tesseract_path", return_value=str(engine)), patch("web_app.tesseract_env", return_value=env), patch("web_app.parse_screenshot_tsv", side_effect=check_positions) as parse:
            read_screenshot_metrics(path)
            parse.assert_called_once()
    print("PASS: bundled OCR returns coordinates without external config files or Homebrew")


if __name__ == "__main__":
    check_packaged_ocr(sys.argv[1])
