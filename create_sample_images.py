"""
샘플 테스트 이미지를 생성하는 스크립트
"""
import os
from PIL import Image, ImageDraw, ImageFont


def create_sample_image(filepath: str, lines: list[str], bg_color="white", text_color="black"):
    """텍스트가 담긴 샘플 이미지를 생성합니다."""
    width, height = 800, 400
    img = Image.new("RGB", (width, height), color=bg_color)
    draw = ImageDraw.Draw(img)

    # 기본 폰트 사용
    try:
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 24)
        font_title = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 28)
    except IOError:
        font = ImageFont.load_default()
        font_title = font

    y = 30
    for i, line in enumerate(lines):
        f = font_title if i == 0 else font
        draw.text((40, y), line, fill=text_color, font=f)
        y += 45

    img.save(filepath)
    print(f"생성됨: {filepath}")


def main():
    os.makedirs("images", exist_ok=True)

    samples = [
        {
            "filename": "images/sample_01_english.png",
            "lines": [
                "OCR Text Extraction Test",
                "Hello, this is a sample image.",
                "The quick brown fox jumps over the lazy dog.",
                "Image processing with Python and Tesseract.",
                "Date: 2026-03-04",
            ],
        },
        {
            "filename": "images/sample_02_invoice.png",
            "lines": [
                "INVOICE #INV-2026-001",
                "Company: ACME Corporation",
                "Product: Widget A x 10 = $250.00",
                "Product: Widget B x  5 = $150.00",
                "Total Amount: $400.00",
                "Due Date: 2026-04-01",
            ],
        },
        {
            "filename": "images/sample_03_report.png",
            "lines": [
                "Monthly Sales Report",
                "Region: Asia Pacific",
                "Q1 Target: 10,000 units",
                "Q1 Actual: 12,340 units",
                "Achievement: 123.4%",
                "Status: EXCEEDED TARGET",
            ],
            "bg_color": "#F0F8FF",
        },
    ]

    for s in samples:
        create_sample_image(
            s["filename"],
            s["lines"],
            bg_color=s.get("bg_color", "white"),
        )

    print("\n샘플 이미지 생성 완료!")
    print("다음 명령어로 OCR을 실행하세요:")
    print("  python image_to_excel.py ./images")


if __name__ == "__main__":
    main()
