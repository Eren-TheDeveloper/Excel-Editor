"""Gerçek ürün bilgisi içermeyen bir deneme Excel'i oluşturur."""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font


def create_sample(output: Path) -> None:
    workbook = Workbook()
    products = workbook.active
    products.title = "Products"
    products.append([
        "Product ID", "Name", "Model", "Price", "Quantity", "Categories",
        "Description", "Image", "Meta Title", "Date Available",
    ])
    products.append([
        1, "Örnek Defter", "K000001", 100, 50, "1",
        "Deneme amacıyla oluşturulmuş ürün açıklaması.",
        "catalog/demo/defter.jpg", "Örnek Defter", "2026-01-01",
    ])
    products["B2"].font = Font(bold=True)
    products["D2"].number_format = "0.00"
    workbook.save(output)
    workbook.close()


if __name__ == "__main__":
    path = Path(__file__).with_name("products.xlsx")
    if path.exists():
        raise SystemExit(f"Dosya zaten mevcut, üzerine yazılmadı: {path}")
    create_sample(path)
    print(f"Örnek kaynak oluşturuldu: {path}")
