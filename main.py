"""Excel ürün içe aktarma otomasyonu.

Kaynak ürün Excel'ini asla değiştirmez. Yeni ürünleri, seçilen bir şablon
ürünün bütün alanlarını kopyalayarak ayrı bir yükleme dosyasına ekler.
"""

from __future__ import annotations

import argparse
import math
import sys
import tkinter as tk
from copy import copy
from datetime import datetime
from pathlib import Path
from typing import Any
from tkinter import filedialog, messagebox, ttk

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet


PRODUCT_SHEET = "Products"
REQUIRED_PRODUCT_COLUMNS = {"Product ID", "Name", "Model", "Price"}
APP_DIRECTORY = Path(__file__).resolve().parent
ENTRY_COLUMNS = [
    "Ürün Adı*",
    "Model (boşsa otomatik)",
    "Fiyat*",
    "Stok (boşsa şablondan)",
    "Görsel Yolu (boşsa şablondan)",
    "Şablon Ürün Modeli*",
    "Kategori ID (boşsa şablondan)",
    "Açıklama (boşsa şablondan)",
    "SEO Anahtar Kelime",
    "İlgili Ürün ID'leri",
    "Yayın Tarihi (YYYY-AA-GG)",
]


class ValidationError(Exception):
    """Kullanıcı tarafından giderilebilecek veri hatası."""


def validate_output_path(output_path: Path, *input_paths: Path) -> None:
    """Kaynak veya giriş dosyasının yanlışlıkla ezilmesini önler."""
    for input_path in input_paths:
        if output_path.resolve() == input_path.resolve() or (
            output_path.exists() and input_path.exists() and output_path.samefile(input_path)
        ):
            raise ValidationError("Sonuç dosyası kaynak veya giriş dosyasıyla aynı olamaz.")


def as_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def header_map(ws: Worksheet) -> dict[str, int]:
    return {as_text(cell.value): cell.column for cell in ws[1] if cell.value is not None}


def product_sheet(workbook_path: Path) -> tuple[Any, Worksheet, dict[str, int]]:
    if not workbook_path.exists():
        raise ValidationError(f"Kaynak dosya bulunamadı: {workbook_path}")
    wb = load_workbook(workbook_path)
    if PRODUCT_SHEET not in wb.sheetnames:
        raise ValidationError(f"'{PRODUCT_SHEET}' sayfası bulunamadı.")
    ws = wb[PRODUCT_SHEET]
    columns = header_map(ws)
    missing = REQUIRED_PRODUCT_COLUMNS - columns.keys()
    if missing:
        raise ValidationError("Kaynak Excel'de zorunlu sütunlar eksik: " + ", ".join(sorted(missing)))
    return wb, ws, columns


def create_entry_template(source_path: Path, output_path: Path) -> None:
    """Yeni ürünlerin girileceği yardımcı Excel'i üretir."""
    validate_output_path(output_path, source_path)
    _, products, columns = product_sheet(source_path)
    wb = Workbook()
    entry = wb.active
    entry.title = "Yeni Ürünler"
    entry.append(ENTRY_COLUMNS)
    entry.freeze_panes = "A2"
    entry.auto_filter.ref = f"A1:K1"

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in entry[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
    widths = [34, 24, 14, 23, 46, 25, 30, 46, 26, 26, 25]
    for number, width in enumerate(widths, start=1):
        entry.column_dimensions[chr(64 + number)].width = width

    info = wb.create_sheet("Kullanım")
    info.column_dimensions["A"].width = 115
    instructions = [
        "1. 'Yeni Ürünler' sayfasında her satıra bir ürün girin.",
        "2. Yıldızlı alanlar zorunludur: ürün adı, fiyat ve şablon ürün modeli.",
        "3. Şablon Ürün Modeli, aşağıdaki 'Mevcut Ürünler' sayfasındaki Model değerlerinden biri olmalıdır.",
        "4. Boş bırakılan alanlar şablon üründen kopyalanır; kategori ID alanına yalnızca doğruluğundan eminseniz yazın.",
        "5. Model boş bırakılırsa K ile başlayan sıradaki benzersiz model kodu otomatik atanır.",
        "6. Bu dosyayı kaydedip otomasyonda 'Yüklemeye hazır Excel üret' seçeneğini kullanın.",
        "Kaynak Excel değiştirilmez; sonuç ayrı bir dosya olarak üretilir.",
    ]
    for line in instructions:
        info.append([line])
    info["A1"].font = Font(bold=True)

    sample = wb.create_sheet("Mevcut Ürünler")
    sample.append(["Model", "Ürün Adı", "Kategori ID", "Fiyat", "Stok", "Görsel Yolu"])
    for cell in sample[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
    sample.freeze_panes = "A2"
    sample.auto_filter.ref = f"A1:F{products.max_row}"
    wanted = ["Model", "Name", "Categories", "Price", "Quantity", "Image"]
    for source_row in products.iter_rows(min_row=2, values_only=True):
        if source_row[columns["Model"] - 1] is not None:
            sample.append([source_row[columns[name] - 1] if name in columns else None for name in wanted])
    for column, width in {"A": 24, "B": 42, "C": 28, "D": 14, "E": 12, "F": 48}.items():
        sample.column_dimensions[column].width = width

    wb.save(output_path)


def entry_header_map(ws: Worksheet) -> dict[str, int]:
    actual = header_map(ws)
    missing = set(ENTRY_COLUMNS[:6]) - actual.keys()
    if missing:
        raise ValidationError(
            "Giriş dosyası beklenen şablonda değil. Eksik sütunlar: " + ", ".join(sorted(missing))
        )
    return actual


def parse_number(value: Any, row_number: int, label: str) -> float | int:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if not math.isfinite(value) or value < 0:
            raise ValidationError(f"Giriş satırı {row_number}: {label} sonlu ve sıfır veya daha büyük olmalıdır.")
        if label == "stok" and value != int(value):
            raise ValidationError(f"Giriş satırı {row_number}: stok tam sayı olmalıdır.")
        return value
    text = as_text(value).replace("₺", "").replace(" ", "")
    if not text:
        raise ValidationError(f"Giriş satırı {row_number}: {label} boş olamaz.")
    # Türkçe 1.250,50 ve 1250.50 yazımlarının ikisini de kabul eder.
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", ".")
    try:
        number = float(text)
    except ValueError as exc:
        raise ValidationError(f"Giriş satırı {row_number}: {label} sayı olmalıdır.") from exc
    if not math.isfinite(number) or number < 0:
        raise ValidationError(f"Giriş satırı {row_number}: {label} sonlu ve sıfır veya daha büyük olmalıdır.")
    if label == "stok" and not number.is_integer():
        raise ValidationError(f"Giriş satırı {row_number}: stok tam sayı olmalıdır.")
    return int(number) if number.is_integer() else number


def parse_date(value: Any, row_number: int) -> datetime:
    if isinstance(value, datetime):
        return value
    text = as_text(value)
    if not text:
        return datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    try:
        return datetime.strptime(text, "%Y-%m-%d")
    except ValueError as exc:
        raise ValidationError(f"Giriş satırı {row_number}: yayın tarihi YYYY-AA-GG biçiminde olmalıdır.") from exc


def next_model(existing_models: set[str]) -> str:
    numbers = []
    for model in existing_models:
        if model.startswith("K") and model[1:].isdigit():
            numbers.append(int(model[1:]))
    number = (max(numbers) if numbers else 0) + 1
    while f"K{number:06d}" in existing_models:
        number += 1
    return f"K{number:06d}"


def copy_product_row(source_ws: Worksheet, source_row: int, target_ws: Worksheet, target_row: int) -> None:
    """Hücre değerleri ile biçimlerini birlikte kopyalar."""
    for column in range(1, source_ws.max_column + 1):
        source = source_ws.cell(source_row, column)
        target = target_ws.cell(target_row, column, source.value)
        if source.has_style:
            target._style = copy(source._style)
        if source.number_format:
            target.number_format = source.number_format
        if source.hyperlink:
            target._hyperlink = copy(source.hyperlink)
        if source.comment:
            target.comment = copy(source.comment)
    target_ws.row_dimensions[target_row].height = source_ws.row_dimensions[source_row].height


def set_if_present(ws: Worksheet, row: int, columns: dict[str, int], column_name: str, value: Any) -> None:
    if column_name in columns:
        ws.cell(row, columns[column_name], value)


def build_upload_from_records(
    source_path: Path,
    entry_rows: list[tuple[int, dict[str, Any]]],
    output_path: Path,
) -> int:
    """Arayüzden veya giriş Excel'inden gelen ürün satırlarıyla sonuç dosyasını üretir."""
    validate_output_path(output_path, source_path)
    source_wb, products, columns = product_sheet(source_path)
    model_column = columns["Model"]
    id_column = columns["Product ID"]
    template_rows: dict[str, int] = {}
    existing_models: set[str] = set()
    product_ids: list[int] = []
    for row in range(2, products.max_row + 1):
        model = as_text(products.cell(row, model_column).value)
        if model:
            template_rows.setdefault(model, row)
            existing_models.add(model)
        product_id = products.cell(row, id_column).value
        if isinstance(product_id, int):
            product_ids.append(product_id)

    if not product_ids:
        raise ValidationError("Kaynak Excel'de kullanılabilir Product ID bulunamadı.")
    next_id = max(product_ids) + 1
    added = 0

    for entry_row_number, values in entry_rows:
        # Tamamen boş satırları atla.
        if not any(as_text(value) for value in values.values()):
            continue
        product_name = as_text(values["Ürün Adı*"])
        template_model = as_text(values["Şablon Ürün Modeli*"])
        if not product_name or not template_model:
            raise ValidationError(f"Giriş satırı {entry_row_number}: ürün adı ve şablon ürün modeli zorunludur.")
        if template_model not in template_rows:
            raise ValidationError(
                f"Giriş satırı {entry_row_number}: '{template_model}' adlı şablon model kaynak Excel'de yok."
            )

        model = as_text(values["Model (boşsa otomatik)"])
        if not model:
            model = next_model(existing_models)
        if model in existing_models:
            raise ValidationError(f"Giriş satırı {entry_row_number}: '{model}' model kodu zaten kullanılıyor.")
        price = parse_number(values["Fiyat*"], entry_row_number, "fiyat")
        quantity_value = values.get("Stok (boşsa şablondan)")
        if as_text(quantity_value):
            quantity_value = parse_number(quantity_value, entry_row_number, "stok")

        new_row = products.max_row + 1
        template_row = template_rows[template_model]
        template_name = as_text(products.cell(template_row, columns["Name"]).value)
        copy_product_row(products, template_row, products, new_row)

        set_if_present(products, new_row, columns, "Product ID", next_id)
        set_if_present(products, new_row, columns, "Name", product_name)
        set_if_present(products, new_row, columns, "Model", model)
        set_if_present(products, new_row, columns, "Price", price)
        set_if_present(
            products,
            new_row,
            columns,
            "Date Available",
            parse_date(values.get("Yayın Tarihi (YYYY-AA-GG)"), entry_row_number),
        )
        if as_text(quantity_value):
            set_if_present(products, new_row, columns, "Quantity", quantity_value)

        optional_mapping = {
            "Görsel Yolu (boşsa şablondan)": "Image",
            "Kategori ID (boşsa şablondan)": "Categories",
            "Açıklama (boşsa şablondan)": "Description",
            "SEO Anahtar Kelime": "SEO Keyword",
            "İlgili Ürün ID'leri": "Related Products",
        }
        for entry_name, product_name_column in optional_mapping.items():
            value = values.get(entry_name)
            if as_text(value):
                set_if_present(products, new_row, columns, product_name_column, value)

        # Şablonda ürün adını kullanan SEO metinlerini yeni ürün adına uyarla.
        for field in ("Meta Title", "Meta Keyword", "Meta Description", "Tags"):
            if field in columns:
                current = products.cell(new_row, columns[field]).value
                if isinstance(current, str) and template_name:
                    products.cell(new_row, columns[field], current.replace(template_name, product_name))

        existing_models.add(model)
        next_id += 1
        added += 1

    if not added:
        raise ValidationError("Giriş dosyasında işlenecek ürün satırı yok.")
    source_wb.save(output_path)
    return added


def build_upload_workbook(source_path: Path, entries_path: Path, output_path: Path) -> int:
    """Giriş Excel'indeki satırları doğrular ve sonuç Excel'ini üretir."""
    validate_output_path(output_path, source_path, entries_path)
    entry_wb = load_workbook(entries_path, data_only=True)
    if "Yeni Ürünler" not in entry_wb.sheetnames:
        raise ValidationError("Giriş dosyasında 'Yeni Ürünler' sayfası bulunamadı.")
    entries = entry_wb["Yeni Ürünler"]
    entry_columns = entry_header_map(entries)
    entry_rows = [
        (
            row_number,
            {name: entries.cell(row_number, column).value for name, column in entry_columns.items()},
        )
        for row_number in range(2, entries.max_row + 1)
    ]
    return build_upload_from_records(source_path, entry_rows, output_path)


def choose_path(question: str, default: str) -> Path:
    answer = input(f"{question} [{default}]: ").strip()
    return Path(answer or default)


class ProductExcelApp:
    """Excel işlemlerini komut satırı yerine pencereden çalıştırır."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Ürün Excel Otomasyonu")
        self.root.minsize(960, 700)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        default_source = APP_DIRECTORY / "A12.xlsx"
        if not default_source.exists():
            default_source = APP_DIRECTORY / "examples" / "products.xlsx"
        self.source_path = tk.StringVar(value=str(default_source))
        self.upload_output_path = tk.StringVar(value=str(APP_DIRECTORY / "yuklemeye_hazir.xlsx"))
        self.status = tk.StringVar(value="Kaynak Excel'i seçin ve şablon modellerini yükleyin.")
        self.template_models: list[str] = []
        self.product_rows: dict[str, dict[str, str]] = {}
        self.product_counter = 0
        self.form_values = {column: tk.StringVar() for column in ENTRY_COLUMNS}

        container = ttk.Frame(root, padding=20)
        container.grid(sticky="nsew")
        container.columnconfigure(0, weight=1)

        ttk.Label(container, text="Ürün Excel Otomasyonu", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            container,
            text="Ürünleri doğrudan bu ekrana girin. Kaynak Excel dosyası değişmez; sonuç yeni bir .xlsx dosyasıdır.",
            wraplength=900,
        ).grid(row=1, column=0, sticky="w", pady=(4, 16))

        notebook = ttk.Notebook(container)
        notebook.grid(row=2, column=0, sticky="nsew")
        container.rowconfigure(2, weight=1)
        self._build_product_entry_tab(notebook)

        ttk.Separator(container).grid(row=3, column=0, sticky="ew", pady=(14, 8))
        ttk.Label(container, textvariable=self.status, wraplength=900).grid(row=4, column=0, sticky="w")

    @staticmethod
    def _entry_row(
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        browse_command: Any,
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=6)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=6)
        ttk.Button(parent, text="Gözat…", command=browse_command).grid(row=row, column=2, padx=(8, 0), pady=6)

    def _build_product_entry_tab(self, notebook: ttk.Notebook) -> None:
        tab = ttk.Frame(notebook, padding=16)
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(3, weight=1)
        notebook.add(tab, text="Ürün girişi")

        source_frame = ttk.LabelFrame(tab, text="Kaynak ve sonuç", padding=10)
        source_frame.grid(row=0, column=0, sticky="ew")
        source_frame.columnconfigure(1, weight=1)
        self._entry_row(source_frame, 0, "Kaynak ürün Excel'i", self.source_path, self._choose_source)
        ttk.Button(source_frame, text="Şablon modellerini yükle", command=self._load_template_models).grid(
            row=0, column=3, padx=(8, 0), pady=6
        )
        self._entry_row(source_frame, 1, "Sonucun kaydedileceği yer", self.upload_output_path, self._choose_upload_output)

        form_frame = ttk.LabelFrame(tab, text="Yeni ürün bilgileri", padding=10)
        form_frame.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        for column in (1, 3):
            form_frame.columnconfigure(column, weight=1)
        fields = [
            ("Ürün Adı*", "Ürün adı"),
            ("Fiyat*", "Fiyat"),
            ("Model (boşsa otomatik)", "Model (boşsa otomatik)"),
            ("Şablon Ürün Modeli*", "Şablon ürün modeli"),
            ("Stok (boşsa şablondan)", "Stok"),
            ("Kategori ID (boşsa şablondan)", "Kategori ID"),
            ("Görsel Yolu (boşsa şablondan)", "Görsel yolu"),
            ("Yayın Tarihi (YYYY-AA-GG)", "Yayın tarihi"),
            ("Açıklama (boşsa şablondan)", "Açıklama"),
            ("SEO Anahtar Kelime", "SEO anahtar kelime"),
            ("İlgili Ürün ID'leri", "İlgili ürün ID'leri"),
        ]
        for index, (field, label) in enumerate(fields):
            row, pair_column = divmod(index, 2)
            column = pair_column * 2
            ttk.Label(form_frame, text=label).grid(row=row, column=column, sticky="w", padx=(0, 8), pady=4)
            if field == "Şablon Ürün Modeli*":
                self.template_model_box = ttk.Combobox(
                    form_frame, textvariable=self.form_values[field], state="normal"
                )
                self.template_model_box.grid(row=row, column=column + 1, sticky="ew", padx=(0, 18), pady=4)
            else:
                ttk.Entry(form_frame, textvariable=self.form_values[field]).grid(
                    row=row, column=column + 1, sticky="ew", padx=(0, 18), pady=4
                )
        ttk.Button(form_frame, text="Ürünü listeye ekle", command=self._add_product).grid(
            row=6, column=3, sticky="e", pady=(10, 0)
        )

        list_frame = ttk.LabelFrame(tab, text="Eklenecek ürünler", padding=10)
        list_frame.grid(row=3, column=0, sticky="nsew", pady=(12, 0))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        columns = ("name", "model", "price", "template", "quantity")
        self.product_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=8)
        headings = {
            "name": "Ürün adı",
            "model": "Model",
            "price": "Fiyat",
            "template": "Şablon model",
            "quantity": "Stok",
        }
        widths = {"name": 270, "model": 130, "price": 110, "template": 160, "quantity": 90}
        for column in columns:
            self.product_tree.heading(column, text=headings[column])
            self.product_tree.column(column, width=widths[column], anchor="w")
        self.product_tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.product_tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.product_tree.configure(yscrollcommand=scrollbar.set)
        buttons = ttk.Frame(list_frame)
        buttons.grid(row=1, column=0, columnspan=2, sticky="e", pady=(8, 0))
        ttk.Button(buttons, text="Seçileni sil", command=self._remove_selected_product).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(buttons, text="Listeyi temizle", command=self._clear_product_list).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(buttons, text="Yüklemeye hazır Excel oluştur", command=self._build_upload).grid(row=0, column=2)

    def _load_template_models(self, show_error: bool = True) -> bool:
        try:
            workbook, products, columns = product_sheet(Path(self.source_path.get().strip()))
            models = sorted(
                {as_text(products.cell(row, columns["Model"]).value) for row in range(2, products.max_row + 1)} - {""}
            )
            workbook.close()
            if not models:
                raise ValidationError("Kaynak Excel'de seçilebilecek model bulunamadı.")
        except (ValidationError, OSError, PermissionError) as exc:
            self.status.set(f"Hata: {exc}")
            if show_error:
                messagebox.showerror("Şablon modelleri yüklenemedi", str(exc), parent=self.root)
            return False
        self.template_models = models
        self.template_model_box["values"] = models
        self.status.set(f"{len(models)} şablon modeli yüklendi. Ürün bilgilerini girip listeye ekleyin.")
        return True

    def _add_product(self) -> None:
        if not self.template_models and not self._load_template_models():
            return
        values = {field: variable.get().strip() for field, variable in self.form_values.items()}
        product_name = values["Ürün Adı*"]
        price = values["Fiyat*"]
        template = values["Şablon Ürün Modeli*"]
        if not product_name or not price or not template:
            messagebox.showerror("Eksik bilgi", "Ürün adı, fiyat ve şablon ürün modeli zorunludur.", parent=self.root)
            return
        if template not in self.template_models:
            messagebox.showerror("Geçersiz şablon", "Şablon ürün modeli kaynak Excel'deki modellerden biri olmalıdır.", parent=self.root)
            return
        try:
            parse_number(price, self.product_counter + 1, "fiyat")
            if values["Stok (boşsa şablondan)"]:
                parse_number(values["Stok (boşsa şablondan)"], self.product_counter + 1, "stok")
            if values["Yayın Tarihi (YYYY-AA-GG)"]:
                parse_date(values["Yayın Tarihi (YYYY-AA-GG)"], self.product_counter + 1)
        except ValidationError as exc:
            messagebox.showerror("Geçersiz bilgi", str(exc), parent=self.root)
            return
        self.product_counter += 1
        item_id = f"product-{self.product_counter}"
        self.product_rows[item_id] = values
        self.product_tree.insert(
            "",
            "end",
            iid=item_id,
            values=(product_name, values["Model (boşsa otomatik)"] or "Otomatik", price, template, values["Stok (boşsa şablondan)"] or "Şablondan"),
        )
        for field, variable in self.form_values.items():
            if field != "Şablon Ürün Modeli*":
                variable.set("")
        self.status.set(f"Ürün listeye eklendi. Toplam: {len(self.product_rows)}")

    def _remove_selected_product(self) -> None:
        selected = self.product_tree.selection()
        if not selected:
            return
        for item_id in selected:
            self.product_tree.delete(item_id)
            self.product_rows.pop(item_id, None)
        self.status.set(f"Listede {len(self.product_rows)} ürün kaldı.")

    def _clear_product_list(self) -> None:
        if not self.product_rows:
            return
        if not messagebox.askyesno("Listeyi temizle", "Listedeki tüm ürünler silinsin mi?", parent=self.root):
            return
        for item_id in self.product_tree.get_children():
            self.product_tree.delete(item_id)
        self.product_rows.clear()
        self.status.set("Ürün listesi temizlendi.")

    def _choose_open_file(self, variable: tk.StringVar, title: str) -> None:
        selected = filedialog.askopenfilename(
            parent=self.root,
            title=title,
            initialdir=str(Path(variable.get()).parent),
            filetypes=[("Excel dosyaları", "*.xlsx"), ("Tüm dosyalar", "*.*")],
        )
        if selected:
            variable.set(selected)

    def _choose_save_file(self, variable: tk.StringVar, title: str) -> None:
        selected = filedialog.asksaveasfilename(
            parent=self.root,
            title=title,
            initialdir=str(Path(variable.get()).parent),
            initialfile=Path(variable.get()).name,
            defaultextension=".xlsx",
            filetypes=[("Excel dosyaları", "*.xlsx")],
        )
        if selected:
            variable.set(selected)

    def _choose_source(self) -> None:
        self._choose_open_file(self.source_path, "Kaynak ürün Excel'ini seçin")

    def _choose_upload_output(self) -> None:
        self._choose_save_file(self.upload_output_path, "Yüklemeye hazır Excel'i kaydedin")

    def _run_action(self, action: Any, success_message: Any) -> None:
        try:
            self.root.config(cursor="wait")
            self.root.update_idletasks()
            result = action()
        except (ValidationError, OSError, PermissionError) as exc:
            self.status.set(f"Hata: {exc}")
            messagebox.showerror("İşlem tamamlanamadı", str(exc), parent=self.root)
        except Exception as exc:  # Beklenmeyen hata için kullanıcıya anlaşılır bildirim.
            self.status.set("Beklenmeyen bir hata oluştu.")
            messagebox.showerror("Beklenmeyen hata", str(exc), parent=self.root)
        else:
            message = success_message(result)
            self.status.set(message)
            messagebox.showinfo("İşlem tamamlandı", message, parent=self.root)
        finally:
            self.root.config(cursor="")

    def _build_upload(self) -> None:
        if not self.product_rows:
            messagebox.showerror("Ürün yok", "Önce en az bir ürünü listeye ekleyin.", parent=self.root)
            return
        source = Path(self.source_path.get().strip())
        output = Path(self.upload_output_path.get().strip())
        entry_rows = [
            (index + 1, self.product_rows[item_id])
            for index, item_id in enumerate(self.product_tree.get_children())
        ]
        self._run_action(
            lambda: build_upload_from_records(source, entry_rows, output),
            lambda added: f"{added} ürün eklendi:\n{output.resolve()}",
        )


def launch_gui() -> int:
    root = tk.Tk()
    style = ttk.Style(root)
    style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"))
    ProductExcelApp(root)
    root.mainloop()
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Ürün Excel içe aktarma otomasyonu")
    subparsers = parser.add_subparsers(dest="command")
    template_parser = subparsers.add_parser("sablon", help="Ürün giriş Excel'i oluştur")
    template_parser.add_argument("source", type=Path)
    template_parser.add_argument("output", type=Path)
    build_parser = subparsers.add_parser("olustur", help="Yüklemeye hazır Excel oluştur")
    build_parser.add_argument("source", type=Path)
    build_parser.add_argument("entries", type=Path)
    build_parser.add_argument("output", type=Path)
    args = parser.parse_args()
    try:
        if args.command == "sablon":
            create_entry_template(args.source, args.output)
            print(f"Giriş şablonu oluşturuldu: {args.output}")
        elif args.command == "olustur":
            added = build_upload_workbook(args.source, args.entries, args.output)
            print(f"{added} ürün eklendi: {args.output}")
        else:
            return launch_gui()
    except ValidationError as exc:
        print(f"HATA: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
