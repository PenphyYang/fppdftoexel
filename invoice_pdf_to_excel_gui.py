import argparse
import queue
import re
import sys
import threading
import traceback
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import pdfplumber
import tkinter as tk
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from tkinter import filedialog, messagebox, ttk


APP_NAME = "发票分析工具"
APP_TITLE = "发票分析工具（ 只是一个路过的网友：Perfect）"
SUMMARY_SHEET_NAME = "发票汇总"
DETAIL_SHEET_NAME = "发票明细"

SKIP_LINE_PATTERNS = [
    "电子发票",
    "发票号码",
    "开票日期",
    "共",
    "第",
    "购买方信息",
    "销售方信息",
    "购",
    "买",
    "销",
    "售",
    "方",
    "信",
    "息",
    "项目名称 规格型号",
    "项目名称",
]

DETAIL_END_MARKERS = ["小 计", "小计", "合 计", "合计", "开票人"]

ITEM_LINE_RE = re.compile(
    r"^(?P<body>.+?)\s+"
    r"(?P<unit>[^\s]+)\s+"
    r"(?P<qty>-?\d+(?:\.\d+)?)\s+"
    r"(?P<unit_price>-?\d+(?:\.\d+)?)\s+"
    r"(?P<amount>-?\d+(?:\.\d+)?)\s+"
    r"(?P<tax_rate>\d+(?:\.\d+)?%)\s+"
    r"(?P<tax_amount>-?\d+(?:\.\d+)?)$"
)


@dataclass
class InvoiceParseResult:
    summary: Dict[str, object]
    items: List[Dict[str, object]]


def clean_text(text: str) -> str:
    return text.replace("\xa0", " ").replace("\u3000", " ").strip()


def compact_spaces(text: str) -> str:
    return re.sub(r"[ \t]+", " ", clean_text(text))


def to_decimal(value: str) -> Optional[Decimal]:
    try:
        return Decimal(str(value).replace(",", "").replace("¥", "").strip())
    except (InvalidOperation, AttributeError):
        return None


def to_float(value: str) -> Optional[float]:
    number = to_decimal(value)
    return float(number) if number is not None else None


def extract_pages(pdf_path: Path) -> List[Dict[str, object]]:
    pages: List[Dict[str, object]] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(x_tolerance=2, y_tolerance=2) or ""
            lines = [compact_spaces(line) for line in text.splitlines() if compact_spaces(line)]
            pages.append(
                {
                    "page_number": page_number,
                    "text": clean_text(text),
                    "lines": lines,
                }
            )
    return pages


def extract_field(pattern: str, text: str) -> str:
    match = re.search(pattern, text, re.MULTILINE)
    return clean_text(match.group(1)) if match else ""


def extract_company_sections(full_text: str) -> Dict[str, str]:
    buyer_match = re.search(
        r"名称：(?P<buyer_name>.*?)\s+名称：(?P<seller_name>.*?)\n",
        full_text,
        re.MULTILINE,
    )
    tax_ids = re.findall(
        r"统一社会信用代码/纳税人识别号：([0-9A-Z]+)",
        full_text,
        re.MULTILINE,
    )
    return {
        "buyer_name": clean_text(buyer_match.group("buyer_name")) if buyer_match else "",
        "seller_name": clean_text(buyer_match.group("seller_name")) if buyer_match else "",
        "buyer_tax_id": tax_ids[0] if len(tax_ids) > 0 else "",
        "seller_tax_id": tax_ids[1] if len(tax_ids) > 1 else "",
    }


def extract_totals(full_text: str) -> Dict[str, object]:
    total_section_match = re.search(
        r"合\s*计\s*\n?\s*¥?([0-9,]+\.\d{2})\s*¥?([0-9,]+\.\d{2})",
        full_text,
        re.MULTILINE,
    )
    amount_total = to_float(total_section_match.group(1)) if total_section_match else None
    tax_total = to_float(total_section_match.group(2)) if total_section_match else None
    grand_total = round(amount_total + tax_total, 2) if amount_total is not None and tax_total is not None else None
    return {
        "amount_total": amount_total,
        "tax_total": tax_total,
        "grand_total": grand_total,
    }


def should_skip_prefix_line(line: str) -> bool:
    if not line:
        return True
    if any(marker in line for marker in DETAIL_END_MARKERS):
        return True
    if any(pattern == line or pattern in line for pattern in SKIP_LINE_PATTERNS):
        return True
    if line.startswith("名称：") or line.startswith("统一社会信用代码/纳税人识别号："):
        return True
    return False


def parse_items_from_page(lines: List[str], page_number: int, source_file: str) -> List[Dict[str, object]]:
    items: List[Dict[str, object]] = []
    in_detail = False
    prefix_buffer: List[str] = []
    last_item: Optional[Dict[str, object]] = None

    for line in lines:
        if not in_detail:
            if re.search(r"项目名称.*税\s*额", line):
                in_detail = True
            continue

        if any(marker in line for marker in DETAIL_END_MARKERS):
            break

        match = ITEM_LINE_RE.match(line)
        if match:
            description_parts = prefix_buffer + [match.group("body")]
            prefix_buffer = []
            description = " ".join(part for part in description_parts if part).strip()
            description = re.sub(r"\s+", " ", description)
            item = {
                "source_file": source_file,
                "page_number": page_number,
                "description": description,
                "unit": match.group("unit"),
                "quantity": to_float(match.group("qty")),
                "unit_price": to_float(match.group("unit_price")),
                "amount": to_float(match.group("amount")),
                "tax_rate": match.group("tax_rate"),
                "tax_amount": to_float(match.group("tax_amount")),
            }
            items.append(item)
            last_item = item
            continue

        if should_skip_prefix_line(line):
            continue

        if last_item is not None:
            last_item["description"] = f"{last_item['description']} {line}".strip()
        else:
            prefix_buffer.append(line)

    return items


def parse_invoice(pdf_path: Path) -> InvoiceParseResult:
    pages = extract_pages(pdf_path)
    full_text = "\n".join(page["text"] for page in pages if page["text"])

    company_info = extract_company_sections(full_text)
    totals = extract_totals(full_text)
    items: List[Dict[str, object]] = []

    for page in pages:
        page_number = int(page["page_number"])
        items.extend(parse_items_from_page(page["lines"], page_number, pdf_path.name))

    item_amount_sum = round(sum(item["amount"] or 0 for item in items), 2) if items else None
    item_tax_sum = round(sum(item["tax_amount"] or 0 for item in items), 2) if items else None

    summary = {
        "source_file": pdf_path.name,
        "source_path": str(pdf_path.resolve()),
        "invoice_title": extract_field(r"^(电子发票[^\n]*)", full_text),
        "invoice_number": extract_field(r"发票号码：([0-9]+)", full_text),
        "invoice_date": extract_field(r"开票日期：([0-9]{4}年[0-9]{2}月[0-9]{2}日)", full_text),
        "issuer": extract_field(r"开票人：([^\n]+)", full_text),
        "page_count": len(pages),
        **company_info,
        **totals,
        "item_count": len(items),
        "item_amount_sum": item_amount_sum,
        "item_tax_sum": item_tax_sum,
        "parse_status": "成功" if full_text else "未提取到文本",
    }
    return InvoiceParseResult(summary=summary, items=items)


def auto_fit_worksheet(worksheet) -> None:
    for column_cells in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 60)

    for cell in worksheet[1]:
        cell.font = Font(bold=True)


def export_to_excel(results: List[InvoiceParseResult], output_path: Path) -> None:
    summary_rows = [result.summary for result in results]
    item_rows = [item for result in results for item in result.items]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame(summary_rows).to_excel(writer, index=False, sheet_name=SUMMARY_SHEET_NAME)
        pd.DataFrame(item_rows).to_excel(writer, index=False, sheet_name=DETAIL_SHEET_NAME)

        workbook = writer.book
        for sheet_name in [SUMMARY_SHEET_NAME, DETAIL_SHEET_NAME]:
            worksheet = workbook[sheet_name]
            auto_fit_worksheet(worksheet)
            worksheet.freeze_panes = "A2"


def process_pdfs(pdf_paths: List[Path], output_path: Path, progress_callback=None) -> List[InvoiceParseResult]:
    results = []
    total = len(pdf_paths)
    for index, pdf_path in enumerate(pdf_paths, start=1):
        if progress_callback:
            progress_callback(f"正在解析：{pdf_path.name}", index - 1, total)
        result = parse_invoice(pdf_path)
        results.append(result)
    if progress_callback:
        progress_callback("正在写入 Excel...", total, total)
    export_to_excel(results, output_path)
    if progress_callback:
        progress_callback(f"处理完成：{output_path.name}", total, total)
    return results


class InvoiceExtractorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("920x620")
        self.root.minsize(840, 560)

        self.queue: "queue.Queue[tuple]" = queue.Queue()
        self.worker_thread: Optional[threading.Thread] = None
        self.pdf_paths: List[Path] = []

        self.set_window_icon()

        self.status_var = tk.StringVar(value="请选择一个或多个发票 PDF 文件。")
        self.output_var = tk.StringVar(value="")
        self.progress_var = tk.DoubleVar(value=0)

        self.build_ui()
        self.root.after(200, self.poll_queue)

    def set_window_icon(self) -> None:
        icon_path = Path(__file__).resolve().parent / "assets" / "invoice_tool_logo.png"
        if icon_path.exists():
            try:
                self.icon_image = tk.PhotoImage(file=str(icon_path))
                self.root.iconphoto(True, self.icon_image)
            except Exception:
                pass

    def build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=16)
        main.pack(fill="both", expand=True)

        header_frame = ttk.Frame(main)
        header_frame.pack(fill="x", pady=(0, 12))

        ttk.Label(header_frame, text=APP_NAME, font=("Microsoft YaHei UI", 18, "bold")).pack(anchor="w")
        ttk.Label(
            header_frame,
            text="批量提取发票 PDF 内容并生成 Excel",
            foreground="#666666",
        ).pack(anchor="w", pady=(4, 0))

        file_frame = ttk.LabelFrame(main, text="PDF 文件", padding=12)
        file_frame.pack(fill="both", expand=True)

        button_row = ttk.Frame(file_frame)
        button_row.pack(fill="x", pady=(0, 10))

        ttk.Button(button_row, text="添加 PDF", command=self.add_files).pack(side="left", padx=(0, 8))
        ttk.Button(button_row, text="移除选中", command=self.remove_selected).pack(side="left", padx=(0, 8))
        ttk.Button(button_row, text="清空列表", command=self.clear_files).pack(side="left")

        list_frame = ttk.Frame(file_frame)
        list_frame.pack(fill="both", expand=True)

        self.file_list = tk.Listbox(list_frame, selectmode=tk.EXTENDED, font=("Microsoft YaHei UI", 10))
        self.file_list.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_list.yview)
        scrollbar.pack(side="right", fill="y")
        self.file_list.config(yscrollcommand=scrollbar.set)

        output_frame = ttk.LabelFrame(main, text="导出设置", padding=12)
        output_frame.pack(fill="x", pady=12)

        ttk.Label(output_frame, text="Excel 输出路径：").grid(row=0, column=0, sticky="w")
        ttk.Entry(output_frame, textvariable=self.output_var).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(output_frame, text="选择路径", command=self.choose_output).grid(row=0, column=2, sticky="e")
        output_frame.columnconfigure(1, weight=1)

        run_frame = ttk.Frame(main)
        run_frame.pack(fill="x")

        ttk.Button(run_frame, text="开始分析", command=self.start_processing).pack(side="left")
        ttk.Button(run_frame, text="清空输出路径", command=self.clear_output).pack(side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(run_frame, variable=self.progress_var, maximum=100)
        self.progress.pack(side="left", fill="x", expand=True, padx=12)

        status_frame = ttk.LabelFrame(main, text="处理状态", padding=12)
        status_frame.pack(fill="both", expand=False, pady=(12, 0))

        ttk.Label(status_frame, textvariable=self.status_var, anchor="w").pack(fill="x")

        tip_text = (
            "导出结果包含两个工作表：发票汇总、发票明细。\n"
            "如果个别发票存在换行拆分或版式差异，建议抽样核对明细描述。"
        )
        ttk.Label(status_frame, text=tip_text, foreground="#555555", justify="left").pack(fill="x", pady=(10, 0))

    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择发票 PDF",
            filetypes=[("PDF 文件", "*.pdf")],
        )
        if not paths:
            return

        existing = {str(path.resolve()) for path in self.pdf_paths}
        for file_path in paths:
            path = Path(file_path)
            resolved = str(path.resolve())
            if resolved not in existing:
                self.pdf_paths.append(path)
                self.file_list.insert(tk.END, path.name)
                existing.add(resolved)

        self.status_var.set(f"已选择 {len(self.pdf_paths)} 个 PDF 文件。")

    def remove_selected(self) -> None:
        selected = list(self.file_list.curselection())
        if not selected:
            return
        for index in reversed(selected):
            self.file_list.delete(index)
            self.pdf_paths.pop(index)
        self.status_var.set(f"剩余 {len(self.pdf_paths)} 个 PDF 文件。")

    def clear_files(self) -> None:
        self.file_list.delete(0, tk.END)
        self.pdf_paths.clear()
        self.status_var.set("文件列表已清空。")

    def choose_output(self) -> None:
        initial_file = f"{APP_NAME}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        path = filedialog.asksaveasfilename(
            title="选择 Excel 输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile=initial_file,
        )
        if path:
            self.output_var.set(path)

    def clear_output(self) -> None:
        self.output_var.set("")

    def start_processing(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("提示", "当前任务还在处理，请稍候。")
            return

        if not self.pdf_paths:
            messagebox.showwarning("提示", "请先添加至少一个 PDF 文件。")
            return

        output_value = self.output_var.get().strip()
        if not output_value:
            self.choose_output()
            output_value = self.output_var.get().strip()
            if not output_value:
                self.status_var.set("已取消输出文件选择。")
                return

        output_path = Path(output_value)
        if output_path.suffix.lower() != ".xlsx":
            output_path = output_path.with_suffix(".xlsx")
            self.output_var.set(str(output_path))

        self.progress_var.set(0)
        self.status_var.set("准备开始处理...")

        self.worker_thread = threading.Thread(
            target=self.run_processing,
            args=(list(self.pdf_paths), output_path),
            daemon=True,
        )
        self.worker_thread.start()

    def run_processing(self, pdf_paths: List[Path], output_path: Path) -> None:
        try:
            def progress(message: str, current: int, total: int) -> None:
                percent = (current / total * 100) if total else 0
                self.queue.put(("progress", message, percent))

            results = process_pdfs(pdf_paths, output_path, progress_callback=progress)
            self.queue.put(("success", str(output_path), len(results)))
        except Exception as exc:
            self.queue.put(("error", str(exc)))

    def poll_queue(self) -> None:
        try:
            while True:
                item = self.queue.get_nowait()
                kind = item[0]

                if kind == "progress":
                    _, message, percent = item
                    self.status_var.set(message)
                    self.progress_var.set(percent)
                elif kind == "success":
                    _, output_path, count = item
                    self.progress_var.set(100)
                    self.status_var.set(f"处理完成，共分析 {count} 份发票。")
                    messagebox.showinfo("完成", f"Excel 已生成：\n{output_path}")
                elif kind == "error":
                    _, error_message = item
                    self.status_var.set("处理失败，请查看错误信息。")
                    messagebox.showerror("错误", error_message)
        except queue.Empty:
            pass
        finally:
            self.root.after(200, self.poll_queue)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="批量提取发票 PDF 内容并导出 Excel。")
    parser.add_argument("--input", nargs="+", help="一个或多个 PDF 文件路径")
    parser.add_argument("--output", help="输出的 Excel 文件路径")
    return parser


def write_error_log(exc: Exception) -> Path:
    if getattr(sys, "frozen", False):
        base_dir = Path(sys.executable).resolve().parent
    else:
        base_dir = Path(__file__).resolve().parent

    log_path = base_dir / "invoice_pdf_to_excel_error.log"
    content = (
        f"time: {datetime.now():%Y-%m-%d %H:%M:%S}\n"
        f"error: {exc}\n\n"
        f"{traceback.format_exc()}"
    )
    log_path.write_text(content, encoding="utf-8")
    return log_path


def main() -> None:
    parser = build_arg_parser()
    args = parser.parse_args()

    if args.input:
        pdf_paths = [Path(path) for path in args.input]
        output_path = Path(args.output) if args.output else Path.cwd() / f"{APP_NAME}.xlsx"
        process_pdfs(pdf_paths, output_path)
        print(f"已生成 Excel：{output_path}")
        return

    root = tk.Tk()
    app = InvoiceExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        log_path = write_error_log(exc)
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("程序错误", f"程序执行失败，错误日志已写入：\n{log_path}")
            root.destroy()
        except Exception:
            pass
        raise
