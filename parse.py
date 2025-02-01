import os
import ast
import glob
import argparse
from pathlib import Path

from vars import *
from templates import *

from bs4 import BeautifulSoup
import xlwings as xw
from xlwings.constants import FileFormat


class Formulas2Html:
    default_args = {
        "update_links": False,
        "read_only": True,
        "ignore_read_only_recommended": True,
    }

    def __init__(self, *, file_path: str, password: str):
        self.file_path = args.file
        self.file_name = os.path.split(file_path)[1]
        self.main_dir = os.getcwd()
        self.all_formulas: dict[str, dict[str, str]] = {}
        self.selected_formulas: dict[str, str] = {}
        self.sheets_names: dict[str, str] = {}
        self.wb = xw.Book(file_path, password=password, **self.default_args)

    def has_formula(self, formula: str):
        return (formula.startswith("=") or formula.startswith("+")) and len(
            formula.strip()
        ) != 1

    def col_name(self, col: int):
        name = ""
        while col >= 0:
            name = chr(col % 26 + 65) + name
            col //= 26
            col -= 1
        return name

    def address_to_string(self, column: int, row: int):
        return f"{self.col_name(column)}{row + 1}"

    def is_input(self, text: str):
        try:
            float(text.strip().replace(",", "."))
            return True
        except:
            return False

    def escape_chars(self, html: str):
        replacements = {
            "<": "&lt;",
            ">": "&gt;",
            '"': "&quot;",
            "'": "&apos;",  # Alternatively, you can use "&#39;" for single quote
        }
        html.replace("&", "&amp;")
        for key, value in replacements.items():
            html = html.replace(key, value)
        return html

    def reflect_data(self, data_file: str, data_text: str):
        try:
            with open(data_file, "r") as f:
                raw = f.read()
                raw = raw[data_text.index("{") :][:-1]
                return ast.literal_eval(raw)
        except FileNotFoundError:
            pass

    def extract_formulas(self, range: xw.Range):
        formulas = {}
        for row in range.rows:
            for cell in row:
                if self.has_formula(cell.formula):
                    formulas[cell.address.replace("$", "")] = (
                        cell.formula.replace("FALSE", '"FALSE"')
                        .replace("TRUE", '"TRUE"')
                        .replace("FIXED(", "ROUND(")
                    )
        return formulas

    def extract_all_formulas(self):
        try:
            with open("formulas.js", "r") as f:
                raw_formulas = f.read()
                raw_formulas = raw_formulas[len("export var formulas = ") :][:-1]
                self.all_formulas = ast.literal_eval(raw_formulas)
        except FileNotFoundError:
            pass
        for idx, sheet in enumerate(self.wb.sheets):
            index_name = f"sheet{idx + 1:03d}"
            self.sheets_names[index_name] = sheet.name
        if len(self.all_formulas) != 0:
            return
        for idx, sheet in enumerate(self.wb.sheets):
            self.all_formulas[sheet.name] = self.extract_formulas(sheet.used_range)
        with open("formulas.js", "w", encoding="utf-8") as f:
            f.write(formulas.format(formulas=self.all_formulas))

    def extract_selection_formulas(self):
        try:
            with open("selected_formulas.js", "r") as f:
                raw_formulas = f.read()
                raw_formulas = raw_formulas[len("export var formulas = ") :][:-1]
                self.all_formulas = ast.literal_eval(raw_formulas)
        except FileNotFoundError:
            pass
        selectedCells = self.wb.app.selection
        for address in selectedCells.address.split(","):
            self.all_formulas[selectedCells.sheet.name] = self.all_formulas.get(
                selectedCells.sheet.name, {}
            )
            self.all_formulas[selectedCells.sheet.name].update(
                self.extract_formulas(selectedCells.sheet.range(selectedCells.address))
            )
            # self.all_formulas[selectedCells.sheet.name].update(
            #    self.extract_formulas(selectedCells.sheet.range(selectedCells.address))
            # )

        with open("selected_formulas.js", "w", encoding="utf-8") as f:
            f.write(formulas.format(formulas=self.all_formulas))

    def extract_html_col(self, *, text: str, cell_id: str, sheet_name: str):
        if text.strip() != "" or cell_id in self.all_formulas[sheet_name]:
            text = self.escape_chars(text)
            if cell_id in self.all_formulas[sheet_name]:
                return output_template.format(cell_id=replacements[cell_id.lower()], text=text)
            # FIXME: extract correct values that contains ","
            elif self.is_input(text):
                if styles == output and not allow_input_on_output:
                    return output_template.format(cell_id=replacements[cell_id.lower()], text=text)
                else:
                    return input_template.format(cell_id=replacements[cell_id.lower()], value=text)
            else:
                return col_template.format(text=text)
        return ""

    def selection_to_html(self):
        selectedCells = self.wb.app.selection
        new_html = ""
        for address in selectedCells.address.split(","):
            for row in selectedCells.sheet.range(selectedCells.address).rows:
                row_html = ""
                for cell in row:
                    id = cell.address.replace("$", "")
                    row_html += self.extract_html_col(
                        text= f"{cell.raw_value}" if cell.raw_value != None else "",
                        cell_id=id,
                        sheet_name=selectedCells.sheet.name,
                    )
                new_html += (
                    row_template.format(columns=row_html) if row_html != "" else ""
                )
        with open(
            os.path.join(self.main_dir, "selected.html"), "w", encoding="utf-8"
        ) as new:
            new.write(html_template.format(body=new_html))

    def table_data_as_js_obj(self):
        d: dict[str, list[float]] = {}
        selectedCells = self.wb.app.selection
        for address in selectedCells.address.split(","):
            for row in selectedCells.sheet.range(selectedCells.address).rows:
                curr = ""
                for i, cell in enumerate(row):
                    if i == 0:
                        if cell.raw_value:
                            curr = f"{cell.raw_value}"
                        d[curr] = []
                    else:
                        if cell.raw_value:
                            d[curr].append(cell.raw_value)
                        else:
                            d[curr].append(0)
        with open(
            os.path.join(self.main_dir, "table.js"), "w", encoding="utf-8"
        ) as data_js:
            data_js.write(data.format(data=d))

    def table_data_to_js(self):
        d: list[dict[str, dict[str, str]]] = []
        try:
            with open(
                os.path.join(self.main_dir, "data.js"), "r", encoding="utf-8"
            ) as f:
                raw = f.read()
                raw = raw[len("export var data = ") :][:-1]
                d = ast.literal_eval(raw)
        except FileNotFoundError:
            pass
        selectedCells = self.wb.app.selection
        table: dict[str, dict[str, str]] = {}
        for address in selectedCells.address.split(","):
            for row in selectedCells.sheet.range(selectedCells.address).rows:
                table[selectedCells.sheet.name] = table.get(
                    selectedCells.sheet.name, {}
                )
                for cell in row:
                    id = cell.address.replace("$", "")
                    if self.has_formula(cell.formula):
                        table[selectedCells.sheet.name][id] = cell.formula
                    elif cell.raw_value:
                        table[selectedCells.sheet.name][id] = f"{cell.raw_value}"
                    else:
                        table[selectedCells.sheet.name][id] = ""
        d.append(table)
        with open(
            os.path.join(self.main_dir, "data.js"), "w", encoding="utf-8"
        ) as data_js:
            data_js.write(data.format(data=d))

    def clean(self):
        for sheet in self.wb.sheets:
            used_range = sheet.used_range
            for row in used_range.rows:
                for cell in row:
                    error = f"[ERROR] Sheet: {sheet.name}, Cell {cell.address}"
                    try:
                        if (
                            cell.merge_cells
                            and not (cell.api.EntireColumn.Hidden)
                            and not (cell.api.EntireRow.Hidden)
                        ):
                            cell.unmerge()
                    except:
                        print(error)
            for shape in sheet.shapes:
                shape.delete()
            for picture in sheet.pictures:
                picture.delete()

    def cleaned_to_html(self):
        # Clean the book
        self.wb.activate(True)
        self.clean()
        os.makedirs("cleaned", exist_ok=True)
        tmp = os.path.join(self.main_dir, "cleaned")
        temp_path = os.path.join(tmp, self.file_name)
        self.wb.save(temp_path)
        self.wb.close()
        # Load the cleaned book and save to html
        wb = xw.Book(temp_path, editable=False, **self.default_args)
        try:
            wb.api.SaveAs(str(Path(temp_path).with_suffix("")), FileFormat.xlHtml)
        except:
            pass
        wb.close()
        for sheet in glob.glob(os.path.join(tmp, "**", "sheet*.htm"), recursive=True):
            # Parse the html and save to html
            name = Path(sheet).stem
            new_html = ""
            with open(
                sheet,
                encoding="iso-8859-1",
            ) as html:
                soup = BeautifulSoup(html, "html.parser")
                table = soup.find("table")
                for row, tr in enumerate(table.find_all("tr", recursive=False)):  # type: ignore
                    row_html = ""
                    column = 0
                    for td in tr.find_all("td", recursive=False):
                        id = self.address_to_string(column, row)
                        row_html += self.extract_html_col(
                            td.text, id, self.sheets_names[name]
                        )
                        offsetx = int(td.get("colspan")) if td.get("colspan") else 1
                        column = column + offsetx
                    new_html += (
                        row_template.format(columns=row_html) if row_html != "" else ""
                    )
            with open(
                os.path.join(self.main_dir, f"{name}.html"), "w", encoding="utf-8"
            ) as new:
                new.write(html_template.format(body=new_html))


parser = argparse.ArgumentParser(
    prog="formulas2excel", description="Converts excel formulas to html format."
)
parser.add_argument("file")
parser.add_argument("-p", "--password")
parser.add_argument("-f", "--formulas")
parser.add_argument("-s", "--selection")
parser.add_argument("-t", "--table")
args = parser.parse_args()

# TODO: extract list boxs too
fm2html = Formulas2Html(file_path=args.file, password=args.password)
with xw.App() as app:
    if args.table:
        fm2html.table_data_to_js()
    elif args.selection:
        fm2html.extract_selection_formulas()
        # fm2html.extract_all_formulas()
        fm2html.selection_to_html()
    elif args.formulas:
        fm2html.extract_all_formulas()
    else:
        fm2html.table_data_as_js_obj()
