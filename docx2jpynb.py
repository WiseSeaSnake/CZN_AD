# =============================================================
# docx2ipynb_v2.py — конвертер DOCX → ipynb с эвристикой по содержимому
# =============================================================
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
import nbformat as nbf
from pathlib import Path
import re


# ---------- 1. Служебные функции ----------
def iter_block_items(parent):
    body = parent.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)
        elif child.tag == qn("w:tbl"):
            yield Table(child, parent)


def table_to_markdown(table):
    rows = []
    for row in table.rows:
        cells = [c.text.strip().replace("\n", " ") for c in row.cells]
        rows.append(cells)
    if not rows:
        return ""
    n_cols = max(len(r) for r in rows)
    rows = [r + [""] * (n_cols - len(r)) for r in rows]
    header = "| " + " | ".join(rows[0]) + " |"
    sep    = "| " + " | ".join(["---"] * n_cols) + " |"
    body   = ["| " + " | ".join(r) + " |" for r in rows[1:]]
    return "\n".join([header, sep] + body)


# ---------- 2. ЭВРИСТИКА: является ли параграф кодом ----------
# Ключевые слова Python
PY_KEYWORDS = (
    "import ", "from ", "def ", "class ", "return ", "print(",
    "for ", "while ", "if ", "elif ", "else:", "try:", "except",
    "with ", "lambda", "yield", "raise ", "assert ",
)

# Известные модули/функции из этого курса
PY_LIBS = (
    "np.", "pd.", "plt.", "sklearn", "sklearn.", "numpy", "pandas",
    "matplotlib", "seaborn", "shap", "RandomForest", "LogisticRegression",
    "KMeans", "make_classification", "make_blobs", "load_iris",
    "train_test_split", "cross_val_score", "GridSearchCV",
    "classification_report", "confusion_matrix", "ConfusionMatrixDisplay",
    "roc_auc_score", "roc_curve", "silhouette_score", "StandardScaler",
    "fit(", "predict(", "fit_predict(", "transform(",
)

# Символы, характерные для кода
CODE_SYMBOLS = ("=", "(", ")", "[", "]", "{", "}", ":", ",", ".", "_")


def is_code_line(text: str) -> bool:
    """Определяет по содержимому строки, является ли она кодом."""
    t = text.strip()
    if not t:
        return False

    # Комментарий
    if t.startswith("#"):
        return True

    # Начинается с ключевого слова Python
    if any(t.startswith(kw) for kw in PY_KEYWORDS):
        return True

    # Содержит известную библиотеку/функцию
    if any(lib in t for lib in PY_LIBS):
        return True

    # Присваивание переменной: "x = ..." или "x: int = ..."
    if re.match(r"^[A-Za-z_][A-Za-z0-9_]*\s*=\s*\S", t):
        return True

    # Вызов метода/функции: "foo.bar(...)" или "foo(...)"
    if re.match(r"^[A-Za-z_][A-Za-z0-9_.]*\s*\(", t):
        return True

    # Строка вида: "plt.bar(...)" / "np.mean(...)"
    if re.match(r"^[A-Za-z_][A-Za-z0-9_]*\.[A-Za-z_]", t):
        return True

    # Отступы (код с блоком) — но не список
    if text.startswith(("    ", "\t")):
        # Список обычно начинается с - • 1. и т.п. — исключаем
        stripped = text.lstrip()
        if not stripped.startswith(("-", "•", "*", "o ")):
            return True

    return False


def is_code_block(lines: list[str]) -> bool:
    """Считаем блок кодом, если ≥ 60% строк похожи на код."""
    if not lines:
        return False
    hits = sum(1 for ln in lines if is_code_line(ln))
    return hits / len(lines) >= 0.6


# ---------- 3. Основная функция ----------
def docx_to_ipynb(docx_path, ipynb_path, title=None):
    doc = Document(docx_path)
    nb = nbf.v4.new_notebook()
    nb.cells = []

    title = title or Path(docx_path).stem
    nb.cells.append(nbf.v4.new_markdown_cell(f"# {title}"))

    md_buffer   = []   # накопление обычного текста
    code_buffer = []   # накопление кода

    def flush_md():
        if md_buffer:
            text = "\n\n".join(md_buffer).strip()
            if text:
                nb.cells.append(nbf.v4.new_markdown_cell(text))
            md_buffer.clear()

    def flush_code():
        if code_buffer:
            # Убираем пустые строки по краям
            while code_buffer and not code_buffer[0].strip():
                code_buffer.pop(0)
            while code_buffer and not code_buffer[-1].strip():
                code_buffer.pop()
            if not code_buffer:
                return
            code = "\n".join(code_buffer)
            # Проверка: блок действительно код?
            if is_code_block(code.split("\n")):
                nb.cells.append(nbf.v4.new_code_cell(code))
            else:
                # Не код — вернуть в markdown
                nb.cells.append(nbf.v4.new_markdown_cell(code))
            code_buffer.clear()

    for block in iter_block_items(doc):
        # --- Параграф ---
        if isinstance(block, Paragraph):
            text = block.text.rstrip()          # сохраняем отступы слева!
            if not text.strip():
                # Пустая строка — разделитель внутри кода
                if code_buffer:
                    code_buffer.append("")
                continue

            style = (block.style.name or "").lower()

            # Заголовок
            if style.startswith("heading"):
                flush_md()
                flush_code()
                level = "".join(filter(str.isdigit, style)) or "1"
                nb.cells.append(nbf.v4.new_markdown_cell(
                    f"{'#' * int(level)} {text.strip()}"
                ))
                continue

            # Код?
            if is_code_line(text):
                flush_md()
                code_buffer.append(text)
            else:
                flush_code()
                md_buffer.append(text.strip())

        # --- Таблица ---
        elif isinstance(block, Table):
            flush_md()
            flush_code()
            md = table_to_markdown(block)
            if md:
                nb.cells.append(nbf.v4.new_markdown_cell(md))

    flush_md()
    flush_code()

    nb["metadata"] = {
        "kernelspec": {"display_name": "Python 3",
                       "language": "python", "name": "python3"},
        "language_info": {"name": "python", "version": "3.11"},
    }

    nbf.write(nb, ipynb_path)
    n_md   = sum(1 for c in nb.cells if c["cell_type"] == "markdown")
    n_code = sum(1 for c in nb.cells if c["cell_type"] == "code")
    print(f"✔ {ipynb_path}")
    print(f"  Всего ячеек: {len(nb.cells)} (markdown: {n_md}, code: {n_code})")


if __name__ == "__main__":
    docx_to_ipynb(
        "input.docx",
        "Интерпретация_результатов.ipynb",
        title="Тема: Интерпретация результатов и анализ",
    )