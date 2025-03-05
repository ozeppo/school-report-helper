import fitz  # PyMuPDF
import os
import re


def extract_name_from_page(page):
    """Ekstrahuje nazwisko i imię ucznia z pierwszej strony segmentu."""
    text = page.get_text("text")
    match = re.search(r"Nazwisko i imię\s*([A-ZŁŚĆŻŹŃÓĘĄ][a-złśćżźńóęą]+\s+[A-ZŁŚĆŻŹŃÓĘĄ][a-złśćżźńóęą]+)", text,
                      re.IGNORECASE)
    return match.group(1) if match else None


def extract_group_from_page(page):
    """Ekstrahuje pełną nazwę grupy ucznia z pierwszej strony segmentu."""
    text = page.get_text("text")
    match = re.search(r"Grupa\s*(.+?)\s+Od", text, re.IGNORECASE)
    return match.group(1).strip() if match else "Nieznana_grupa"


def split_pdf(input_pdf, output_base):
    doc = fitz.open(input_pdf)
    filtered_pages = [page for page in doc if page.get_text("text").strip()]

    if not filtered_pages:
        return  # Jeśli cały dokument jest pusty, nie przetwarzamy go

    segments = []  # Lista przechowująca zakresy stron dla każdego ucznia
    current_start = None  # Indeks początku segmentu
    group_name = "Nieznana_grupa"

    for i, page in enumerate(filtered_pages):
        name = extract_name_from_page(page)
        if name:
            if current_start is not None:
                segments.append((current_start, i - 1))
            current_start = i  # Nowy segment zaczyna się tutaj
            group_name = extract_group_from_page(page)

    if current_start is not None:
        segments.append((current_start, len(filtered_pages) - 1))

    group_folder = os.path.join(output_base, group_name)
    os.makedirs(group_folder, exist_ok=True)

    for start, end in segments:
        name = extract_name_from_page(filtered_pages[start]) or f"Nieznany_uczen_{start}"
        output_path = os.path.join(group_folder, f"{name}.pdf")

        new_pdf = fitz.open()
        for i in range(start, end + 1):
            new_pdf.insert_pdf(doc, from_page=doc.index(filtered_pages[i]), to_page=doc.index(filtered_pages[i]))

        new_pdf.save(output_path)
        new_pdf.close()
        print(f"Zapisano: {output_path}")


def process_folder(input_folder):
    output_base = os.path.join(input_folder, "output")
    os.makedirs(output_base, exist_ok=True)

    for filename in os.listdir(input_folder):
        if filename.lower().endswith(".pdf"):
            input_pdf = os.path.join(input_folder, filename)
            split_pdf(input_pdf, output_base)


if __name__ == "__main__":
    input_folder = "documents/input"  # Podmień na właściwą ścieżkę
    process_folder(input_folder)
