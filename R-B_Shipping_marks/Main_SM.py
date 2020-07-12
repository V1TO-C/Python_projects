import os
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
from docxcompose.composer import Composer


def main():
    input_anzahl = input("Anzahl: ")
    input_pages = int(input("Number of pages: "))
    input_load_doc_name = "Shipping_M"  # potom upravit na input()
    input_created_doc_name = input("Name of created file: ") + ".docx"

    make_improved_document(input_load_doc_name, input_anzahl)

    compose_doc(input_created_doc_name, "temporary_copy.docx", input_pages)

    change_style_paragraphs_tables(input_created_doc_name, input_pages, input_created_doc_name)


def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None


def make_improved_document(load_file, anzahl):
    doc = Document(load_file)
    # improved_doc = Document()
    # art_num = doc.paragraphs[0]
    # art_num_str = f"1-{art_num}"

    table = doc.tables[0]
    table.cell(4, 1).text = anzahl

    # table_art_num = table.cell(6,1).text = f"1-{pages}"

    delete_paragraph(doc.paragraphs[1])
    doc.add_paragraph("")

    doc.save("temporary_copy.docx")


def compose_doc(created, copy, pages):
    doc_main = Document()
    doc_main.save("temporary_main.docx")
    master = Document("temporary_main.docx")
    composer = Composer(master)
    doc = Document(copy)
    for i in range(pages):
        composer.append(doc)
    composer.save(created)
    os.remove("temporary_main.docx")
    os.remove("temporary_copy.docx")


def change_style_paragraphs_tables(load_doc, pages, saved_doc):
    composed_doc = Document(load_doc)
    style = composed_doc.styles["Normal"]
    font = style.font
    font.name = 'Arial'
    font.size = Pt(9)

    delete_counter = len(composed_doc.paragraphs)  # na jedné kopii je 5 paragr., rozlišit sudý/lichý počet stran
    if delete_counter % 2 == 0:                    # umazat čtyři paragr. za tabulkou pro rozložení 2x na list
        for par in range(1, len(composed_doc.paragraphs), 10):
            for i in range(1, 4):
                delete_paragraph(composed_doc.paragraphs[delete_counter - i])
            delete_counter -= 10
    else:
        for par in range(1, len(composed_doc.paragraphs), 10):
            for i in range(1, 4):
                delete_paragraph(composed_doc.paragraphs[delete_counter - i - 5])
            delete_counter -= 10

    table_count = 1
    for tab in range(pages):
        table2 = composed_doc.tables[tab]
        table2.cell(6, 1).text = f"{table_count}-{pages}"  # úprava buněk s počtem stran
        table_count += 1
        for c in range(len(table2.columns)):  # úprava formátu textu buněk
            for r in range(len(table2.rows)):
                copy = table2.cell(r, c).text
                table2.cell(r, c).style = composed_doc.styles['Normal']
                table2.cell(r, c).text = copy

    for table in composed_doc.tables:  # zarovnání první buňky nastřed adresy
        table.cell(0, 0).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    composed_doc.save(saved_doc)


if __name__ == '__main__':
    main()
