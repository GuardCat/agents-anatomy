#!/usr/bin/env -S uv run
# /// script
# dependencies = [
#     "python-docx",
# ]
# ///

import sys
import os
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table

def parse_docx(file_path, output_dir="extracted_media"):
    try:
        doc = Document(file_path)
    except Exception as e:
        print(f"Ошибка чтения: {e}")
        return

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    print(f"# Анализ документа: {file_path}\n")

    img_counter = 1

    def process_drawing(drawing_element):
        """Извлекает картинку, сохраняет ее и возвращает Markdown-тег"""
        nonlocal img_counter
        try:
            blips = drawing_element.xpath('.//a:blip')
            if not blips:
                return ""
            
            embed_id = blips[0].get(qn('r:embed'))
            if not embed_id or embed_id not in doc.part.related_parts:
                return ""
            
            image_part = doc.part.related_parts[embed_id]
            ext = image_part.content_type.split('/')[-1].replace('jpeg', 'jpg')
            fname = f"img_{img_counter}.{ext}"
            fpath = os.path.join(output_dir, fname)
            
            with open(fpath, "wb") as f:
                f.write(image_part.blob)
            
            extents = drawing_element.xpath('.//wp:extent')
            dim_str = ""
            if extents:
                cx = int(extents[0].get('cx', 0)) / 360000
                cy = int(extents[0].get('cy', 0)) / 360000
                dim_str = f" ({cx:.2f}x{cy:.2f} см)"
            
            placeholder = f"![Изображение {img_counter}{dim_str}]({fpath})"
            img_counter += 1
            return placeholder
        except Exception as e:
            return f"[Ошибка извлечения медиа: {e}]"

    # Обходим все блочные элементы документа строго в порядке их следования
    for block in doc.element.body:
        
        # ЕСЛИ ЭТО АБЗАЦ
        if block.tag.endswith('p'):
            para_text = ""
            for child in block:
                if child.tag.endswith('r'):
                    for run_child in child:
                        if run_child.tag.endswith('t') and run_child.text:
                            para_text += run_child.text
                        elif run_child.tag.endswith('drawing'):
                            para_text += process_drawing(run_child)
                elif child.tag.endswith('hyperlink'):
                    rel_id = child.get(qn('r:id'))
                    url = doc.part.rels[rel_id].target_ref if rel_id in doc.part.rels else ""
                    link_text = "".join([t.text for t in child.xpath('.//w:t') if t.text])
                    para_text += f"[{link_text}]({url})" if url else link_text

            full_text = para_text.strip()
            if full_text:
                print(full_text + "\n")

        # ЕСЛИ ЭТО ТАБЛИЦА
        elif block.tag.endswith('tbl'):
            table = Table(block, doc._body)
            for r_idx, row in enumerate(table.rows):
                # Извлекаем текст из ячеек, убирая переносы строк для Markdown-формата
                cells = [c.text.replace('\n', ' ').strip() for c in row.cells]
                print("| " + " | ".join(cells) + " |")
                
                # Добавляем разделитель Markdown-таблицы после первой строки
                if r_idx == 0:
                    print("|" + "|".join(["---"] * len(cells)) + "|")
            print() # Пустая строка после таблицы для корректного рендеринга Markdown

if __name__ == "__main__":
    if len(sys.argv) > 1:
        path = sys.argv[1]
        out = sys.argv[2] if len(sys.argv) > 2 else f"{os.path.splitext(os.path.basename(path))[0]}_media"
        parse_docx(path, out)
    else:
        print("Использование: docx_reader <путь_к_файлу> [папка_сохранения]")