import os
import fitz  # PyMuPDF
import tkinter as tk
from PIL import Image, ImageTk
import pytesseract
import pandas as pd


# (선택) Tesseract 실행파일 경로 직접 지정 (Windows)
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# ------------------------------------
# 1) PDF -> PIL.Image 렌더링 함수
# ------------------------------------
def render_pdf_page_as_image(pdf_path, page_index=0, zoom_x=2.0, zoom_y=2.0):
    """
    PDF에서 page_index 페이지를 (zoom_x, zoom_y) 배율로 렌더링하여
    Pillow Image 객체로 반환.
    """
    doc = fitz.open(pdf_path)
    if page_index >= len(doc):
        doc.close()
        raise ValueError(f"{pdf_path}에는 page_index={page_index} 페이지가 없습니다.")

    page = doc[page_index]
    mat = fitz.Matrix(zoom_x, zoom_y)
    pix = page.get_pixmap(matrix=mat)
    doc.close()

    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return img


# ------------------------------------
# 2) 사용자 드래그로 테이블 좌표 선택
# ------------------------------------
def select_table_region(pdf_path, zoom_x=2.0, zoom_y=2.0):
    """
    첫 번째 PDF 파일을 열어, page=0을 (zoom_x, zoom_y)로 렌더링한 뒤
    Tkinter Canvas에서 사용자가 드래그로 사각형을 지정.

    반환:
      ( (left, top, right, bottom), (img_width, img_height) )
      즉, 선택된 영역(픽셀 좌표), 그리고 렌더링된 전체 이미지 크기.
    """
    # (1) PDF 렌더링 (첫 페이지)
    img = render_pdf_page_as_image(pdf_path, page_index=0, zoom_x=zoom_x, zoom_y=zoom_y)
    w, h = img.width, img.height

    # (2) Tkinter GUI
    root = tk.Tk()
    root.title("테이블 영역 선택")
    root.geometry(f"{w}x{h}")  # 윈도우 크기를 이미지와 동일하게

    canvas = tk.Canvas(root, width=w, height=h)
    canvas.pack()

    tk_img = ImageTk.PhotoImage(img)
    canvas.create_image(0, 0, anchor=tk.NW, image=tk_img)

    rect_id = None
    start_x, start_y = 0, 0

    # 최종 선택 영역
    selected_area = [0, 0, 0, 0]  # (left, top, right, bottom)

    def on_mouse_press(event):
        nonlocal start_x, start_y, rect_id
        start_x, start_y = event.x, event.y
        if rect_id is not None:
            canvas.delete(rect_id)
            rect_id = None

    def on_mouse_drag(event):
        nonlocal rect_id
        if rect_id is not None:
            canvas.delete(rect_id)
        rect_id = canvas.create_rectangle(
            start_x, start_y, event.x, event.y,
            outline="red", width=2, dash=(2, 2)
        )

    def on_mouse_release(event):
        nonlocal selected_area
        end_x, end_y = event.x, event.y
        left, right = sorted([start_x, end_x])
        top, bottom = sorted([start_y, end_y])
        selected_area = [left, top, right, bottom]

    def on_ok():
        root.destroy()

    canvas.bind("<ButtonPress-1>", on_mouse_press)
    canvas.bind("<B1-Motion>", on_mouse_drag)
    canvas.bind("<ButtonRelease-1>", on_mouse_release)

    btn_ok = tk.Button(root, text="OK", command=on_ok)
    btn_ok.pack()

    root.mainloop()

    return selected_area, (w, h)


# ------------------------------------
# 3) 지정 좌표로 OCR 수행
# ------------------------------------
def ocr_extract_text(pdf_path, region, zoom_x=2.0, zoom_y=2.0):
    """
    region: (left, top, right, bottom) 픽셀 좌표 (select_table_region()에서 얻은 것)
    PDF를 동일한 (zoom_x, zoom_y) 배율로 렌더링하고, 해당 픽셀 영역을 crop 후 OCR 수행.
    """
    img = render_pdf_page_as_image(pdf_path, page_index=0, zoom_x=zoom_x, zoom_y=zoom_y)
    w, h = img.width, img.height

    left, top, right, bottom = region
    # 혹시나 범위가 이미지 밖이면 보정
    left = max(0, min(left, w))
    right = max(0, min(right, w))
    top = max(0, min(top, h))
    bottom = max(0, min(bottom, h))

    if right - left < 1 or bottom - top < 1:
        return ""

    cropped_img = img.crop((left, top, right, bottom))

    # OCR
    text = pytesseract.image_to_string(cropped_img)
    return text.strip()


# ------------------------------------
# 4) 메인 흐름
# ------------------------------------
def process_pdfs_in_folder():
    # 사용자에게 폴더 경로 입력
    folder_path = input("PDF 폴더 경로를 입력하세요: ").strip()
    if not os.path.isdir(folder_path):
        print("폴더가 존재하지 않습니다.")
        return

    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
    if not pdf_files:
        print("해당 폴더에 PDF 파일이 없습니다.")
        return
    pdf_files.sort()  # 이름순 정렬

    # 첫 PDF에서 좌표 드래그
    first_pdf = os.path.join(folder_path, pdf_files[0])
    print(f"[INFO] 첫 번째 PDF({first_pdf})에서 테이블 영역을 선택하세요.")
    # 확대 배율 설정 (2.0 정도가 보통)
    zoom_x = zoom_y = 2.0

    selected_area, (img_w, img_h) = select_table_region(first_pdf, zoom_x, zoom_y)
    if selected_area == [0, 0, 0, 0]:
        print("[WARN] 영역을 제대로 지정하지 않았습니다. 종료합니다.")
        return

    # 모든 PDF에 대해 동일 배율로 OCR
    results = []
    for pdf_name in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_name)
        print(f"[INFO] OCR 처리 중: {pdf_path} ...")
        try:
            ocr_text = ocr_extract_text(pdf_path, selected_area, zoom_x, zoom_y)
        except Exception as e:
            print(f"[ERROR] OCR 실패 ({pdf_name}): {e}")
            ocr_text = ""

        results.append({
            "FileName": pdf_name,
            "ExtractedText": ocr_text
        })

    # 결과 저장
    df = pd.DataFrame(results)
    output_excel = os.path.join(folder_path, "extracted_data.xlsx")
    df.to_excel(output_excel, index=False)
    print(f"[INFO] OCR 결과가 '{output_excel}'에 저장되었습니다.")

    input("\n모든 작업이 완료되었습니다. Enter 키를 누르면 종료합니다...")


if __name__ == "__main__":
    process_pdfs_in_folder()
