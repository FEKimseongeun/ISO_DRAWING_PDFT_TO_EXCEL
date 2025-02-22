import fitz  # PyMuPDF
import pandas as pd
import os

def get_corrected_rect(page, rect):
    """
    PDF가 90도, 180도, 270도 회전된 경우 좌표를 변환하는 함수.
    """
    rotation = page.rotation
    page_width = page.rect.width
    page_height = page.rect.height

    if rotation == 90:
        return fitz.Rect(rect.y0, page_width - rect.x1, rect.y1, page_width - rect.x0)
    elif rotation == 180:
        return fitz.Rect(page_width - rect.x1, page_height - rect.y1, page_width - rect.x0, page_height - rect.y0)
    elif rotation == 270:
        return fitz.Rect(page_height - rect.y1, rect.x0, page_height - rect.y0, rect.x1)
    return rect  # 0도 회전이면 원래 좌표 반환

def find_text_position(doc, search_text):
    """
    주어진 doc(PDF)에서 특정 키워드의 위치(좌표) 목록을 찾아 반환.
    """
    positions = []
    for page_index, page in enumerate(doc):
        bboxes = page.search_for(search_text)
        for bbox in bboxes:
            corrected_bbox = get_corrected_rect(page, bbox)  # 회전 변환 적용
            positions.append((page_index, corrected_bbox))
    return positions

def extract_text_in_region(page, rect):
    """
    주어진 page의 특정 영역(rect)에서 텍스트를 추출 (회전 보정 포함).
    """
    corrected_rect = get_corrected_rect(page, rect)  # 회전 보정
    text = page.get_text("text", clip=corrected_rect)
    return text.strip() if text else ""

def extract_data_from_pdf(pdf_path):
    """
    PDF에서 표 및 ISO_NO, REV.NO 데이터를 추출하는 함수.
    """
    # 좌표 세트 1
    table_rect_set1 = fitz.Rect(39.89999771118164, 760.1726684570312, 557.47998046875, 820.8424682617188)
    iso_no_rect_set1 = fitz.Rect(960.6903686523438, 811.55810546875, 1500.6903686523438, 816.8818969726562)
    rev_no_rect_set1 = fitz.Rect(1156.469970703125, 812.6487426757812, 1172.7728271484375, 820.2630004882812)

    # 좌표 세트 2
    table_rect_set2 = fitz.Rect(10.4000244140625, 535.4000244140625, 389.02935546875, 567.79730224609375)
    iso_no_rect_set2 = fitz.Rect(687.4000244140625, 565.4000244140625, 780.02935546875, 580.79730224609375)
    rev_no_rect_set2 = fitz.Rect(820.4000244140625, 575.4000244140625, 829.02935546875, 579.79730224609375)

    try:
        with fitz.open(pdf_path) as doc:
            # ✅ 키워드 좌표 찾기
            try:
                nps_positions = find_text_position(doc, "NPS")
                toxic_positions = find_text_position(doc, "TOXIC")
                kosha_positions = find_text_position(doc, "KOSHA")
                iso_positions = find_text_position(doc, "ISO DWG. NO.")
                rev_positions = find_text_position(doc, "REV. NO")
            except Exception as e:
                print(f"⚠️ 키워드 좌표 검색 중 오류 발생: {e}")
                return None

            if not (nps_positions and toxic_positions and kosha_positions and iso_positions and rev_positions):
                print(f"❌ {pdf_path}에서 필요한 키워드를 찾을 수 없습니다.")
                return None

            # 예시로 첫 번째 결과만 사용 (기존 로직 그대로)
            nps_page_idx, nps_rect = nps_positions[0]
            iso_page_idx, iso_rect = iso_positions[0]
            rev_page_idx, rev_rect = rev_positions[0]

            # table, iso, rev는 모두 NPS가 있는 페이지라고 가정(또는 필요 시 iso_page_idx, rev_page_idx로 각기 불러도 됨)
            nps_page = doc[nps_page_idx]
            iso_page = doc[iso_page_idx]
            rev_page = doc[rev_page_idx]

            # -----------------------------
            # 1) table_rect 첫 번째 세트 시도
            # -----------------------------
            table_text = extract_text_in_region(nps_page, table_rect_set1)

            if table_text.strip():
                # 첫 번째 세트 성공 -> iso_no_rect, rev_no_rect도 첫 번째 세트 사용
                iso_no_text = extract_text_in_region(iso_page, iso_no_rect_set1)
                rev_no_text = extract_text_in_region(rev_page, rev_no_rect_set1)
            else:
                # -----------------------------
                # 2) table_rect 두 번째 세트 시도
                # -----------------------------
                table_text = extract_text_in_region(nps_page, table_rect_set2)
                if not table_text.strip():
                    # 두 번째 세트도 실패하면 데이터 없음 처리
                    print(f"❌ {pdf_path}에서 표 추출 실패 (두 좌표 세트 모두 텍스트 없음)")
                    return None
                # 두 번째 세트 성공 -> iso_no_rect, rev_no_rect도 두 번째 세트 사용
                iso_no_text = extract_text_in_region(iso_page, iso_no_rect_set2)
                rev_no_text = extract_text_in_region(rev_page, rev_no_rect_set2)

            # ✅ 표 데이터를 파싱하여 DataFrame 생성
            #    - 이 부분은 PDF 구조에 따라 잘게 split되지 않을 수 있으니 적절한 후처리 필요
            columns = [
                "NPS", "SPEC", "OPER.PRESS", "OPER.TEMP", "DESIGN.PRESS", "DESIGN.TEMP",
                "TEST.PRESS", "MEDIUM", "INSULATION.TYPE", "INSULATION.THK", "TRACING.TEMP",
                "PAUT", "UT", "PT", "MT", "PAINTCODE", "PWHT", "STEAMOUT", "TOXIC"
            ]
            # 표를 줄단위로 나누고, 공백으로 split
            rows = [line.split() for line in table_text.split("\n") if line.strip()]

            # 열 개수가 일치하지 않으면(혹은 표 형태가 다르면) 에러가 날 수 있으니 예외 처리 필요
            # 여기서는 예시로 columns 개수와 동일하다고 가정
            if not rows:
                print(f"❌ {pdf_path}에서 표 데이터가 비어 있습니다.")
                return None
            if any(len(r) != len(columns) for r in rows):
                print(f"⚠️ {pdf_path}에서 표 구조가 예상과 다릅니다. 데이터 길이 불일치.")
                # 필요 시 여기서 후처리...
                return None

            df = pd.DataFrame(rows, columns=columns)

            # ISO_NO, REV.NO 추가 컬럼
            df["ISO_NO"] = iso_no_text
            df["REV.NO"] = rev_no_text

            return df

    except Exception as e:
        print(f"⚠️ PDF 처리 중 오류 발생: {e} (파일 스킵)")
        return None

def process_pdfs_in_folder():
    """
    폴더 내 모든 PDF 파일을 처리하고 CSV로 저장하는 함수.
    """
    folder_path = input("📂 PDF 폴더 경로를 입력하세요: ").strip()
    if not os.path.exists(folder_path):
        print("❌ 폴더가 존재하지 않습니다.")
        return

    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
    if not pdf_files:
        print("❌ 폴더에 PDF가 없습니다.")
        return

    all_data = []
    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        print(f"▶ 처리 중: {pdf_path}")
        df = extract_data_from_pdf(pdf_path)
        if df is not None and not df.empty:
            all_data.append(df)

    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        output_csv = os.path.join(folder_path, "extracted_data.csv")
        final_df.to_csv(output_csv, index=False)
        print(f"✅ CSV 저장 완료: {output_csv}")
    else:
        print("❌ 추출할 데이터가 없습니다.")


    # 작업 완료 후 콘솔 창이 바로 닫히지 않도록 대기
    input("\n모든 작업이 완료되었습니다. 콘솔을 종료하려면 Enter 키를 누르세요...")
# 실행 예시
if __name__ == "__main__":
    process_pdfs_in_folder()
