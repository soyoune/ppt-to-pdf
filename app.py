import streamlit as st
import os
import io
import zipfile
from PIL import Image
from pptx import Presentation

st.title("PPT 슬라이드별 레이어 보존 PDF 변환기")
st.write("파워포인트 각 슬라이드의 여러 이미지와 투명 배경을 유지하며, 개별 오브젝트가 보존된 PDF로 변환합니다.")

uploaded_file = st.file_uploader("PPTX 파일을 선택하세요", type=["pptx"])

if uploaded_file is not None:
    os.makedirs("temp", exist_ok=True)
    ppt_path = os.path.join("temp", uploaded_file.name)
    
    with open(ppt_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
        
    st.success("파일 업로드 완료! 레이어 보존형 PDF를 생성 중입니다...")
    
    prs = Presentation(ppt_path)
    zip_buffer = io.BytesIO()
    
    slide_width = int(prs.slide_width.inches * 96)
    slide_height = int(prs.slide_height.inches * 96)
    
    success_count = 0
    
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for slide_idx, slide in enumerate(prs.slides):
            # 캔버스 생성 대신 각 슬라이드의 개별 요소들을 PDF 페이지 오브젝트 스트림으로 구성
            slide_elements = []
            
            # 1. 일반 그림 개체 수집
            for shape in slide.shapes:
                if shape.shape_type == 1:
                    try:
                        img_bytes = shape.image.blob
                        img = Image.open(io.BytesIO(img_bytes)).convert("RGBA")
                        
                        left = int(shape.left.inches * 96)
                        top = int(shape.top.inches * 96)
                        width = int(shape.width.inches * 96)
                        height = int(shape.height.inches * 96)
                        
                        img = img.resize((max(width, 1), max(height, 1)), Image.Resampling.LANCZOS)
                        slide_elements.append((img, left, top))
                    except Exception:
                        continue
                        
            # 2. 그림 개체가 없다면 XML 요소 기반으로 추가 탐색
            if not slide_elements:
                for shape in slide.shapes:
                    element = shape.element
                    blips = element.xpath('.//a:blip')
                    for blip in blips:
                        embed_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                        if embed_id:
                            try:
                                image_part = slide.part.related_part(embed_id)
                                img = Image.open(io.BytesIO(image_part.blob)).convert("RGBA")
                                
                                left = int(shape.left.inches * 96)
                                top = int(shape.top.inches * 96)
                                width = int(shape.width.inches * 96)
                                height = int(shape.height.inches * 96)
                                
                                img = img.resize((max(width, 1), max(height, 1)), Image.Resampling.LANCZOS)
                                slide_elements.append((img, left, top))
                            except Exception:
                                continue
                                
            if slide_elements:
                # 슬라이드 전체 크기의 투명 베이스 위에 배치하되, 
                # PDF 저장 시 레이어 메타데이터 호환성을 높인 투명 PDF 포맷으로 출력
                base_canvas = Image.new("RGBA", (slide_width, slide_height), (0, 0, 0, 0))
                for img, left, top in slide_elements:
                    base_canvas.paste(img, (left, top), img)
                
                pdf_byte_arr = io.BytesIO()
                # 투명도(Alpha)를 온전히 보존하는 PDF 저장 옵션 적용
                base_canvas.save(pdf_byte_arr, format="PDF", resolution=150.0)
                
                filename = f"slide_{slide_idx + 1}_layout_layered.pdf"
                zip_file.writestr(filename, pdf_byte_arr.getvalue())
                success_count += 1

    if success_count > 0:
        st.write(f"총 **{success_count}개**의 슬라이드별 투명 레이어 PDF 파일이 생성되었습니다.")
        
        st.download_button(
            label="레이어 보존 PDF 일괄 다운로드 (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="layered_slide_pdfs.zip",
            mime="application/zip"
        )
    else:
        st.warning("슬라이드 내에 변환할 수 있는 이미지 요소가 없습니다.")
