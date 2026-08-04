import streamlit as st
import os
import io
from PIL import Image
from pptx import Presentation

st.title("PPT 이미지 투명 PDF 변환기")
st.write("파워포인트 각 슬라이드의 여러 이미지 요소를 하나로 합성하여 투명 PDF로 변환해 드립니다.")

uploaded_file = st.file_uploader("PPTX 파일을 선택하세요", type=["pptx"])

if uploaded_file is not None:
    os.makedirs("temp", exist_ok=True)
    ppt_path = os.path.join("temp", uploaded_file.name)
    
    with open(ppt_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
        
    st.success("파일 업로드 완료! 슬라이드별 이미지를 합성 중입니다...")
    
    prs = Presentation(ppt_path)
    slide_pil_images = []
    
    # 파워포인트 기본 슬라이드 크기 (16:9 기준, 픽셀 단위로 환산)
    slide_width = int(prs.slide_width.inches * 96)
    slide_height = int(prs.slide_height.inches * 96)
    
    for slide in prs.slides:
        # 투명 배경을 지원하는 빈 캔버스 생성 (RGBA 모드)
        base_canvas = Image.new("RGBA", (slide_width, slide_height), (255, 255, 255, 0))
        image_found = False
        
        # 1. 일반 그림 개체 합성
        for shape in slide.shapes:
            if shape.shape_type == 1:
                try:
                    img_bytes = shape.image.blob
                    img = Image.open(io.BytesIO(img_bytes)).convert("RGBA")
                    
                    # PPT 내 위치 및 크기 계산 (인치 또는 파이썬-pptx 단위를 픽셀로 변환)
                    left = int(shape.left.inches * 96)
                    top = int(shape.top.inches * 96)
                    width = int(shape.width.inches * 96)
                    height = int(shape.height.inches * 96)
                    
                    # 이미지 크기 조절
                    img = img.resize((max(width, 1), max(height, 1)), Image.Resampling.LANCZOS)
                    
                    # 캔버스 위에 해당 위치에 맞춰 덮어쓰기 (알파 채널 활용)
                    base_canvas.paste(img, (left, top), img)
                    image_found = True
                except Exception:
                    continue
                    
        # 2. 그림 개체가 없다면 XML 요소 기반으로 추가 탐색
        if not image_found:
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
                            base_canvas.paste(img, (left, top), img)
                            image_found = True
                        except Exception:
                            continue
                            
        # 이미지가 하나라도 발견된 슬라이드라면 PDF 페이지용 이미지 리스트에 추가
        if image_found:
            # PDF 저장을 위해 투명도가 포함된 경우 RGB로 변환 (필요시)
            rgb_canvas = Image.new("RGB", base_canvas.size, (255, 255, 255))
            rgb_canvas.paste(base_canvas, mask=base_canvas.split()[3]) # 알파 채널 마스크 적용
            slide_pil_images.append(rgb_canvas)

    if slide_pil_images:
        st.write(f"총 **{len(slide_pil_images)}개**의 슬라이드가 성공적으로 합성되었습니다.")
        
        try:
            pdf_buffer = io.BytesIO()
            slide_pil_images[0].save(
                pdf_buffer, 
                format="PDF", 
                save_all=True, 
                append_images=slide_pil_images[1:]
            )
            
            st.download_button(
                label="투명 PDF 다운로드",
                data=pdf_buffer.getvalue(),
                file_name="converted_slides.pdf",
                mime="application/pdf"
            )
        except Exception as e:
            st.error(f"PDF 변환 중 오류가 발생했습니다: {e}")
    else:
        st.warning("슬라이드 내에 변환할 수 있는 이미지 요소가 없습니다.")
