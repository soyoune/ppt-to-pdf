import streamlit as st
import os
import io
from PIL import Image
from pptx import Presentation

st.title("PPT 이미지 투명 PDF 변환기")
st.write("파워포인트 각 슬라이드의 대표 이미지를 추출하여 하나의 투명 PDF로 변환해 드립니다.")

uploaded_file = st.file_uploader("PPTX 파일을 선택하세요", type=["pptx"])

if uploaded_file is not None:
    os.makedirs("temp", exist_ok=True)
    ppt_path = os.path.join("temp", uploaded_file.name)
    
    with open(ppt_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
        
    st.success("파일 업로드 완료! 슬라이드별 이미지를 수집 중입니다...")
    
    prs = Presentation(ppt_path)
    pil_images = []
    
    # 슬라이드별로 고유 이미지를 순서대로 수집
    for slide_idx, slide in enumerate(prs.slides):
        slide_img_bytes = None
        
        # 1. 일반 그림 개체 탐색
        for shape in slide.shapes:
            if shape.shape_type == 1:
                try:
                    slide_img_bytes = shape.image.blob
                    break
                except Exception:
                    continue
                    
        # 2. 그림 개체가 없다면 XML 요소에서 탐색
        if not slide_img_bytes:
            for shape in slide.shapes:
                element = shape.element
                blips = element.xpath('.//a:blip')
                for blip in blips:
                    embed_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                    if embed_id:
                        try:
                            image_part = slide.part.related_part(embed_id)
                            slide_img_bytes = image_part.blob
                            break
                        except Exception:
                            continue
                if slide_img_bytes:
                    break
                    
        # 해당 슬라이드에서 이미지를 찾았으면 PIL Image로 변환 후 리스트에 추가
        if slide_img_bytes:
            try:
                img = Image.open(io.BytesIO(slide_img_bytes))
                # RGBA 모드(투명도 포함)를 RGB로 변환 (PDF 저장을 위해 필요할 수 있음)
                if img.mode in ("RGBA", "P"):
                    img = img.convert("RGB")
                pil_images.append(img)
            except Exception:
                continue

    if pil_images:
        st.write(f"총 **{len(pil_images)}개**의 슬라이드별 고유 이미지를 매칭했습니다.")
        
        try:
            pdf_buffer = io.BytesIO()
            # 첫 번째 이미지를 기준으로 다중 페이지 PDF 저장
            pil_images[0].save(
                pdf_buffer, 
                format="PDF", 
                save_all=True, 
                append_images=pil_images[1:]
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
