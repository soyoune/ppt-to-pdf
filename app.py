import streamlit as st
import os
import zipfile
from pptx import Presentation
import io

st.title("PPT 이미지 투명 PDF/PNG 변환기")
st.write("파워포인트 파일을 업로드하면 각 슬라이드의 투명도가 유지된 이미지 요소들을 추출해 드립니다.")

uploaded_file = st.file_uploader("PPTX 파일을 선택하세요", type=["pptx"])

if uploaded_file is not None:
    os.makedirs("temp", exist_ok=True)
    ppt_path = os.path.join("temp", uploaded_file.name)
    
    with open(ppt_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
        
    st.success("파일 업로드 완료! 요소를 추출 중입니다...")
    
    prs = Presentation(ppt_path)
    extracted_images = []
    
    for slide_idx, slide in enumerate(prs.slides):
        slide_image_count = 0
        
        # 1. 일반 그림 개체 탐색
        for shape in slide.shapes:
            if shape.has_image:
                try:
                    image = shape.image
                    image_bytes = image.blob
                    image_ext = image.ext
                    
                    slide_image_count += 1
                    image_filename = f"slide_{slide_idx + 1}_img_{slide_image_count}.{image_ext}"
                    extracted_images.append((image_filename, image_bytes))
                except Exception:
                    continue
                    
        # 2. 슬라이드 배경에 이미지가 있는 경우 탐색
        try:
            background = slide.background
            if background and background.fill and background.fill.type:
                # 배경 이미지 블롭 추출 시도
                bg_part = background.fill._element
                # XML에서 blip 요소(이미지 데이터)가 있는지 확인
                blips = bg_part.xpath('.//a:blip')
                for blip in blips:
                    embed_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                    if embed_id:
                        image_part = slide.part.related_part(embed_id)
                        image_bytes = image_part.blob
                        # 확장자 결정 (기본 png 또는 파트에서 유추)
                        image_ext = image_part.content_type.split('/')[-1]
                        if image_ext == 'jpeg':
                            image_ext = 'jpg'
                            
                        slide_image_count += 1
                        image_filename = f"slide_{slide_idx + 1}_bg_{slide_image_count}.{image_ext}"
                        extracted_images.append((image_filename, image_bytes))
        except Exception:
            pass

    if extracted_images:
        st.write(f"총 **{len(extracted_images)}개**의 투명 이미지 요소를 발견했습니다.")
        
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for filename, img_bytes in extracted_images:
                zip_file.writestr(filename, img_bytes)
                
        st.download_button(
            label="투명 이미지 일괄 다운로드 (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="transparent_images.zip",
            mime="application/zip"
        )
    else:
        st.warning("슬라이드 내에 추출 가능한 이미지 요소가 없습니다.")
