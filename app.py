import streamlit as st
import os
import zipfile
from pptx import Presentation
import io

st.title("PPT 이미지 투명 PDF/PNG 변환기")
st.write("파워포인트 파일을 업로드하면 각 슬라이드의 투명도가 유지된 이미지 요소들을 추출해 드립니다.")

uploaded_file = st.file_uploader("PPTX 파일을 선택하세요", type=["pptx"])

if uploaded_file is not None:
    # 임시 디렉토리 생성
    os.makedirs("temp", exist_ok=True)
    ppt_path = os.path.join("temp", uploaded_file.name)
    
    with open(ppt_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
        
    st.success("파일 업로드 완료! 요소를 추출 중입니다...")
    
    prs = Presentation(ppt_path)
    extracted_images = []
    
    # 슬라이드별로 이미지 개수를 독립적으로 카운트하기 위한 변수
    for slide_idx, slide in enumerate(prs.slides):
        slide_image_count = 0  # 슬라이드 바뀔 때마다 0부터 카운트
        for shape in slide.shapes:
            if shape.shape_type == 1:  # 그림(Picture) 타입
                image = shape.image
                image_bytes = image.blob
                image_ext = image.ext
                
                slide_image_count += 1
                image_filename = f"slide_{slide_idx + 1}_img_{slide_image_count}.{image_ext}"
                extracted_images.append((image_filename, image_bytes))
                
    if extracted_images:
        st.write(f"총 **{len(extracted_images)}개**의 투명 이미지 요소를 발견했습니다.")
        
        # ZIP 파일로 묶어서 다운로드 제공
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
