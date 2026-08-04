import streamlit as st
import os
import io
import zipfile
from PIL import Image
from pptx import Presentation

st.title("PPT 개별 스티커 레이어 분리 추출기")
st.write("파워포인트 각 슬라이드의 여러 이미지를 개별 레이어(투명 PNG)로 분리하여 추출합니다.")

uploaded_file = st.file_uploader("PPTX 파일을 선택하세요", type=["pptx"])

if uploaded_file is not None:
    os.makedirs("temp", exist_ok=True)
    ppt_path = os.path.join("temp", uploaded_file.name)
    
    with open(ppt_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
        
    st.success("파일 업로드 완료! 개별 스티커 레이어를 추출 중입니다...")
    
    prs = Presentation(ppt_path)
    zip_buffer = io.BytesIO()
    
    total_elements = 0
    
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for slide_idx, slide in enumerate(prs.slides):
            elem_count = 0
            
            # 1. 일반 그림 개체 개별 추출
            for shape in slide.shapes:
                if shape.shape_type == 1:
                    try:
                        img_bytes = shape.image.blob
                        img = Image.open(io.BytesIO(img_bytes)).convert("RGBA")
                        
                        width = int(shape.width.inches * 96)
                        height = int(shape.height.inches * 96)
                        
                        img = img.resize((max(width, 1), max(height, 1)), Image.Resampling.LANCZOS)
                        
                        img_byte_arr = io.BytesIO()
                        img.save(img_byte_arr, format="PNG")
                        
                        elem_count += 1
                        total_elements += 1
                        filename = f"slide_{slide_idx + 1}_sticker_{elem_count}.png"
                        zip_file.writestr(filename, img_byte_arr.getvalue())
                    except Exception:
                        continue
                        
            # 2. 그림 개체가 없다면 XML 요소 기반으로 추가 탐색하여 개별 추출
            if elem_count == 0:
                for shape in slide.shapes:
                    element = shape.element
                    blips = element.xpath('.//a:blip')
                    for blip in blips:
                        embed_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                        if embed_id:
                            try:
                                image_part = slide.part.related_part(embed_id)
                                img = Image.open(io.BytesIO(image_part.blob)).convert("RGBA")
                                
                                width = int(shape.width.inches * 96)
                                height = int(shape.height.inches * 96)
                                
                                img = img.resize((max(width, 1), max(height, 1)), Image.Resampling.LANCZOS)
                                
                                img_byte_arr = io.BytesIO()
                                img.save(img_byte_arr, format="PNG")
                                
                                elem_count += 1
                                total_elements += 1
                                filename = f"slide_{slide_idx + 1}_sticker_{elem_count}.png"
                                zip_file.writestr(filename, img_byte_arr.getvalue())
                            except Exception:
                                continue

    if total_elements > 0:
        st.write(f"총 **{total_elements}개**의 개별 스티커 요소가 독립된 투명 레이어 파일로 추출되었습니다.")
        
        st.download_button(
            label="개별 스티커 레이어 일괄 다운로드 (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="separated_sticker_layers.zip",
            mime="application/zip"
        )
    else:
        st.warning("슬라이드 내에 추출 가능한 이미지 요소가 없습니다.")
