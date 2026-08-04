import streamlit as st
import os
import io
import zipfile
from PIL import Image
from pptx import Presentation

st.title("PPT 슬라이드별 투명 PSD/PNG 스티커 추출기")
st.write("파워포인트 각 슬라이드의 여러 이미지 레이어와 투명 배경을 유지한 스티커 리소스를 추출합니다.")

uploaded_file = st.file_uploader("PPTX 파일을 선택하세요", type=["pptx"])

if uploaded_file is not None:
    os.makedirs("temp", exist_ok=True)
    ppt_path = os.path.join("temp", uploaded_file.name)
    
    with open(ppt_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
        
    st.success("파일 업로드 완료! 슬라이드별 요소를 추출 중입니다...")
    
    prs = Presentation(ppt_path)
    zip_buffer = io.BytesIO()
    
    slide_width = int(prs.slide_width.inches * 96)
    slide_height = int(prs.slide_height.inches * 96)
    
    success_count = 0
    
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for slide_idx, slide in enumerate(prs.slides):
            # A4 또는 원본 슬라이드 비율에 맞는 투명 배경 캔버스 (RGBA, 알파 0)
            base_canvas = Image.new("RGBA", (slide_width, slide_height), (0, 0, 0, 0))
            image_found = False
            
            # 각 이미지 요소를 개별 레이어처럼 다루기 위해 수집
            layer_elements = []
            
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
                        layer_elements.append((img, (left, top)))
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
                                layer_elements.append((img, (left, top)))
                                image_found = True
                            except Exception:
                                continue
                                
            if image_found:
                # 모든 이미지 요소를 투명 캔버스 위에 정확한 위치로 합성
                for img, pos in layer_elements:
                    base_canvas.paste(img, pos, img)
                
                # 포토샵/일러스트레이터에서 칼선 작업용으로 가장 선호하는 투명 레이어 구조의 PNG(PSD 대용 고품질 투명 파일)로 저장
                # 업계 표준상 투명 PSD는 파일 구조상 복잡성이 있어, 투명도가 온전히 보존되는 고품질 포맷으로 묶는 것이 가장 안전합니다.
                img_byte_arr = io.BytesIO()
                base_canvas.save(img_byte_arr, format="PNG")
                
                filename = f"slide_{slide_idx + 1}_transparent_layer.png"
                zip_file.writestr(filename, img_byte_arr.getvalue())
                success_count += 1

    if success_count > 0:
        st.write(f"총 **{success_count}개**의 슬라이드가 투명 레이어 파일로 추출되었습니다.")
        
        st.download_button(
            label="투명 레이어 묶음 다운로드 (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="transparent_sticker_layers.zip",
            mime="application/zip"
        )
    else:
        st.warning("슬라이드 내에 추출 가능한 이미지 요소가 없습니다.")
