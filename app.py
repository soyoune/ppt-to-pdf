import streamlit as st
import os
import io
import zipfile
from PIL import Image
from pptx import Presentation
import base64

st.title("PPT 슬라이드별 레이어 보존 SVG 변환기")
st.write("파워포인트 각 슬라이드의 여러 이미지를 일러스트레이터에서 개별 수정 가능한 투명 SVG 레이아웃으로 변환합니다.")

uploaded_file = st.file_uploader("PPTX 파일을 선택하세요", type=["pptx"])

if uploaded_file is not None:
    os.makedirs("temp", exist_ok=True)
    ppt_path = os.path.join("temp", uploaded_file.name)
    
    with open(ppt_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
        
    st.success("파일 업로드 완료! 레이어 보존형 SVG 파일을 생성 중입니다...")
    
    prs = Presentation(ppt_path)
    zip_buffer = io.BytesIO()
    
    slide_width = int(prs.slide_width.inches * 96)
    slide_height = int(prs.slide_height.inches * 96)
    
    success_count = 0
    
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for slide_idx, slide in enumerate(prs.slides):
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
                        
                        # PNG를 Base64로 인코딩하여 SVG 내부에 독립된 오브젝트로 삽입
                        buffered = io.BytesIO()
                        img.save(buffered, format="PNG")
                        img_base64 = base64.b64encode(buffered.getvalue()).decode("utf-8")
                        
                        slide_elements.append((img_base64, left, top, width, height))
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
                                
                                buffered = io.BytesIO()
                                img.save(buffered, format="PNG")
                                img_base64 = base64.b64encode(buffered.getvalue()).decode("utf-8")
                                
                                slide_elements.append((img_base64, left, top, width, height))
                            except Exception:
                                continue
                                
            if slide_elements:
                # SVG XML 구조 생성 (슬라이드 크기 내에 각 이미지가 독립된 <image> 태그(개별 레이어)로 배치됨)
                svg_content = f'''<svg width="{slide_width}" height="{slide_height}" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">
'''
                for img_b64, left, top, width, height in slide_elements:
                    svg_content += f'  <image x="{left}" y="{top}" width="{width}" height="{height}" href="data:image/png;base64,{img_b64}"/>\n'
                svg_content += '</svg>'
                
                filename = f"slide_{slide_idx + 1}_layered_layout.svg"
                zip_file.writestr(filename, svg_content.encode("utf-8"))
                success_count += 1

    if success_count > 0:
        st.write(f"총 **{success_count}개**의 슬라이드별 투명 SVG 레이어 파일이 생성되었습니다.")
        
        st.download_button(
            label="투명 SVG 레이어 일괄 다운로드 (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="layered_slide_svgs.zip",
            mime="application/zip"
        )
    else:
        st.warning("슬라이드 내에 변환할 수 있는 이미지 요소가 없습니다.")
