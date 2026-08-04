import streamlit as st
import os
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image
import io
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas

st.title("PPT 이미지 투명 PDF 변환기")
st.write("PPTX 파일을 업로드하면 각 슬라이드의 이미지들을 투명도가 유지된 PDF 파일로 변환합니다.")

uploaded_file = st.file_uploader("파워포인트 파일(.pptx)을 업로드하세요", type=["pptx"])

def extract_images_with_position(shapes, image_list):
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            extract_images_with_position(shape.shapes, image_list)
        elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            image_list.append({
                'image': shape.image,
                'left': shape.left,
                'top': shape.top,
                'width': shape.width,
                'height': shape.height
            })

if uploaded_file is not None:
    with st.spinner("투명 PDF로 변환하는 중입니다..."):
        prs = Presentation(uploaded_file)
        slide_width = prs.slide_width
        slide_height = prs.slide_height
        
        pdf_buffer = io.BytesIO()
        
        # 가로 비율이 더 넓은 경우 가로형 PDF 페이지 크기 설정
        is_landscape = slide_width > slide_height
        page_size = landscape(letter) if is_landscape else letter
        page_w, page_h = page_size
        
        c = canvas.Canvas(pdf_buffer, pagesize=page_size)
        
        for slide in prs.slides:
            slide_images = []
            extract_images_with_position(slide.shapes, slide_images)
            
            if not slide_images:
                c.showPage()
                continue
            
            # 임시 투명 PNG 캔버스 생성용 고해상도 설정
            canvas_w = 1920
            canvas_h = int(1920 * (slide_height / slide_width))
            base_canvas = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))
            
            for item in slide_images:
                try:
                    img_stream = io.BytesIO(item['image'].blob)
                    pil_img = Image.open(img_stream).convert("RGBA")
                    
                    x = int(item['left'] / slide_width * canvas_w)
                    y = int(item['top'] / slide_height * canvas_h)
                    w = int(item['width'] / slide_width * canvas_w)
                    h = int(item['height'] / slide_height * canvas_h)
                    
                    resized_img = pil_img.resize((w, h), Image.Resampling.LANCZOS)
                    base_canvas.paste(resized_img, (x, y), resized_img)
                except Exception as e:
                    pass
            
            # 임시 이미지 파일로 저장 후 ReportLab PDF에 삽입 (투명도 유지)
            temp_img_path = "temp_slide_img.png"
            base_canvas.save(temp_img_path, "PNG")
            
            c.drawImage(temp_img_path, 0, 0, width=page_w, height=page_h, mask='auto')
            c.showPage()
            
            if os.path.exists(temp_img_path):
                os.remove(temp_img_path)
                
        c.save()
        pdf_buffer.seek(0)
        
        st.success("PDF 변환이 완료되었습니다!")
        st.download_button(
            label="투명 PDF 다운로드",
            data=pdf_buffer,
            file_name="converted_transparent.pdf",
            mime="application/pdf"
        )
