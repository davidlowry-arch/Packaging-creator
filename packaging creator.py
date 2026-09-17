import os
import re
import tempfile
import pandas as pd
from PIL import Image as PILImage
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, Image, Frame, KeepInFrame, HRFlowable, Table, TableStyle, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from datetime import datetime

# --- Configuration & Constants ---
LABEL_WIDTH = 63 * mm
# --- Configuration & Constants ---
COLS = 3  
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")

# Set width to exactly 1/3 of A4 width
LABEL_WIDTH = A4[0] / COLS  

def get_layout_geometry(rows=3):
    """Calculates label height for exact edge-to-edge page tiling."""
    # Set height to exactly 1/4 or 1/3 of A4 height
    label_height = A4[1] / float(rows)

    # Zero out all margins and gutters to perfectly tile the page
    margin_x = 0.0
    gutter_x = 0.0
    margin_y = 0.0
    gutter_y = 0.0

    return label_height, margin_x, gutter_x, margin_y, gutter_y

LABEL_HEIGHT, MARGIN_X, GUTTER_X, MARGIN_Y, GUTTER_Y = get_layout_geometry(3)

pdfmetrics.registerFont(TTFont('Arial', r'C:\Windows\Fonts\arial.ttf'))
pdfmetrics.registerFont(TTFont('Arial-Bold', r'C:\Windows\Fonts\arialbd.ttf'))
pdfmetrics.registerFont(TTFont('Times-Bold', r'C:\Windows\Fonts\timesbd.ttf'))
pdfmetrics.registerFont(TTFont('Rockwell-Bold', r'C:\Windows\Fonts\rockb.ttf'))

# Cultural / Calligraphic Fonts
pdfmetrics.registerFont(TTFont('Amiri-Bold', os.path.join(ASSETS_DIR, 'Amiri-Bold.ttf')))
pdfmetrics.registerFont(TTFont('ElMessiri-Bold', os.path.join(ASSETS_DIR, 'ElMessiri-Bold.ttf')))
pdfmetrics.registerFont(TTFont('Cinzel-Bold', os.path.join(ASSETS_DIR, 'CinzelDecorative-Bold.ttf')))
pdfmetrics.registerFont(TTFont('Noto-Bold', os.path.join(ASSETS_DIR, 'NotoSerif-Bold.ttf')))

RESOURCE_FONTS = {
    1: 'Rockwell-Bold',
    2: 'Amiri-Bold',
    3: 'ElMessiri-Bold',
    4: 'Cinzel-Bold',
    5: 'Noto-Bold'
}

styles = getSampleStyleSheet()
style_header = ParagraphStyle('Header', parent=styles['Normal'], fontName='Arial', fontSize=8, alignment=TA_CENTER, spaceAfter=6)
style_desc = ParagraphStyle('Desc', parent=styles['Normal'], fontSize=8, alignment=TA_CENTER, spaceAfter=2)
style_copy = ParagraphStyle('Copy', parent=styles['Normal'], fontName='Arial', fontSize=5, textColor='grey', alignment=TA_CENTER, spaceAfter=6)


def get_resized_image_path(img_path):
    if not os.path.exists(img_path):
        return img_path
        
    directory, filename = os.path.split(img_path)
    name, ext = os.path.splitext(filename)
    
    cache_dir = os.path.join(directory, "resized_cache")
    if not os.path.exists(cache_dir):
        os.makedirs(cache_dir)
        
    cached_path = os.path.join(cache_dir, f"{name}_450px{ext}")
    if os.path.exists(cached_path):
        return cached_path
        
    try:
        with PILImage.open(img_path) as img:
            wpercent = (450 / float(img.width))
            hsize = int((float(img.height) * float(wpercent)))
            resample_filter = getattr(PILImage, 'Resampling', PILImage).LANCZOS
            img = img.resize((450, hsize), resample_filter)
            
            if ext.lower() in ['.jpg', '.jpeg'] and img.mode in ('RGBA', 'P', 'LA'):
                img = img.convert('RGB')
                
            img.save(cached_path)
        return cached_path
    except Exception as e:
        print(f"Warning: Could not resize {filename}. Error: {e}")
        return img_path


def build_label_story(resources, include_images, title_size, image_size_mm, scale_factor=1.0, extra_gap=0, cc_shift=0):
    desc_font_size = 9 * scale_factor
    copy_font_size = 8 * scale_factor  
    
    style_desc_scaled = ParagraphStyle(
        'DescScaled', parent=styles['Normal'], 
        fontSize=desc_font_size, leading=desc_font_size * 1.2, 
        alignment=TA_CENTER, spaceAfter=2 * scale_factor
    )
    style_copy_scaled = ParagraphStyle(
        'CopyScaled', parent=styles['Normal'], fontName='Arial', 
        fontSize=copy_font_size, leading=copy_font_size * 1.2, 
        textColor='grey', alignment=TA_CENTER, spaceAfter=6 * scale_factor
    )
    
    story = []
    
    for i, (_, row) in enumerate(resources.iterrows()):
        hide_divider = row.get('HideDivider', False) if 'HideDivider' in resources.columns else False
        
        if i > 0 and not hide_divider:
            if extra_gap > 0: story.append(Spacer(1, extra_gap / 2))
            story.append(HRFlowable(
                width="50%", 
                thickness=0.5 * scale_factor, 
                color=colors.black, 
                spaceBefore=2 * scale_factor, 
                spaceAfter=2 * scale_factor, 
                hAlign='CENTER'
            ))
            if extra_gap > 0: story.append(Spacer(1, extra_gap / 2))
        
        include_title = row.get('IncludeTitle', True) if 'IncludeTitle' in resources.columns else True
        
        if include_title:
            is_small = row.get('SmallTitle', False) if 'SmallTitle' in resources.columns else False
            active_title_size = (title_size * 0.75) if is_small else title_size 
            scaled_title_size = active_title_size * scale_factor
            
            title_font = row.get('TitleFont', 'Rockwell-Bold') if 'TitleFont' in resources.columns else 'Rockwell-Bold'
            
            dynamic_title_style = ParagraphStyle(
                'DynamicTitle', 
                parent=styles['Normal'], 
                fontName=title_font, 
                fontSize=scaled_title_size, 
                leading=scaled_title_size * 1.2, 
                alignment=TA_CENTER, 
                spaceAfter=2 * scale_factor
            )
            
            story.append(Paragraph(str(row['Resource']), dynamic_title_style))
            if extra_gap > 0: story.append(Spacer(1, extra_gap))
        
        if include_images:
            if pd.notna(row['Image']) and str(row['Image']).strip():
                image_filename = str(row['Image']).strip()
                img_path = os.path.join(ASSETS_DIR, image_filename)
                
                if os.path.exists(img_path):
                    processed_img_path = get_resized_image_path(img_path)
                    scaled_img_height_pts = image_size_mm * scale_factor * mm
                    max_w_pts = (LABEL_WIDTH - (4 * mm)) * scale_factor 
                    story.append(Image(processed_img_path, width=max_w_pts, height=scaled_img_height_pts, kind='proportional'))
                    if extra_gap > 0: story.append(Spacer(1, extra_gap))
        
        if pd.notna(row['Description']) and row['Description']:
            story.append(Paragraph(str(row['Description']), style_desc_scaled))
            if extra_gap > 0: story.append(Spacer(1, extra_gap))
                    
        copyright_text = str(row['Copyright']) if pd.notna(row['Copyright']) and str(row['Copyright']).strip() else ""
        has_cc_img = False
        cc_img_obj = None
        
        text_only_cc = row.get('TextOnlyCC', False) if 'TextOnlyCC' in resources.columns else False
        raw_cc = str(row['CC_License']).strip() if pd.notna(row['CC_License']) else ""
        license_type = None
        
        if "by-nc-sa/4.0" in raw_cc:
            license_type = "by-nc-sa"
            cc_label = "CC BY-NC-SA 4.0"
            cc_file = "by-nc-sa.png"
        elif "by-nc-nd/4.0" in raw_cc:
            license_type = "by-nc-nd"
            cc_label = "CC BY-NC-ND 4.0"
            cc_file = "by-nc-nd.png"
            
        if license_type:
            if text_only_cc:
                if copyright_text:
                    copyright_text += f" - {cc_label}"
                else:
                    copyright_text = cc_label
            elif include_images:
                cc_img_path = os.path.join(ASSETS_DIR, cc_file)
                if os.path.exists(cc_img_path):
                    cc_w = 24 * scale_factor
                    cc_h = 8 * scale_factor
                    cc_img_obj = Image(cc_img_path, width=cc_w*mm, height=cc_h*mm, kind='bound')
                    has_cc_img = True
        
        if has_cc_img or copyright_text:
            if has_cc_img and copyright_text:
                style_copy_table = ParagraphStyle('CopyTable', parent=style_copy_scaled, alignment=TA_LEFT)
                copy_p = Paragraph(copyright_text, style_copy_table)
                col_widths = [cc_w * mm + (2 * mm), None] 
                core_flowable = Table([[cc_img_obj, copy_p]], colWidths=col_widths)
                core_flowable.setStyle(TableStyle([
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('LEFTPADDING', (0,0), (-1,-1), 0),
                    ('RIGHTPADDING', (0,0), (-1,-1), 0),
                    ('TOPPADDING', (0,0), (-1,-1), 0),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 0),
                ]))
            elif has_cc_img:
                core_flowable = cc_img_obj
            else:
                core_flowable = Paragraph(copyright_text, style_copy_scaled)
                
            if cc_shift != 0 and has_cc_img:
                shift_table_style = TableStyle([
                    ('LEFTPADDING', (0,0), (-1,-1), 0),
                    ('RIGHTPADDING', (0,0), (-1,-1), 0),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 0),
                    ('TOPPADDING', (0,0), (-1,-1), 0),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
                ])
                if cc_shift > 0:
                    cc_final = Table([[Spacer(cc_shift * 2, 1), core_flowable]], colWidths=[cc_shift * 2, None])
                else:
                    cc_final = Table([[core_flowable, Spacer(abs(cc_shift) * 2, 1)]], colWidths=[None, abs(cc_shift) * 2])
                    
                cc_final.setStyle(shift_table_style)
                story.append(cc_final)
            else:
                story.append(core_flowable)
            
    return story


def measure_story_height(story, width, canv):
    dummy_frame = KeepInFrame(width, 1000*mm, story)
    calculated_width, calculated_height = dummy_frame.wrapOn(canv, width, 1000*mm)
    return calculated_height


def generate_pdf(jobs, output_filename="SD_Labels.pdf", rows_per_page=3, img_v_offset=0.0, img_zoom=1.0):
    c = canvas.Canvas(output_filename, pagesize=A4)
    c.setLineWidth(0.5)

    label_height, margin_x, gutter_x, margin_y, gutter_y = get_layout_geometry(rows_per_page)
    labels_per_page = COLS * rows_per_page
    labels_drawn = 0
    
    for job in jobs:
        count = job['count']
        
        # --- NEW: Extract the job's specific snapshot, fallback to global if missing ---
        job_v_offset = job.get('img_v_offset', img_v_offset)
        job_zoom = job.get('img_zoom', img_zoom)
        
        # --- Direct Image Import Job ---
        if job.get('is_image_import', False):
            raw_img_file = job['imported_image_path']

            for _ in range(count):
                if labels_drawn > 0 and labels_drawn % labels_per_page == 0:
                    c.showPage()
                col = labels_drawn % COLS
                row = (rows_per_page - 1) - ((labels_drawn // COLS) % rows_per_page)
                x = margin_x + (col * (LABEL_WIDTH + gutter_x))
                y = margin_y + (row * (label_height + gutter_y))
                
                c.setFillColor(colors.white)
                c.rect(x, y, LABEL_WIDTH, label_height, fill=1, stroke=0)
                
                if os.path.exists(raw_img_file):
                    c.saveState()
                    path = c.beginPath()
                    path.rect(x, y, LABEL_WIDTH, label_height)
                    c.clipPath(path, stroke=0)

                    with PILImage.open(raw_img_file) as im:
                        im_w, im_h = im.size

                    fit_scale = min(LABEL_WIDTH / im_w, label_height / im_h)
                    draw_w = im_w * fit_scale * job_zoom
                    draw_h = im_h * fit_scale * job_zoom

                    center_x = x + (LABEL_WIDTH / 2.0)
                    center_y = y + (label_height / 2.0) + (job_v_offset * label_height * 0.5)

                    draw_x = center_x - (draw_w / 2.0)
                    draw_y = center_y - (draw_h / 2.0)

                    c.drawImage(
                        raw_img_file,
                        draw_x,
                        draw_y,
                        width=draw_w,
                        height=draw_h,
                        preserveAspectRatio=True,
                        mask='auto'
                    )
                    c.restoreState()
                else:
                    c.rect(x, y, LABEL_WIDTH, label_height)
                labels_drawn += 1
            continue

        # --- Standard Dynamic Label Job ---
        language = job['language']
        resources = job['df']
        border_filename = job.get('border_filename', 'border.png')
        scale_factor = job['scale_factor']
        vertical_shift = job['vertical_shift']
        cc_shift = job['cc_shift']
        
        for _ in range(count):
            if labels_drawn > 0 and labels_drawn % labels_per_page == 0:
                c.showPage() 
                
            col = labels_drawn % COLS
            row = (rows_per_page - 1) - ((labels_drawn // COLS) % rows_per_page)
            
            x = margin_x + (col * (LABEL_WIDTH + gutter_x))
            y = margin_y + (row * (label_height + gutter_y))
            
            # 1. Background Fill
            c.setFillColor(colors.white)
            c.rect(x, y, LABEL_WIDTH, label_height, fill=1, stroke=0)
            
            # 2. Save state & apply user transforms (Zoom/Nudge) using JOB SPECIFIC settings
            c.saveState()
            
            path = c.beginPath()
            path.rect(x, y, LABEL_WIDTH, label_height)
            c.clipPath(path, stroke=0)
            
            center_x = x + (LABEL_WIDTH / 2.0)
            center_y = y + (label_height / 2.0)
            
            c.translate(center_x, center_y)
            c.translate(0, job_v_offset * label_height * 0.5)
            c.scale(job_zoom, job_zoom)
            c.translate(-center_x, -center_y)
            
            # 3. Dedicated Language Header (8mm reserved gap)
            header_space = 8.0 * mm
            border_h = label_height - header_space
            
            border_img_path = os.path.join(ASSETS_DIR, border_filename)
            if os.path.exists(border_img_path):
                c.drawImage(border_img_path, x, y, width=LABEL_WIDTH, height=border_h, mask='auto')
            else:
                c.setStrokeColor(colors.black)
                c.rect(x, y, LABEL_WIDTH, border_h, fill=0, stroke=1)
            
            c.setFont("Arial", 8 * (0.85 if rows_per_page == 4 else 1.0))
            c.setFillColor(colors.black)
            spaced_language = " ".join(language.upper())
            text_x = x + (LABEL_WIDTH / 2)
            text_y = y + border_h + (2.5 * mm) 
            c.drawCentredString(text_x, text_y, spaced_language)
            
            # 4. Content Layout Frame
            max_w = (LABEL_WIDTH - (4*mm)) * scale_factor
            max_h_shrunk = (border_h - (4*mm)) * scale_factor
            frame_x = x + 2*mm + ((LABEL_WIDTH - (4*mm)) - max_w) / 2
            
            t_size = 9 if rows_per_page == 4 else 10
            img_size = 8 if rows_per_page == 4 else 12
            
            base_story = build_label_story(resources, include_images=True, title_size=t_size, image_size_mm=img_size, scale_factor=scale_factor, cc_shift=cc_shift)
            base_height = measure_story_height(base_story, max_w, c)
            
            has_images = True
            if base_height > max_h_shrunk:
                has_images = False
                t_size = 8 if rows_per_page == 4 else 10
                img_size = 0
                final_story = build_label_story(resources, include_images=False, title_size=t_size, image_size_mm=img_size, scale_factor=scale_factor, cc_shift=cc_shift)
            else:
                final_story = base_story
                while True:
                    t_size += 1
                    img_size += 1.0 if rows_per_page == 4 else 1.5
                    if t_size > 28:
                        break
                    test_story = build_label_story(resources, include_images=True, title_size=t_size, image_size_mm=img_size, scale_factor=scale_factor, cc_shift=cc_shift)
                    test_height = measure_story_height(test_story, max_w, c)
                    if test_height > max_h_shrunk:
                        t_size -= 1
                        img_size -= 1.0 if rows_per_page == 4 else 1.5
                        break
                    final_story = test_story
            
            current_height = measure_story_height(final_story, max_w, c)
            physical_h = (border_h - (4 * mm)) 
            leftover_space = physical_h - current_height
            
            if leftover_space > (3 * mm): 
                estimated_gaps = max(1, len(resources) * 3)
                gap_size = (leftover_space * 0.35) / estimated_gaps 
                final_story = build_label_story(
                    resources, 
                    include_images=has_images, 
                    title_size=t_size, 
                    image_size_mm=img_size, 
                    scale_factor=scale_factor, 
                    extra_gap=gap_size, 
                    cc_shift=cc_shift
                )
                
            final_h = measure_story_height(final_story, max_w, c)
            frame_y = y + 2*mm + (physical_h - final_h) / 2 - vertical_shift
            
            frame = Frame(frame_x, frame_y, max_w, final_h + (1*mm), 
                          leftPadding=0, bottomPadding=0, rightPadding=0, topPadding=0)
            kif = KeepInFrame(max_w, final_h + (1*mm), final_story, mode='shrink')
            frame.addFromList([kif], c)
            
            # Restore state
            c.restoreState()
            
            labels_drawn += 1
        
    c.save()
    print(f"Success! Generated {output_filename} with a total of {labels_drawn} labels.")


def parse_input_to_jobs(user_input, df, rows_per_page=3):
    job_strings = [j.strip() for j in user_input.split(';') if j.strip()]
    jobs = []
    default_count = COLS * rows_per_page

    islamic_borders = {
        1: 'islamic_border.png',
        2: 'islamic_border2.jpg',
        3: 'islamic_border3.jpg',
        4: 'islamic_border4.jpg',
        5: 'islamic_border5.png',
    }

    def extract_font(token):
        f_count = 0
        while token.endswith('f'):
            f_count += 1
            token = token[:-1]
        font_idx = 1 + f_count
        return RESOURCE_FONTS.get(font_idx, 'Rockwell-Bold'), token

    for job_str in job_strings:
        count = default_count
        count_match = re.search(r'\((\d+)\)$', job_str.strip())
        if count_match:
            count = int(count_match.group(1))
            job_str = job_str[:count_match.start()].strip()

        clean_path = job_str.strip('\'"')
        if clean_path.lower().endswith(('.png', '.jpg', '.jpeg')) and os.path.exists(clean_path):
            jobs.append({
                'is_image_import': True,
                'imported_image_path': clean_path,
                'count': count
            })
            continue

        parts = [p.strip() for p in job_str.split(',')]
        language_name = parts[0]

        use_islamic = False
        border_filename = 'border.png'
        custom_shrink = None
        vertical_shift = 0
        cc_shift = 0
        resource_requests = []

        for item in parts[1:]:
            item_lower = item.lower()
            if item_lower and all(ch == 'i' for ch in item_lower):
                use_islamic = True
                i_count = len(item_lower)
                border_filename = islamic_borders.get(i_count, f'islamic_border{i_count}.jpg')
            elif item_lower.startswith('cc') and '-cc' not in item_lower:
                try:
                    cc_shift = float(item_lower[2:])
                except ValueError:
                    pass
            elif item_lower.startswith('s'):
                try:
                    custom_shrink = float(item_lower[1:])
                except ValueError:
                    pass
            elif item_lower.startswith('l'):
                try:
                    vertical_shift = float(item_lower[1:])
                except ValueError:
                    pass
            elif '&' in item_lower:
                sub_items = [s.strip() for s in item_lower.split('&') if s.strip()]
                for idx, sub in enumerate(sub_items):
                    text_only_cc = False
                    if '-cc' in sub:
                        text_only_cc = True
                        sub = sub.replace('-cc', '')
                    include_title = not sub.endswith('--')
                    include_desc = not sub.endswith('-')
                    clean_sub = sub.rstrip('-')
                    font_name, clean_sub = extract_font(clean_sub)
                    try:
                        line_num = int(clean_sub)
                        resource_requests.append({
                            'line_num': line_num,
                            'include_desc': include_desc,
                            'include_title': include_title,
                            'hide_divider': idx > 0,
                            'show_copyright': idx == len(sub_items) - 1,
                            'small_title': True,
                            'text_only_cc': text_only_cc,
                            'title_font': font_name
                        })
                    except ValueError:
                        pass
            else:
                text_only_cc = False
                if '-cc' in item_lower:
                    text_only_cc = True
                    item_lower = item_lower.replace('-cc', '')
                include_title = not item_lower.endswith('--')
                include_desc = not (item_lower.endswith('-') or item_lower.endswith('--'))
                clean_item = item_lower.rstrip('-')
                font_name, clean_item = extract_font(clean_item)
                try:
                    line_num = int(clean_item)
                    resource_requests.append({
                        'line_num': line_num,
                        'include_desc': include_desc,
                        'include_title': include_title,
                        'hide_divider': False,
                        'show_copyright': True,
                        'text_only_cc': text_only_cc,
                        'title_font': font_name
                    })
                except ValueError:
                    pass

        scale_factor = 1.0
        if custom_shrink is not None:
            scale_factor = 1.0 - (custom_shrink / 100.0)
        elif use_islamic:
            scale_factor = 0.85

        selected_resources = []
        for req in resource_requests:
            line_num = req['line_num']
            row_index = line_num - 1

            col_b = df.iloc[row_index, 1] if df.shape[1] > 1 else ''
            col_c = df.iloc[row_index, 2] if df.shape[1] > 2 else ''
            col_d = df.iloc[row_index, 3] if df.shape[1] > 3 else ''
            col_e = df.iloc[row_index, 4] if df.shape[1] > 4 else ''
            col_f = df.iloc[row_index, 5] if df.shape[1] > 5 else ''

            selected_resources.append({
                'Language': language_name,
                'Description': (col_b if pd.notna(col_b) else '') if req['include_desc'] else '',
                'Resource': col_c if pd.notna(col_c) else '',
                'Copyright': (col_d if pd.notna(col_d) else '') if req['show_copyright'] else '',
                'CC_License': col_e if req['show_copyright'] else '',
                'Image': col_f,
                'HideDivider': req['hide_divider'],
                'SmallTitle': req.get('small_title', False),
                'IncludeTitle': req['include_title'],
                'TextOnlyCC': req['text_only_cc'],
                'TitleFont': req['title_font']
            })

        target_df = pd.DataFrame(selected_resources)

        jobs.append({
            'language': language_name,
            'df': target_df,
            'use_islamic': use_islamic,
            'border_filename': border_filename,
            'scale_factor': scale_factor,
            'vertical_shift': vertical_shift,
            'cc_shift': cc_shift,
            'count': count,
        })

    return jobs