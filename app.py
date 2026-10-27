import io
from flask import Flask, render_template_string, request, send_file

# PDF Generation Libraries
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Word Document Generation Libraries
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

app = Flask(__name__)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Professional CV Maker</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background-color: #f4f7f6; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .cv-card { background: #ffffff; border-radius: 12px; box-shadow: 0 8px 20px rgba(0,0,0,0.06); overflow: hidden; }
        .section-header { border-bottom: 2px solid #2b3a4a; padding-bottom: 6px; color: #2b3a4a; font-weight: 700; margin-bottom: 18px; text-transform: uppercase; font-size: 1.05rem; }
        .btn-ai { background-color: #6c5ce7; color: white; border: none; }
        .btn-ai:hover { background-color: #5a4bcf; color: white; }
        .dynamic-item { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 15px; margin-bottom: 15px; position: relative; }
        .btn-remove { position: absolute; top: 10px; right: 12px; color: #e53e3e; cursor: pointer; font-weight: bold; background: none; border: none; font-size: 1.1rem; }
        .btn-remove:hover { color: #c53030; }
    </style>
    <script async src="https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js?client=ca-pub-3410254230554312"
     crossorigin="anonymous"></script>
</head>
<body>
    <div class="container my-5" style="max-width: 950px;">
        <div class="cv-card p-4 p-md-5">
            <h2 class="text-center fw-bold mb-1" style="color: #2b3a4a;">Professional CV Maker</h2>
            <p class="text-center text-muted mb-4">Fill in your information and generate your formatted CV instantly in PDF or Word format.</p>
            
            <form id="cvForm" action="/generate-cv" method="POST">
                <input type="hidden" name="output_format" id="output_format" value="pdf">
                
                <div class="row g-3">
                    <!-- LEFT SIDEBAR COLUMN -->
                    <div class="col-md-5 pe-md-4" style="border-right: 1px solid #e2e8f0;">
                        <div class="section-header text-primary">Sidebar & Header Info</div>
                        
                        <div class="mb-3">
                            <label class="form-label fw-bold">Full Name *</label>
                            <input type="text" name="full_name" class="form-control" placeholder="e.g. John Doe" required>
                        </div>
                        <div class="mb-3">
                            <label class="form-label fw-bold">Professional Title *</label>
                            <input type="text" name="job_title" class="form-control" placeholder="e.g. Senior Software Engineer" required>
                        </div>
                        <div class="mb-3">
                            <label class="form-label fw-semibold">Address / Location</label>
                            <textarea name="contact_address" class="form-control" rows="2" placeholder="e.g. 123 Street Name, New York, USA"></textarea>
                        </div>
                        <div class="mb-3">
                            <label class="form-label fw-semibold">Phone Number</label>
                            <input type="text" name="phone" class="form-control" placeholder="e.g. +1 234 567 890">
                        </div>
                        <div class="mb-3">
                            <label class="form-label fw-semibold">Email Address</label>
                            <input type="email" name="email" class="form-control" placeholder="e.g. johndoe@example.com">
                        </div>
                        <div class="mb-3">
                            <label class="form-label fw-semibold">Personal Info (Key-Value pairs)</label>
                            <textarea name="personal_details" class="form-control" rows="3" placeholder="Nationality: American&#10;LinkedIn: linkedin.com/in/johndoe"></textarea>
                        </div>
                        <div class="mb-3">
                            <label class="form-label fw-semibold">Core Skills (One per line)</label>
                            <textarea name="skills" class="form-control" rows="5" placeholder="Project Management&#10;Python Programming&#10;Data Analysis"></textarea>
                        </div>
                        <div class="mb-3">
                            <label class="form-label fw-semibold">Languages</label>
                            <textarea name="languages" class="form-control" rows="3" placeholder="English — Native&#10;Spanish — Intermediate"></textarea>
                        </div>
                    </div>

                    <!-- RIGHT MAIN CONTENT COLUMN -->
                    <div class="col-md-7 ps-md-4">
                        <div class="section-header text-primary">Main Content</div>
                        
                        <!-- Summary Section -->
                        <div class="mb-4">
                            <div class="d-flex justify-content-between align-items-center mb-1">
                                <label class="form-label fw-semibold mb-0">Professional Summary</label>
                                <button type="button" class="btn btn-ai btn-sm" onclick="enhanceSummaryWithAI()">✨ AI Polish Text</button>
                            </div>
                            <textarea id="summary_input" name="summary" class="form-control" rows="4" placeholder="Dedicated professional with experience in..."></textarea>
                        </div>

                        <!-- Work Experience Section -->
                        <div class="mb-4">
                            <div class="d-flex justify-content-between align-items-center mb-2">
                                <label class="form-label fw-bold mb-0">Work Experience</label>
                                <button type="button" class="btn btn-outline-primary btn-sm" onclick="addExperience()">+ Add Experience</button>
                            </div>
                            <div id="experience_container"></div>
                        </div>

                        <!-- Education Section -->
                        <div class="mb-4">
                            <div class="d-flex justify-content-between align-items-center mb-2">
                                <label class="form-label fw-bold mb-0">Education</label>
                                <button type="button" class="btn btn-outline-primary btn-sm" onclick="addEducation()">+ Add Education</button>
                            </div>
                            <div id="education_container"></div>
                        </div>

                        <!-- Certifications Section -->
                        <div class="mb-4">
                            <div class="d-flex justify-content-between align-items-center mb-2">
                                <label class="form-label fw-bold mb-0">Certifications & Training</label>
                                <button type="button" class="btn btn-outline-primary btn-sm" onclick="addCertification()">+ Add Certification</button>
                            </div>
                            <div id="cert_container"></div>
                        </div>

                    </div>
                </div>

                <!-- DOWNLOAD ACTION BUTTONS -->
                <div class="row mt-4 pt-3 border-top">
                    <div class="col-md-6 mb-2 mb-md-0">
                        <button type="button" class="btn btn-danger btn-lg w-100 fw-bold" onclick="submitForm('pdf')">
                            📄 Download PDF
                        </button>
                    </div>
                    <div class="col-md-6">
                        <button type="button" class="btn btn-primary btn-lg w-100 fw-bold" onclick="submitForm('docx')">
                            📝 Download Word (.docx)
                        </button>
                    </div>
                </div>
            </form>
        </div>
    </div>

    <script>
        function addExperience(title = '', duration = '', company = '', desc = '') {
            var container = document.getElementById('experience_container');
            var id = 'exp_' + Math.random().toString(36).substring(2, 9);
            var item = document.createElement('div');
            item.className = 'dynamic-item';
            item.id = id;
            item.innerHTML = `
                <button type="button" class="btn-remove" onclick="removeElement('${id}')">✕</button>
                <div class="row g-2 mb-2">
                    <div class="col-md-7">
                        <input type="text" name="exp_title[]" value="${title}" class="form-control form-control-sm" placeholder="Job Title (e.g. Project Manager)">
                    </div>
                    <div class="col-md-5">
                        <input type="text" name="exp_duration[]" value="${duration}" class="form-control form-control-sm" placeholder="Dates (e.g. 2021 - Present)">
                    </div>
                    <div class="col-12">
                        <input type="text" name="exp_company[]" value="${company}" class="form-control form-control-sm" placeholder="Company Name & Location (e.g. ABC Corp)">
                    </div>
                    <div class="col-12 mt-2">
                        <textarea name="exp_desc[]" class="form-control form-control-sm" rows="2" placeholder="Key achievements or responsibilities...">${desc}</textarea>
                        <button type="button" class="btn btn-ai btn-sm mt-1" onclick="aiFixText(this)">✨ AI Polish Bullet Points</button>
                    </div>
                </div>
            `;
            container.appendChild(item);
        }

        function addEducation(degree = '', institute = '', year = '') {
            var container = document.getElementById('education_container');
            var id = 'edu_' + Math.random().toString(36).substring(2, 9);
            var item = document.createElement('div');
            item.className = 'dynamic-item';
            item.id = id;
            item.innerHTML = `
                <button type="button" class="btn-remove" onclick="removeElement('${id}')">✕</button>
                <div class="row g-2">
                    <div class="col-md-6">
                        <input type="text" name="edu_degree[]" value="${degree}" class="form-control form-control-sm" placeholder="Degree / Qualification">
                    </div>
                    <div class="col-md-4">
                        <input type="text" name="edu_institute[]" value="${institute}" class="form-control form-control-sm" placeholder="University / School">
                    </div>
                    <div class="col-md-2">
                        <input type="text" name="edu_year[]" value="${year}" class="form-control form-control-sm" placeholder="Year">
                    </div>
                </div>
            `;
            container.appendChild(item);
        }

        function addCertification(name = '', issuer = '', year = '') {
            var container = document.getElementById('cert_container');
            var id = 'cert_' + Math.random().toString(36).substring(2, 9);
            var item = document.createElement('div');
            item.className = 'dynamic-item';
            item.id = id;
            item.innerHTML = `
                <button type="button" class="btn-remove" onclick="removeElement('${id}')">✕</button>
                <div class="row g-2">
                    <div class="col-md-6">
                        <input type="text" name="cert_name[]" value="${name}" class="form-control form-control-sm" placeholder="Certification Title">
                    </div>
                    <div class="col-md-4">
                        <input type="text" name="cert_issuer[]" value="${issuer}" class="form-control form-control-sm" placeholder="Issuing Organization">
                    </div>
                    <div class="col-md-2">
                        <input type="text" name="cert_year[]" value="${year}" class="form-control form-control-sm" placeholder="Year">
                    </div>
                </div>
            `;
            container.appendChild(item);
        }

        function removeElement(id) {
            var elem = document.getElementById(id);
            if (elem) {
                elem.parentNode.removeChild(elem);
            }
        }

        function submitForm(format) {
            var form = document.getElementById('cvForm');
            document.getElementById('output_format').value = format;
            if (form.checkValidity()) {
                form.submit();
            } else {
                form.reportValidity();
            }
        }

        function enhanceSummaryWithAI() {
            var textarea = document.getElementById('summary_input');
            var text = textarea.value.trim();
            if(!text) {
                alert('Please type a brief summary in the box first.');
                return;
            }
            text = text.replace(/freelaner/gi, 'freelancer')
                       .replace(/spaced/gi, 'specialized')
                       .replace(/experiance/gi, 'experience')
                       .replace(/managment/gi, 'management');
            
            if(!text.toLowerCase().startsWith('results-driven') && !text.toLowerCase().startsWith('experienced')) {
                text = "Results-driven professional with strong expertise: " + text;
            }
            textarea.value = text;
        }

        function aiFixText(btnElement) {
            var textarea = btnElement.previousElementSibling;
            var text = textarea.value;
            if(!text) return;

            var lines = text.split('\\n');
            var cleanedLines = [];
            for (var i = 0; i < lines.length; i++) {
                var line = lines[i].trim();
                if (line.length > 0) {
                    var fixed = line.replace(/^[▸•\\-\\*]\\s*/, '');
                    cleanedLines.push("• " + fixed.charAt(0).toUpperCase() + fixed.slice(1));
                }
            }
            textarea.value = cleanedLines.join('\\n');
        }

        window.onload = function() {
            addExperience();
            addEducation();
            addCertification();
        };
    </script>
</body>
</html>
"""

def build_pdf_cv(data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=15, leftMargin=15, topMargin=15, bottomMargin=15)
    
    styles = getSampleStyleSheet()
    name_style = ParagraphStyle('NameStyle', parent=styles['Heading1'], fontSize=18, leading=22, textColor=colors.HexColor('#1E293B'), fontName='Helvetica-Bold', spaceAfter=2)
    subtitle_style = ParagraphStyle('SubTitleStyle', parent=styles['Normal'], fontSize=10, leading=13, textColor=colors.HexColor('#0F766E'), fontName='Helvetica-Bold', spaceAfter=12)
    
    left_header_style = ParagraphStyle('LeftHeader', parent=styles['Heading2'], fontSize=10, leading=13, textColor=colors.HexColor('#0F766E'), fontName='Helvetica-Bold', spaceBefore=10, spaceAfter=4)
    left_body_style = ParagraphStyle('LeftBody', parent=styles['Normal'], fontSize=8.5, leading=12, textColor=colors.HexColor('#334155'))
    
    right_header_style = ParagraphStyle('RightHeader', parent=styles['Heading2'], fontSize=11, leading=14, textColor=colors.HexColor('#1E293B'), fontName='Helvetica-Bold', spaceBefore=12, spaceAfter=4)
    right_body_style = ParagraphStyle('RightBody', parent=styles['Normal'], fontSize=9, leading=13, textColor=colors.HexColor('#334155'))

    elements = []

    # 1. Left Column Content (Sidebar with Name & Title at the Top)
    left_flowables = []
    
    # Name and Title inside Sidebar Header
    left_flowables.append(Paragraph(data.get('full_name', 'YOUR NAME').upper(), name_style))
    left_flowables.append(Paragraph(data.get('job_title', ''), subtitle_style))

    if data.get('contact_address') or data.get('phone') or data.get('email'):
        left_flowables.append(Paragraph("CONTACT", left_header_style))
        contact_txt = ""
        if data.get('contact_address'): contact_txt += f"<b>Address:</b><br/>{data.get('contact_address')}<br/>"
        if data.get('phone'): contact_txt += f"<b>Phone:</b><br/>{data.get('phone')}<br/>"
        if data.get('email'): contact_txt += f"<b>Email:</b><br/>{data.get('email')}<br/>"
        left_flowables.append(Paragraph(contact_txt, left_body_style))

    if data.get('personal_details'):
        left_flowables.append(Paragraph("PERSONAL DETAILS", left_header_style))
        p_txt = data.get('personal_details').replace('\n', '<br/>')
        left_flowables.append(Paragraph(p_txt, left_body_style))

    if data.get('skills'):
        left_flowables.append(Paragraph("CORE SKILLS", left_header_style))
        skills_lines = [s.strip() for s in data.get('skills').split('\n') if s.strip()]
        formatted_skills = "<br/>".join([f"• {s}" for s in skills_lines])
        left_flowables.append(Paragraph(formatted_skills, left_body_style))

    if data.get('languages'):
        left_flowables.append(Paragraph("LANGUAGES", left_header_style))
        lang_txt = data.get('languages').replace('\n', '<br/>')
        left_flowables.append(Paragraph(lang_txt, left_body_style))

    # 2. Right Column Content (Main Body)
    right_flowables = []
    
    if data.get('summary'):
        right_flowables.append(Paragraph("PROFESSIONAL SUMMARY", right_header_style))
        right_flowables.append(Paragraph(data.get('summary').replace('\n', '<br/>'), right_body_style))

    exp_titles = request.form.getlist('exp_title[]')
    exp_durations = request.form.getlist('exp_duration[]')
    exp_companies = request.form.getlist('exp_company[]')
    exp_descs = request.form.getlist('exp_desc[]')

    if any(exp_titles):
        right_flowables.append(Paragraph("WORK EXPERIENCE", right_header_style))
        for i in range(len(exp_titles)):
            if exp_titles[i]:
                header = f"<b>{exp_titles[i]}</b>"
                if exp_companies[i]: header += f"<br/><i>{exp_companies[i]}</i>"
                if exp_durations[i]: header += f" &nbsp;|&nbsp; <b>{exp_durations[i]}</b>"
                right_flowables.append(Paragraph(header, right_body_style))
                if exp_descs[i]:
                    right_flowables.append(Paragraph(exp_descs[i].replace('\n', '<br/>'), right_body_style))
                right_flowables.append(Spacer(1, 5))

    edu_degrees = request.form.getlist('edu_degree[]')
    edu_institutes = request.form.getlist('edu_institute[]')
    edu_years = request.form.getlist('edu_year[]')

    if any(edu_degrees):
        right_flowables.append(Paragraph("EDUCATION", right_header_style))
        for i in range(len(edu_degrees)):
            if edu_degrees[i]:
                edu_line = f"<b>{edu_degrees[i]}</b>"
                if edu_institutes[i]: edu_line += f"<br/>{edu_institutes[i]}"
                if edu_years[i]: edu_line += f" | {edu_years[i]}"
                right_flowables.append(Paragraph(edu_line, right_body_style))
                right_flowables.append(Spacer(1, 4))

    cert_names = request.form.getlist('cert_name[]')
    cert_issuers = request.form.getlist('cert_issuer[]')
    cert_years = request.form.getlist('cert_year[]')

    if any(cert_names):
        right_flowables.append(Paragraph("CERTIFICATIONS & TRAINING", right_header_style))
        for i in range(len(cert_names)):
            if cert_names[i]:
                cert_line = f"<b>{cert_names[i]}</b>"
                if cert_issuers[i]: cert_line += f" - {cert_issuers[i]}"
                if cert_years[i]: cert_line += f" ({cert_years[i]})"
                right_flowables.append(Paragraph(cert_line, right_body_style))
                right_flowables.append(Spacer(1, 3))

    # 3. Full-Height 2-Column Table Layout
    two_col_table = Table([[left_flowables, right_flowables]], colWidths=[190, 360])
    two_col_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (0,0), colors.HexColor('#F0F4F8')),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 12),
        ('BOTTOMPADDING', (0,0), (-1,-1), 12),
        ('LEFTPADDING', (0,0), (0,0), 12),
        ('RIGHTPADDING', (0,0), (0,0), 12),
        ('LEFTPADDING', (1,0), (1,0), 15),
        ('RIGHTPADDING', (1,0), (1,0), 10),
    ]))

    elements.append(two_col_table)
    doc.build(elements)
    buffer.seek(0)
    return buffer

def set_cell_background(cell, fill_hex):
    """Utility to set cell background color in docx."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def build_docx_cv(data):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.4)
        section.right_margin = Inches(0.4)

    # Main 2-Column Layout
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    cell_left = table.cell(0, 0)
    cell_right = table.cell(0, 1)
    cell_left.width = Inches(2.4)
    cell_right.width = Inches(4.7)

    # Colored Left Column (Full Sidebar)
    set_cell_background(cell_left, "F0F4F8")

    # Name and Title inside Sidebar Header
    p_name = cell_left.add_paragraph()
    run_name = p_name.add_run(data.get('full_name', 'YOUR NAME').upper())
    run_name.bold = True
    run_name.font.size = Pt(18)
    run_name.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)

    p_sub = cell_left.add_paragraph()
    run_sub = p_sub.add_run(data.get('job_title', ''))
    run_sub.bold = True
    run_sub.font.size = Pt(10)
    run_sub.font.color.rgb = RGBColor(0x0F, 0x76, 0x6E)

    if data.get('contact_address') or data.get('phone') or data.get('email'):
        h = cell_left.add_paragraph()
        r = h.add_run("CONTACT")
        r.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0x0F, 0x76, 0x6E)
        
        c_text = f"{data.get('contact_address')}\nPhone: {data.get('phone')}\nEmail: {data.get('email')}"
        p = cell_left.add_paragraph(c_text)
        p.style.font.size = Pt(8.5)

    if data.get('personal_details'):
        h = cell_left.add_paragraph()
        r = h.add_run("\nPERSONAL DETAILS")
        r.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0x0F, 0x76, 0x6E)
        p = cell_left.add_paragraph(data.get('personal_details'))
        p.style.font.size = Pt(8.5)

    if data.get('skills'):
        h = cell_left.add_paragraph()
        r = h.add_run("\nCORE SKILLS")
        r.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0x0F, 0x76, 0x6E)
        skills_lines = [s.strip() for s in data.get('skills').split('\n') if s.strip()]
        p = cell_left.add_paragraph("\n".join([f"• {s}" for s in skills_lines]))
        p.style.font.size = Pt(8.5)

    if data.get('languages'):
        h = cell_left.add_paragraph()
        r = h.add_run("\nLANGUAGES")
        r.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0x0F, 0x76, 0x6E)
        p = cell_left.add_paragraph(data.get('languages'))
        p.style.font.size = Pt(8.5)

    # Right Column Formatting
    if data.get('summary'):
        h = cell_right.add_paragraph()
        r = h.add_run("PROFESSIONAL SUMMARY")
        r.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
        p = cell_right.add_paragraph(data.get('summary'))
        p.style.font.size = Pt(9.5)

    exp_titles = request.form.getlist('exp_title[]')
    exp_durations = request.form.getlist('exp_duration[]')
    exp_companies = request.form.getlist('exp_company[]')
    exp_descs = request.form.getlist('exp_desc[]')

    if any(exp_titles):
        h = cell_right.add_paragraph()
        r = h.add_run("\nWORK EXPERIENCE")
        r.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)

        for i in range(len(exp_titles)):
            if exp_titles[i]:
                p_item = cell_right.add_paragraph()
                r_title = p_item.add_run(f"{exp_titles[i]}\n")
                r_title.bold = True
                if exp_companies[i]:
                    r_comp = p_item.add_run(f"{exp_companies[i]}  |  {exp_durations[i]}\n")
                    r_comp.italic = True
                if exp_descs[i]:
                    p_desc = cell_right.add_paragraph(exp_descs[i])
                    p_desc.style.font.size = Pt(9)

    edu_degrees = request.form.getlist('edu_degree[]')
    edu_institutes = request.form.getlist('edu_institute[]')
    edu_years = request.form.getlist('edu_year[]')

    if any(edu_degrees):
        h = cell_right.add_paragraph()
        r = h.add_run("\nEDUCATION")
        r.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)

        for i in range(len(edu_degrees)):
            if edu_degrees[i]:
                p_edu = cell_right.add_paragraph()
                r_deg = p_edu.add_run(f"{edu_degrees[i]}\n")
                r_deg.bold = True
                p_edu.add_run(f"{edu_institutes[i]}  |  {edu_years[i]}")
                p_edu.style.font.size = Pt(9)

    cert_names = request.form.getlist('cert_name[]')
    cert_issuers = request.form.getlist('cert_issuer[]')
    cert_years = request.form.getlist('cert_year[]')

    if any(cert_names):
        h = cell_right.add_paragraph()
        r = h.add_run("\nCERTIFICATIONS & TRAINING")
        r.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)

        for i in range(len(cert_names)):
            if cert_names[i]:
                p_cert = cell_right.add_paragraph()
                r_c = p_cert.add_run(f"•  {cert_names[i]}")
                r_c.bold = True
                p_cert.add_run(f" - {cert_issuers[i]} ({cert_years[i]})")
                p_cert.style.font.size = Pt(9)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

@app.route('/')
def home():
    return render_template_string(HTML_TEMPLATE)

@app.route('/generate-cv', methods=['POST'])
def generate_cv():
    data = request.form
    output_format = data.get('output_format', 'pdf')
    
    clean_name = data.get('full_name', 'CV').strip().replace(' ', '_')
    if not clean_name:
        clean_name = "Candidate"
    
    if output_format == 'docx':
        docx_buffer = build_docx_cv(data)
        return send_file(
            docx_buffer,
            as_attachment=True,
            download_name=f"{clean_name}_CV.docx",
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    else:
        pdf_buffer = build_pdf_cv(data)
        return send_file(
            pdf_buffer,
            as_attachment=True,
            download_name=f"{clean_name}_CV.pdf",
            mimetype='application/pdf'
        )

@app.route('/ads.txt')
def ads_txt():
    content = "google.com, pub-3410254230554312, DIRECT, f08c47fec0942fa0"
    return content, 200, {'Content-Type': 'text/plain'}

if __name__ == '__main__':
    app.run(debug=True, port=5000)
