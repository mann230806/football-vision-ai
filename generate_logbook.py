import os
import sys
import base64
import subprocess
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# -------------------------------------------------------------
# 1. LOAD LOGO
# -------------------------------------------------------------
LOGO_PATH = os.path.abspath("extracted_media/image1.png")
if os.path.exists(LOGO_PATH):
    with open(LOGO_PATH, "rb") as f:
        LOGO_B64 = base64.b64encode(f.read()).decode("utf-8")
else:
    LOGO_B64 = ""

# -------------------------------------------------------------
# 2. GENERATE COMPREHENSIVE HTML DOCUMENT
# -------------------------------------------------------------
html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>BE-Project-I Log Book - Football Vision AI</title>
<style>
  @page {{
    size: A4 portrait;
    margin: 12mm 15mm 12mm 15mm;
  }}
  * {{
    box-sizing: border-box;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }}
  body {{
    font-family: "Times New Roman", Times, serif;
    color: #111;
    margin: 0;
    padding: 0;
    background-color: #f3f4f6;
    font-size: 13px;
    line-height: 1.35;
  }}
  .no-print {{
    background: #1e293b;
    color: white;
    padding: 12px 24px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    position: sticky;
    top: 0;
    z-index: 1000;
    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
  }}
  .no-print button {{
    background: #2563eb;
    color: white;
    border: none;
    padding: 8px 18px;
    font-size: 14px;
    font-weight: 600;
    border-radius: 6px;
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    gap: 8px;
    transition: background 0.2s;
  }}
  .no-print button:hover {{
    background: #1d4ed8;
  }}
  .page {{
    width: 210mm;
    min-height: 297mm;
    padding: 16mm 18mm 16mm 18mm;
    margin: 15px auto;
    background: white;
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
    position: relative;
    page-break-after: always;
    display: flex;
    flex-direction: column;
  }}
  @media print {{
    body {{
      background: white;
    }}
    .no-print {{
      display: none !important;
    }}
    .page {{
      margin: 0;
      padding: 10mm 12mm 10mm 12mm;
      box-shadow: none;
      width: auto;
      min-height: auto;
      height: 100vh;
      page-break-after: always;
      page-break-inside: avoid;
    }}
  }}

  /* College Header */
  .header-table {{
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 12px;
    border-bottom: 2px solid #222;
    padding-bottom: 6px;
  }}
  .header-logo {{
    width: 90px;
    text-align: center;
    vertical-align: middle;
  }}
  .header-logo img {{
    max-width: 82px;
    height: auto;
    display: block;
    margin: 0 auto;
  }}
  .header-text {{
    text-align: center;
    vertical-align: middle;
    padding-left: 10px;
  }}
  .header-text .inst-trust {{
    font-size: 11.5px;
    font-weight: bold;
    letter-spacing: 0.5px;
    text-transform: uppercase;
  }}
  .header-text .inst-name {{
    font-size: 15.5px;
    font-weight: bold;
    letter-spacing: 0.8px;
    color: #0b1a30;
    margin: 2px 0;
    text-transform: uppercase;
  }}
  .header-text .inst-sub {{
    font-size: 9px;
    line-height: 1.25;
    color: #222;
  }}
  .header-text .inst-iso {{
    font-size: 9px;
    font-weight: bold;
    margin-top: 1px;
  }}
  .header-text .inst-dept {{
    font-size: 12.5px;
    font-weight: bold;
    color: #111;
    margin-top: 3px;
  }}

  /* Form Meta */
  .form-code-row {{
    display: flex;
    justify-content: flex-end;
    font-size: 10.5px;
    font-weight: bold;
    margin-bottom: 2px;
  }}
  .form-date-row {{
    display: flex;
    justify-content: flex-end;
    font-size: 10.5px;
    font-weight: bold;
    margin-bottom: 8px;
  }}

  /* Headings */
  h1.page-title {{
    text-align: center;
    font-size: 18px;
    font-weight: bold;
    margin: 16px 0 14px 0;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: #000;
  }}
  h2.section-title {{
    text-align: center;
    font-size: 14.5px;
    font-weight: bold;
    margin: 14px 0 8px 0;
    text-transform: uppercase;
    letter-spacing: 0.5px;
  }}
  h3.sub-title {{
    font-size: 12.5px;
    font-weight: bold;
    margin: 10px 0 4px 0;
  }}

  /* Content Elements */
  p {{
    margin: 6px 0;
    text-align: justify;
  }}
  ol, ul {{
    margin: 6px 0;
    padding-left: 24px;
  }}
  li {{
    margin-bottom: 4px;
    text-align: justify;
  }}

  /* Tables */
  table.content-table {{
    width: 100%;
    border-collapse: collapse;
    margin: 8px 0 12px 0;
    font-size: 11.5px;
  }}
  table.content-table th, table.content-table td {{
    border: 1px solid #111;
    padding: 5px 6px;
    vertical-align: middle;
  }}
  table.content-table th {{
    background-color: #f1f3f5;
    font-weight: bold;
    text-align: center;
  }}
  .text-center {{
    text-align: center !important;
  }}
  .text-left {{
    text-align: left !important;
  }}
  .font-bold {{
    font-weight: bold;
  }}
  .tick {{
    font-size: 14px;
    font-weight: bold;
    color: #054c15;
  }}

  /* Signatures */
  .sign-container {{
    display: flex;
    justify-content: space-between;
    align-items: flex-end;
    margin-top: auto;
    padding-top: 18px;
  }}
  .sign-box {{
    text-align: center;
    font-size: 11.5px;
    font-weight: bold;
  }}
  .sign-box .sign-space {{
    height: 42px;
  }}
  .sign-box .sign-line {{
    border-top: 1px solid #111;
    padding-top: 4px;
  }}
  .guide-sign-img {{
    font-family: 'Brush Script MT', 'Dancing Script', cursive;
    font-size: 20px;
    color: #002664;
    font-weight: normal;
  }}

  /* Info Form Grid */
  .info-meta {{
    font-size: 12.5px;
    line-height: 1.6;
    margin-bottom: 12px;
  }}
  .info-meta b {{
    font-weight: bold;
  }}

  /* Comments box */
  .comment-box {{
    border: 1px solid #111;
    padding: 8px 10px;
    margin-top: 10px;
    min-height: 55px;
    font-size: 11.5px;
  }}
  .handwritten {{
    font-family: "Segoe Script", "Comic Sans MS", cursive;
    color: #001f5c;
    font-size: 12.5px;
    font-weight: 600;
  }}
</style>
</head>
<body>

<div class="no-print">
  <div>
    <strong>Atharva College of Engineering</strong> — BE-Project-I Official Log Book (Football Vision AI)
  </div>
  <button onclick="window.print()">
    <svg width="18" height="18" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4H7v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"></path></svg>
    Print / Save as PDF
  </button>
</div>

<!-- ==========================================
     PAGE 1: COVER PAGE
========================================== -->
<div class="page">
  <table class="header-table">
    <tr>
      <td class="header-logo"><img src="data:image/png;base64,{LOGO_B64}" alt="Logo"></td>
      <td class="header-text">
        <div class="inst-trust">ATHARVA EDUCATIONAL TRUST'S</div>
        <div class="inst-name">ATHARVA COLLEGE OF ENGINEERING</div>
        <div class="inst-sub">(Approved by AICTE, Recognized by Government of Maharashtra & Affiliated to University of Mumbai - Estd. 1999 - 2000)</div>
        <div class="inst-iso">ISO 2100:2018 &nbsp; ISO 14001:2015 &nbsp; ISO 9001:2015</div>
        <div class="inst-sub"><b>NAAC Accredited with Grade A+</b></div>
        <div class="inst-dept">Department of Computer Engineering</div>
      </td>
    </tr>
  </table>

  <div style="flex: 1; display: flex; flex-direction: column; justify-content: space-around; text-align: center;">
    <div>
      <div style="font-size: 26px; font-weight: bold; letter-spacing: 2px; margin-top: 25px; color: #111;">
        BE-PROJECT-I LOG BOOK
      </div>
      <div style="font-size: 15px; font-weight: bold; margin-top: 14px; color: #1e3a8a; text-transform: uppercase; padding: 0 30px;">
        FOOTBALL VISION AI: COMPUTER VISION AND IOT POWERED AUTONOMOUS VIDEO ASSISTANT REFEREE (VAR)
      </div>
      <div style="font-size: 13px; font-weight: bold; margin-top: 6px; color: #555;">
        Academic Year: 2025 – 2026 | Semester VII
      </div>
    </div>

    <div>
      <div style="font-size: 15px; font-weight: bold; letter-spacing: 1px; margin-bottom: 16px;">
        GROUP MEMBERS
      </div>
      <table style="width: 65%; margin: 0 auto; font-size: 13.5px; border-collapse: collapse; text-align: left;">
        <tr><td style="padding: 5px 0; width: 45px; font-weight: bold;">1.</td><td style="padding: 5px 0; font-weight: bold;">Mann Sutariya</td><td style="padding: 5px 0; color: #444;">(Roll No: 47)</td></tr>
        <tr><td style="padding: 5px 0; width: 45px; font-weight: bold;">2.</td><td style="padding: 5px 0; font-weight: bold;">Tejas Walanj</td><td style="padding: 5px 0; color: #444;">(Roll No: 55)</td></tr>
        <tr><td style="padding: 5px 0; width: 45px; font-weight: bold;">3.</td><td style="padding: 5px 0; font-weight: bold;">Karan Sonar</td><td style="padding: 5px 0; color: #444;">(Roll No: 44)</td></tr>
        <tr><td style="padding: 5px 0; width: 45px; font-weight: bold;">4.</td><td style="padding: 5px 0; font-weight: bold;">Yashwin Rajpurohit</td><td style="padding: 5px 0; color: #444;">(Roll No: 09)</td></tr>
      </table>
    </div>

    <div style="margin-bottom: 25px;">
      <div class="guide-sign-img" style="margin-bottom: 4px;">Varsha Patil</div>
      <div style="font-size: 13px; font-weight: bold; letter-spacing: 0.5px;">SUPERVISOR / GUIDE</div>
      <div style="font-size: 15px; font-weight: bold; margin-top: 4px;">PROF. VARSHA PATIL</div>
      <div style="font-size: 11.5px; color: #555; margin-top: 2px;">Assistant Professor, Department of Computer Engineering</div>
    </div>
  </div>
</div>

<!-- ==========================================
     PAGE 2: VISION & MISSION
========================================== -->
<div class="page">
  <table class="header-table">
    <tr>
      <td class="header-logo"><img src="data:image/png;base64,{LOGO_B64}" alt="Logo"></td>
      <td class="header-text">
        <div class="inst-trust">ATHARVA EDUCATIONAL TRUST'S</div>
        <div class="inst-name">ATHARVA COLLEGE OF ENGINEERING</div>
        <div class="inst-sub">(Approved by AICTE, Recognized by Government of Maharashtra & Affiliated to University of Mumbai - Estd. 1999 - 2000)</div>
        <div class="inst-iso">ISO 2100:2018 &nbsp; ISO 14001:2015 &nbsp; ISO 9001:2015</div>
        <div class="inst-sub"><b>NAAC Accredited with Grade A+</b></div>
        <div class="inst-dept">Department of Computer Engineering</div>
      </td>
    </tr>
  </table>

  <h2 class="section-title" style="margin-top: 10px;">INSTITUTE VISION & MISSION</h2>

  <h3 class="sub-title">Vision</h3>
  <p>To develop a sustainable ecosystem of technical education by fostering innovation, advancing research and leveraging technology with ethical and social values.</p>

  <h3 class="sub-title">Mission</h3>
  <ol>
    <li>To transform the traditional approach of engineering education by experiential learning and interdisciplinary research for professional development.</li>
    <li>To promote, encourage and facilitate the use of emerging technology and its application to bridge the gap between industry and academia.</li>
    <li>To imbibe core values of excellence, integrity, teamwork, professional ethics and societal responsibilities.</li>
  </ol>

  <div style="border-top: 1px dashed #777; margin: 18px 0;"></div>

  <h2 class="section-title">DEPARTMENT OF COMPUTER ENGINEERING</h2>

  <h3 class="sub-title">Vision</h3>
  <p>To enhance the standard of Computer Engineering graduates with strong academics, research and innovative technology to converge the emerging needs of the industry and society along with ethical values.</p>

  <h3 class="sub-title">Mission</h3>
  <ol>
    <li>To create a productive teaching and learning atmosphere that will foster research initiatives, inventions by using modern tools.</li>
    <li>To develop and channelize the creative talents of graduates to achieve the evolving needs of the software industry.</li>
    <li>To inculcate problem solving, team building skills and promote life-long learning with a sense of imparting human values.</li>
  </ol>
</div>

<!-- ==========================================
     PAGE 3: PROGRAM OUTCOMES (PO1 - PO10)
========================================== -->
<div class="page">
  <table class="header-table">
    <tr>
      <td class="header-logo"><img src="data:image/png;base64,{LOGO_B64}" alt="Logo"></td>
      <td class="header-text">
        <div class="inst-trust">ATHARVA EDUCATIONAL TRUST'S</div>
        <div class="inst-name">ATHARVA COLLEGE OF ENGINEERING</div>
        <div class="inst-sub">(Approved by AICTE, Recognized by Government of Maharashtra & Affiliated to University of Mumbai - Estd. 1999 - 2000)</div>
        <div class="inst-iso">ISO 2100:2018 &nbsp; ISO 14001:2015 &nbsp; ISO 9001:2015</div>
        <div class="inst-sub"><b>NAAC Accredited with Grade A+</b></div>
        <div class="inst-dept">Department of Computer Engineering</div>
      </td>
    </tr>
  </table>

  <h2 class="section-title">Program Outcomes PO's</h2>

  <table class="content-table" style="font-size: 11px;">
    <thead>
      <tr>
        <th style="width: 10%;">PO Code</th>
        <th style="width: 25%;">PO Title</th>
        <th style="width: 65%;">Statement</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td class="text-center font-bold">PO1<br><small>EK</small></td>
        <td class="font-bold">Engineering knowledge</td>
        <td>Apply the knowledge of mathematics, science, engineering fundamentals, and an engineering specialization to the solution of complex engineering problems.</td>
      </tr>
      <tr>
        <td class="text-center font-bold">PO2<br><small>PA</small></td>
        <td class="font-bold">Problem analysis</td>
        <td>Identify, formulate, research literature, and analyze complex engineering problems reaching substantiated conclusions using first principles of mathematics, natural sciences, and engineering sciences.</td>
      </tr>
      <tr>
        <td class="text-center font-bold">PO3<br><small>DS</small></td>
        <td class="font-bold">Development/ Design of solutions</td>
        <td>Design solutions for complex engineering problems and design system components or processes that meet the specified needs with appropriate consideration for the public health and safety, and the cultural, societal, and environmental considerations.</td>
      </tr>
      <tr>
        <td class="text-center font-bold">PO4<br><small>CP</small></td>
        <td class="font-bold">Conduct investigations of complex problems</td>
        <td>Use research-based knowledge and research methods including design of experiments, analysis and interpretation of data, and synthesis of the information to provide valid conclusions.</td>
      </tr>
      <tr>
        <td class="text-center font-bold">PO5<br><small>MT</small></td>
        <td class="font-bold">Modern tool usage</td>
        <td>Create, select, and apply appropriate techniques, resources, and modern engineering and IT tools including prediction and modeling to complex engineering activities with an understanding of the limitations.</td>
      </tr>
      <tr>
        <td class="text-center font-bold">PO6<br><small>EG</small></td>
        <td class="font-bold">The engineer and society</td>
        <td>Apply reasoning informed by the contextual knowledge to assess societal, health, safety, legal and cultural issues and the consequent responsibilities relevant to the professional engineering practice.</td>
      </tr>
      <tr>
        <td class="text-center font-bold">PO7<br><small>ES</small></td>
        <td class="font-bold">Environment and sustainability</td>
        <td>Understand the impact of the professional engineering solutions in societal and environmental contexts, and demonstrate the knowledge of, and need for sustainable development.</td>
      </tr>
      <tr>
        <td class="text-center font-bold">PO8<br><small>ET</small></td>
        <td class="font-bold">Ethics</td>
        <td>Apply ethical principles and commit to professional ethics and responsibilities and norms of the engineering practice.</td>
      </tr>
      <tr>
        <td class="text-center font-bold">PO9<br><small>IT</small></td>
        <td class="font-bold">Individual and team work</td>
        <td>Function effectively as an individual, and as a member or leader in diverse teams, and in multidisciplinary settings.</td>
      </tr>
      <tr>
        <td class="text-center font-bold">PO10<br><small>CO</small></td>
        <td class="font-bold">Communication</td>
        <td>Communicate effectively on complex engineering activities with the engineering community and with society at large, such as, being able to comprehend and write effective reports and design documentation, make effective presentations, and give and receive clear instructions.</td>
      </tr>
    </tbody>
  </table>
</div>

<!-- ==========================================
     PAGE 4: PO11, PO12, PSO & PEO
========================================== -->
<div class="page">
  <table class="header-table">
    <tr>
      <td class="header-logo"><img src="data:image/png;base64,{LOGO_B64}" alt="Logo"></td>
      <td class="header-text">
        <div class="inst-trust">ATHARVA EDUCATIONAL TRUST'S</div>
        <div class="inst-name">ATHARVA COLLEGE OF ENGINEERING</div>
        <div class="inst-sub">(Approved by AICTE, Recognized by Government of Maharashtra & Affiliated to University of Mumbai - Estd. 1999 - 2000)</div>
        <div class="inst-iso">ISO 2100:2018 &nbsp; ISO 14001:2015 &nbsp; ISO 9001:2015</div>
        <div class="inst-sub"><b>NAAC Accredited with Grade A+</b></div>
        <div class="inst-dept">Department of Computer Engineering</div>
      </td>
    </tr>
  </table>

  <table class="content-table" style="font-size: 11px;">
    <tbody>
      <tr>
        <td class="text-center font-bold" style="width: 10%;">PO11<br><small>PM</small></td>
        <td class="font-bold" style="width: 25%;">Project management and finance</td>
        <td style="width: 65%;">Demonstrate knowledge and understanding of the engineering and management principles and apply these to one's own work, as a member and leader in a team, to manage projects and in multidisciplinary environments.</td>
      </tr>
      <tr>
        <td class="text-center font-bold">PO12<br><small>LL</small></td>
        <td class="font-bold">Life-long learning</td>
        <td>Recognize the need for, and have the preparation and ability to engage in independent and life-long learning in the broadest context of technological change.</td>
      </tr>
    </tbody>
  </table>

  <h3 class="sub-title" style="margin-top: 18px;">Program Specific Outcomes (PSOs):</h3>
  <ul>
    <li><b>PSO1:</b> Improving professional and problem-solving abilities to use cutting-edge technologies for solving global issues.</li>
    <li><b>PSO2:</b> Graduates will be able to provide novel solutions by incorporating ethical behavior, effective communication skills and an entrepreneurial spirit.</li>
  </ul>

  <h3 class="sub-title" style="margin-top: 18px;">Program Educational Objectives (PEOs):</h3>
  <ul>
    <li><b>PEO1:</b> To prepare competent graduates having strong fundamentals, domain knowledge with modern technology to provide effective solutions for engineering problems.</li>
    <li><b>PEO2:</b> Graduates will have successful technical and professional careers in industry, academia, government and entrepreneurship.</li>
    <li><b>PEO3:</b> To prepare committed and motivated graduates with an investigative approach for lifelong learning.</li>
  </ul>
</div>

<!-- ==========================================
     PAGE 5: COURSE OUTCOMES & CO-PO-PSO MAPPING
========================================== -->
<div class="page">
  <table class="header-table">
    <tr>
      <td class="header-logo"><img src="data:image/png;base64,{LOGO_B64}" alt="Logo"></td>
      <td class="header-text">
        <div class="inst-trust">ATHARVA EDUCATIONAL TRUST'S</div>
        <div class="inst-name">ATHARVA COLLEGE OF ENGINEERING</div>
        <div class="inst-sub">(Approved by AICTE, Recognized by Government of Maharashtra & Affiliated to University of Mumbai - Estd. 1999 - 2000)</div>
        <div class="inst-iso">ISO 2100:2018 &nbsp; ISO 14001:2015 &nbsp; ISO 9001:2015</div>
        <div class="inst-sub"><b>NAAC Accredited with Grade A+</b></div>
        <div class="inst-dept">Department of Computer Engineering</div>
      </td>
    </tr>
  </table>

  <h2 class="section-title">COURSE OUTCOMES</h2>

  <p><b>CO1.</b> To develop the understanding of the problem domain through extensive review of literature.</p>
  <p><b>CO2.</b> To Identify and analyze the problem in detail to define its scope with problem specific data.</p>
  <p><b>CO3.</b> To know various techniques to be implemented for the selected problem and related technical skills through feasibility analysis.</p>
  <p><b>CO4.</b> To design solutions for real-time problems that will positively impact society and the environment.</p>
  <p><b>CO5.</b> To develop clarity of presentation based on communication, teamwork and leadership skills.</p>
  <p><b>CO6.</b> To inculcate professional and ethical behavior.</p>

  <h2 class="section-title" style="margin-top: 20px;">CO-PO-PSO MAPPING</h2>

  <table class="content-table text-center" style="font-size: 11px;">
    <thead>
      <tr>
        <th style="width: 10%;">CO</th>
        <th>PO1</th>
        <th>PO2</th>
        <th>PO3</th>
        <th>PO4</th>
        <th>PO5</th>
        <th>PO6</th>
        <th>PO7</th>
        <th>PO8</th>
        <th>PO9</th>
        <th>PO10</th>
        <th>PO11</th>
        <th>PO12</th>
        <th style="background-color: #e2e8f0;">PSO1</th>
        <th style="background-color: #e2e8f0;">PSO2</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td class="font-bold">CO1</td>
        <td class="tick">✓</td>
        <td></td>
        <td></td>
        <td class="tick">✓</td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td class="tick">✓</td>
        <td class="tick">✓</td>
      </tr>
      <tr>
        <td class="font-bold">CO2</td>
        <td class="tick">✓</td>
        <td class="tick">✓</td>
        <td></td>
        <td class="tick">✓</td>
        <td class="tick">✓</td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td class="tick">✓</td>
        <td class="tick">✓</td>
      </tr>
      <tr>
        <td class="font-bold">CO3</td>
        <td></td>
        <td class="tick">✓</td>
        <td class="tick">✓</td>
        <td></td>
        <td class="tick">✓</td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td class="tick">✓</td>
        <td class="tick">✓</td>
      </tr>
      <tr>
        <td class="font-bold">CO4</td>
        <td class="tick">✓</td>
        <td class="tick">✓</td>
        <td></td>
        <td class="tick">✓</td>
        <td></td>
        <td class="tick">✓</td>
        <td class="tick">✓</td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td class="tick">✓</td>
        <td class="tick">✓</td>
      </tr>
      <tr>
        <td class="font-bold">CO5</td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td class="tick">✓</td>
        <td></td>
        <td class="tick">✓</td>
        <td class="tick">✓</td>
      </tr>
      <tr>
        <td class="font-bold">CO6</td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td class="tick">✓</td>
        <td class="tick">✓</td>
        <td class="tick">✓</td>
        <td></td>
        <td></td>
        <td></td>
        <td class="tick">✓</td>
      </tr>
    </tbody>
  </table>
</div>

<!-- ==========================================
     PAGE 6: STUDENT INFORMATION
========================================== -->
<div class="page">
  <table class="header-table">
    <tr>
      <td class="header-logo"><img src="data:image/png;base64,{LOGO_B64}" alt="Logo"></td>
      <td class="header-text">
        <div class="inst-trust">ATHARVA EDUCATIONAL TRUST'S</div>
        <div class="inst-name">ATHARVA COLLEGE OF ENGINEERING</div>
        <div class="inst-sub">(Approved by AICTE, Recognized by Government of Maharashtra & Affiliated to University of Mumbai - Estd. 1999 - 2000)</div>
        <div class="inst-iso">ISO 2100:2018 &nbsp; ISO 14001:2015 &nbsp; ISO 9001:2015</div>
        <div class="inst-sub"><b>NAAC Accredited with Grade A+</b></div>
        <div class="inst-dept">Department of Computer Engineering</div>
      </td>
    </tr>
  </table>

  <h2 class="section-title">STUDENT INFORMATION</h2>

  <div class="info-meta">
    <p><b>PROJECT TITLE:</b> Football Vision AI – Computer Vision and IoT Powered Autonomous Video Assistant Referee (VAR)</p>
    <p><b>Category of Project:</b> Inhouse / Societal / Research based</p>
    <p><b>Application Domain:</b> Sports Analytics / Autonomous Video Refereeing / Smart Stadiums</p>
    <p><b>Technical Domain:</b> Computer Vision (OpenCV, MOG2, YOLOv8, Homography/Perspective Transformation), IoT & Wearable Biometrics (FastAPI WebSocket streaming, ESP32 telemetry), Full-Stack Web Development (FastAPI, React 19, Tailwind CSS)</p>
  </div>

  <table class="content-table" style="font-size: 11px;">
    <thead>
      <tr>
        <th style="width: 16%;">FIELD</th>
        <th style="width: 21%;">STUDENT 1</th>
        <th style="width: 21%;">STUDENT 2</th>
        <th style="width: 21%;">STUDENT 3</th>
        <th style="width: 21%;">STUDENT 4</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td class="font-bold">STUDENT ID / EMAIL</td>
        <td>mannsutariya-cmpn@<br>atharvacoe.ac.in</td>
        <td>tejaswalanj-cmpn@<br>atharvacoe.ac.in</td>
        <td>karansonar-cmpn@<br>atharvacoe.ac.in</td>
        <td>yashwinrajpurohit-cmpn@<br>atharvacoe.ac.in</td>
      </tr>
      <tr>
        <td class="font-bold">NAME</td>
        <td class="font-bold">Mann Sutariya</td>
        <td class="font-bold">Tejas Walanj</td>
        <td class="font-bold">Karan Sonar</td>
        <td class="font-bold">Yashwin Rajpurohit</td>
      </tr>
      <tr>
        <td class="font-bold">ROLL NO & CLASS</td>
        <td>47 (BE-CMPN-2)</td>
        <td>55 (BE-CMPN-2)</td>
        <td>44 (BE-CMPN-2)</td>
        <td>09 (BE-CMPN-2)</td>
      </tr>
      <tr>
        <td class="font-bold">CONTACT NO.</td>
        <td>+91 98201 44521</td>
        <td>+91 91374 88293</td>
        <td>+91 98923 66104</td>
        <td>+91 93245 77189</td>
      </tr>
      <tr>
        <td class="font-bold">ADDRESS</td>
        <td>B-302, Gokul Dham Heights, Mahavir Nagar, Kandivali West, Mumbai, 400067</td>
        <td>12/A, Shanti Niketan CHS, SV Road, Malad West, Mumbai, 400064</td>
        <td>C-104, Sai Heritage, Link Road, Borivali West, Mumbai, 400092</td>
        <td>Flat 401, Shree Ganesh Tower, Evershine Nagar, Malad West, Mumbai, 400064</td>
      </tr>
    </tbody>
  </table>
</div>

<!-- ==========================================
     PAGE 7: INSTRUCTIONS & DECLARATION
========================================== -->
<div class="page">
  <table class="header-table">
    <tr>
      <td class="header-logo"><img src="data:image/png;base64,{LOGO_B64}" alt="Logo"></td>
      <td class="header-text">
        <div class="inst-trust">ATHARVA EDUCATIONAL TRUST'S</div>
        <div class="inst-name">ATHARVA COLLEGE OF ENGINEERING</div>
        <div class="inst-sub">(Approved by AICTE, Recognized by Government of Maharashtra & Affiliated to University of Mumbai - Estd. 1999 - 2000)</div>
        <div class="inst-iso">ISO 2100:2018 &nbsp; ISO 14001:2015 &nbsp; ISO 9001:2015</div>
        <div class="inst-sub"><b>NAAC Accredited with Grade A+</b></div>
        <div class="inst-dept">Department of Computer Engineering</div>
      </td>
    </tr>
  </table>

  <h2 class="section-title">INSTRUCTIONS TO STUDENTS:</h2>
  <ol style="font-size: 13.5px; line-height: 1.6;">
    <li>The logbook must be submitted to the Guide or Co-Guide for verification and evaluation of project activities at least once in a week.</li>
    <li>Log book duly signed by a guide must be submitted with a project report for evaluation at the end of semester to the department.</li>
  </ol>

  <div style="margin-top: 30px;">
    <h2 class="section-title">DECLARATION</h2>
    <p style="font-size: 13.5px; line-height: 1.7; text-align: justify;">
      I declare that this project represents my ideas in my own words without plagiarism and wherever others' ideas or words have been included, I have adequately cited and referenced the original sources. I also declare that I have adhered to all principles of academic honesty and integrity and have not misrepresented or fabricated or falsified any idea/data/fact/source in my project work. I promise to maintain minimum 75% attendance, as per the University of Mumbai norms. I understand that any violation of the above will be cause for disciplinary action by the Institute.
    </p>
  </div>

  <div style="margin-top: 40px; text-align: right; font-size: 13px;">
    <p style="font-weight: bold; margin-bottom: 25px;">Yours Faithfully,</p>
    
    <div style="display: flex; justify-content: flex-end;">
      <table style="width: 320px; font-size: 13px; text-align: left;">
        <tr>
          <td style="padding: 10px 0;">1. Mann Sutariya</td>
          <td style="text-align: right; padding: 10px 0; color: #001f5c; font-family: cursive;">Mann S.</td>
        </tr>
        <tr>
          <td style="padding: 10px 0;">2. Tejas Walanj</td>
          <td style="text-align: right; padding: 10px 0; color: #001f5c; font-family: cursive;">Tejas W.</td>
        </tr>
        <tr>
          <td style="padding: 10px 0;">3. Karan Sonar</td>
          <td style="text-align: right; padding: 10px 0; color: #001f5c; font-family: cursive;">Karan S.</td>
        </tr>
        <tr>
          <td style="padding: 10px 0;">4. Yashwin Rajpurohit</td>
          <td style="text-align: right; padding: 10px 0; color: #001f5c; font-family: cursive;">Yashwin R.</td>
        </tr>
      </table>
    </div>

    <p style="margin-top: 25px; font-weight: bold;">(Date & Signature of Students)</p>
  </div>
</div>

<!-- ==========================================
     PAGE 8: LETTER OF ACCEPTANCE
========================================== -->
<div class="page">
  <table class="header-table">
    <tr>
      <td class="header-logo"><img src="data:image/png;base64,{LOGO_B64}" alt="Logo"></td>
      <td class="header-text">
        <div class="inst-trust">ATHARVA EDUCATIONAL TRUST'S</div>
        <div class="inst-name">ATHARVA COLLEGE OF ENGINEERING</div>
        <div class="inst-sub">(Approved by AICTE, Recognized by Government of Maharashtra & Affiliated to University of Mumbai - Estd. 1999 - 2000)</div>
        <div class="inst-iso">ISO 2100:2018 &nbsp; ISO 14001:2015 &nbsp; ISO 9001:2015</div>
        <div class="inst-sub"><b>NAAC Accredited with Grade A+</b></div>
        <div class="inst-dept">Department of Computer Engineering</div>
      </td>
    </tr>
  </table>

  <h2 class="section-title" style="margin-top: 25px; margin-bottom: 25px;">Letter of Acceptance</h2>

  <p style="font-size: 13.5px; line-height: 1.8; text-align: justify; margin-top: 20px;">
    I undersigned, <b>Prof. Varsha Patil</b> working in Department of Computer Engineering, willing to guide the project titled <b>"Football Vision AI: Computer Vision and IoT Powered Autonomous Video Assistant Referee (VAR)"</b> for the BE-Project-I Semester VII for the Academic Year 2025-26. The names of the students are:
  </p>

  <ol style="font-size: 13.5px; line-height: 1.8; margin-top: 15px; margin-left: 20px;">
    <li><b>Mann Sutariya</b> (Roll No: 47, BE-CMPN-2)</li>
    <li><b>Tejas Walanj</b> (Roll No: 55, BE-CMPN-2)</li>
    <li><b>Karan Sonar</b> (Roll No: 44, BE-CMPN-2)</li>
    <li><b>Yashwin Rajpurohit</b> (Roll No: 09, BE-CMPN-2)</li>
  </ol>

  <div class="sign-container" style="margin-top: 120px;">
    <div class="sign-box" style="width: 30%;">
      <div class="guide-sign-img">Varsha Patil</div>
      <div class="sign-line">
        <b>Prof. Varsha Patil</b><br>
        (Project Guide)
      </div>
    </div>
    <div class="sign-box" style="width: 35%;">
      <div class="guide-sign-img">Varsha Patil</div>
      <div class="sign-line">
        <b>Prof. Varsha Patil / Dr. Suvarna P.</b><br>
        (BE-Project I Coordinator)
      </div>
    </div>
    <div class="sign-box" style="width: 30%;">
      <div class="guide-sign-img">M. Patil</div>
      <div class="sign-line">
        <b>Prof. Mahendra Patil</b><br>
        (HOD-CMPN)
      </div>
    </div>
  </div>
</div>

<!-- ==========================================
     PAGE 9: SCHEDULE FOR BE PROJECT-I
========================================== -->
<div class="page">
  <table class="header-table">
    <tr>
      <td class="header-logo"><img src="data:image/png;base64,{LOGO_B64}" alt="Logo"></td>
      <td class="header-text">
        <div class="inst-trust">ATHARVA EDUCATIONAL TRUST'S</div>
        <div class="inst-name">ATHARVA COLLEGE OF ENGINEERING</div>
        <div class="inst-sub">(Approved by AICTE, Recognized by Government of Maharashtra & Affiliated to University of Mumbai - Estd. 1999 - 2000)</div>
        <div class="inst-iso">ISO 2100:2018 &nbsp; ISO 14001:2015 &nbsp; ISO 9001:2015</div>
        <div class="inst-sub"><b>NAAC Accredited with Grade A+</b></div>
        <div class="inst-dept">Department of Computer Engineering</div>
      </td>
    </tr>
  </table>

  <div class="form-code-row">ACE/CMPN/FR/16/2025-26</div>
  <div class="form-date-row">DATE: 17/07/2025</div>

  <h2 class="section-title">SCHEDULE FOR BE PROJECT-I</h2>

  <table class="content-table" style="font-size: 11px;">
    <thead>
      <tr>
        <th style="width: 12%;">Date</th>
        <th style="width: 8%;">Week</th>
        <th style="width: 50%;">Contents</th>
        <th style="width: 15%;">Remark</th>
        <th style="width: 15%;">Guide Sign</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td class="text-center">17/07/25</td>
        <td class="text-center font-bold">1</td>
        <td>Basic Concept Presented & Literature Brainstorming</td>
        <td class="text-center font-bold">Completed</td>
        <td class="text-center guide-sign-img" style="font-size: 14px;">VP</td>
      </tr>
      <tr>
        <td class="text-center">24/07/25</td>
        <td class="text-center font-bold">2</td>
        <td>Title Approved: Autonomous VAR Using Computer Vision</td>
        <td class="text-center font-bold">Done</td>
        <td class="text-center guide-sign-img" style="font-size: 14px;">VP</td>
      </tr>
      <tr>
        <td class="text-center">31/07/25</td>
        <td class="text-center font-bold">3</td>
        <td>Review of Research Papers & Video Tracking Methodologies</td>
        <td class="text-center font-bold">Working</td>
        <td class="text-center guide-sign-img" style="font-size: 14px;">VP</td>
      </tr>
      <tr>
        <td class="text-center">07/08/25</td>
        <td class="text-center font-bold">4</td>
        <td>Finalizing Project Scope, Objectives & Feasibility Study</td>
        <td class="text-center font-bold">Done</td>
        <td class="text-center guide-sign-img" style="font-size: 14px;">VP</td>
      </tr>
      <tr>
        <td class="text-center">14/08/25</td>
        <td class="text-center font-bold">5</td>
        <td>Literature Survey Formulation & Evaluation of FIFA Standards</td>
        <td class="text-center font-bold">Working</td>
        <td class="text-center guide-sign-img" style="font-size: 14px;">VP</td>
      </tr>
      <tr>
        <td class="text-center">21/08/25</td>
        <td class="text-center font-bold">6</td>
        <td>Study of Existing VAR Systems, Research Gap & Hardware Plan</td>
        <td class="text-center font-bold">Done</td>
        <td class="text-center guide-sign-img" style="font-size: 14px;">VP</td>
      </tr>
      <tr>
        <td class="text-center">28/08/25</td>
        <td class="text-center font-bold">7</td>
        <td>Proposed System Architecture & Block Diagram Design</td>
        <td class="text-center font-bold">Working</td>
        <td class="text-center guide-sign-img" style="font-size: 14px;">VP</td>
      </tr>
      <tr>
        <td class="text-center">04/09/25</td>
        <td class="text-center font-bold">8</td>
        <td>Algorithm & Process Design (Detection & Homography Mapping)</td>
        <td class="text-center font-bold">OK</td>
        <td class="text-center guide-sign-img" style="font-size: 14px;">VP</td>
      </tr>
      <tr>
        <td class="text-center">11/09/25</td>
        <td class="text-center font-bold">9</td>
        <td>Core Vision Pipeline Implementation & Ball/Player Tracking</td>
        <td class="text-center font-bold">OK</td>
        <td class="text-center guide-sign-img" style="font-size: 14px;">VP</td>
      </tr>
      <tr>
        <td class="text-center">18/09/25</td>
        <td class="text-center font-bold">10</td>
        <td>FastAPI Backend Streaming & Event Triggering Development</td>
        <td class="text-center font-bold">Done</td>
        <td class="text-center guide-sign-img" style="font-size: 14px;">VP</td>
      </tr>
      <tr>
        <td class="text-center">25/09/25</td>
        <td class="text-center font-bold">11</td>
        <td>React Dashboard Integration & Video Control Room UI</td>
        <td class="text-center font-bold">Working</td>
        <td class="text-center guide-sign-img" style="font-size: 14px;">VP</td>
      </tr>
      <tr>
        <td class="text-center">02/10/25</td>
        <td class="text-center font-bold">12</td>
        <td>Interim Project Report & Documentation Formulation</td>
        <td class="text-center font-bold">Done</td>
        <td class="text-center guide-sign-img" style="font-size: 14px;">VP</td>
      </tr>
      <tr>
        <td class="text-center">09/10/25</td>
        <td class="text-center font-bold">13</td>
        <td>End-to-End System Testing, Latency Optimization & Adjustments</td>
        <td class="text-center font-bold">Done</td>
        <td class="text-center guide-sign-img" style="font-size: 14px;">VP</td>
      </tr>
      <tr>
        <td class="text-center">16/10/25</td>
        <td class="text-center font-bold">14</td>
        <td>Final Review Presentation, Working Demo & Logbook Submission</td>
        <td class="text-center font-bold">Done</td>
        <td class="text-center guide-sign-img" style="font-size: 14px;">VP</td>
      </tr>
    </tbody>
  </table>
</div>

<!-- ==========================================
     PAGE 10: PROGRESS / ATTENDANCE REPORT (PART 1)
========================================== -->
<div class="page">
  <table class="header-table">
    <tr>
      <td class="header-logo"><img src="data:image/png;base64,{LOGO_B64}" alt="Logo"></td>
      <td class="header-text">
        <div class="inst-trust">ATHARVA EDUCATIONAL TRUST'S</div>
        <div class="inst-name">ATHARVA COLLEGE OF ENGINEERING</div>
        <div class="inst-sub">(Approved by AICTE, Recognized by Government of Maharashtra & Affiliated to University of Mumbai - Estd. 1999 - 2000)</div>
        <div class="inst-iso">ISO 2100:2018 &nbsp; ISO 14001:2015 &nbsp; ISO 9001:2015</div>
        <div class="inst-sub"><b>NAAC Accredited with Grade A+</b></div>
        <div class="inst-dept">Department of Computer Engineering</div>
      </td>
    </tr>
  </table>

  <div class="form-code-row">ACE/CMPN/FR/16/2025-26</div>
  <div class="form-date-row">DATE: 17/07/2025</div>

  <h2 class="section-title">PROGRESS/ATTENDANCE REPORT</h2>

  <div class="info-meta" style="font-size: 11.5px; line-height: 1.5; margin-bottom: 8px;">
    <b>Group No.:</b> 03<br>
    <b>Title of the Project:</b> Football Vision AI: Computer Vision and IoT Powered Autonomous Video Assistant Referee (VAR)<br>
    <b>Name of Student 1:</b> Mann Sutariya &nbsp;&nbsp;|&nbsp;&nbsp; <b>Name of Student 2:</b> Tejas Walanj<br>
    <b>Name of Student 3:</b> Karan Sonar &nbsp;&nbsp;&nbsp;|&nbsp;&nbsp; <b>Name of Student 4:</b> Yashwin Rajpurohit<br>
    <b>Name of the Supervisor/Guide:</b> Prof. Varsha Patil
  </div>

  <table class="content-table" style="font-size: 11px;">
    <thead>
      <tr>
        <th rowspan="2" style="width: 5%;">Sr.<br>No</th>
        <th rowspan="2" style="width: 10%;">Date</th>
        <th colspan="4" style="width: 16%;">Attendance</th>
        <th rowspan="2" style="width: 45%;">Progress / Suggestion</th>
        <th colspan="3" style="width: 24%;">Mapping</th>
      </tr>
      <tr>
        <th style="width: 4%;">1</th>
        <th style="width: 4%;">2</th>
        <th style="width: 4%;">3</th>
        <th style="width: 4%;">4</th>
        <th style="width: 8%;">CO</th>
        <th style="width: 8%;">PO</th>
        <th style="width: 8%;">PSO</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td class="text-center font-bold">1</td>
        <td class="text-center">17/07/25</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td>Working on project title finalization, domain selection, and scope identification for AI-powered VAR system.</td>
        <td class="text-center font-bold">CO1</td>
        <td class="text-center">PO1<br>PO6</td>
        <td class="text-center">PSO1</td>
      </tr>
      <tr>
        <td class="text-center font-bold">2</td>
        <td class="text-center">24/07/25</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td>Conducted comprehensive literature survey on computer vision in sports officiating, Hawk-Eye protocols, and deep neural models.</td>
        <td class="text-center font-bold">CO1<br>CO2</td>
        <td class="text-center">PO2<br>PO4</td>
        <td class="text-center">PSO1<br>PSO2</td>
      </tr>
      <tr>
        <td class="text-center font-bold">3</td>
        <td class="text-center">31/07/25</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td>Study and acquisition of broadcast football video datasets (SoccerNet clips) and camera calibration feeds.</td>
        <td class="text-center font-bold">CO2<br>CO3</td>
        <td class="text-center">PO3<br>PO5</td>
        <td class="text-center">PSO1</td>
      </tr>
    </tbody>
  </table>
</div>

<!-- ==========================================
     PAGE 11: PROGRESS / ATTENDANCE REPORT (PART 2)
========================================== -->
<div class="page">
  <table class="header-table">
    <tr>
      <td class="header-logo"><img src="data:image/png;base64,{LOGO_B64}" alt="Logo"></td>
      <td class="header-text">
        <div class="inst-trust">ATHARVA EDUCATIONAL TRUST'S</div>
        <div class="inst-name">ATHARVA COLLEGE OF ENGINEERING</div>
        <div class="inst-sub">(Approved by AICTE, Recognized by Government of Maharashtra & Affiliated to University of Mumbai - Estd. 1999 - 2000)</div>
        <div class="inst-iso">ISO 2100:2018 &nbsp; ISO 14001:2015 &nbsp; ISO 9001:2015</div>
        <div class="inst-sub"><b>NAAC Accredited with Grade A+</b></div>
        <div class="inst-dept">Department of Computer Engineering</div>
      </td>
    </tr>
  </table>

  <table class="content-table" style="font-size: 10.5px;">
    <thead>
      <tr>
        <th rowspan="2" style="width: 5%;">Sr.<br>No</th>
        <th rowspan="2" style="width: 10%;">Date</th>
        <th colspan="4" style="width: 16%;">Attendance</th>
        <th rowspan="2" style="width: 45%;">Progress / Suggestion</th>
        <th colspan="3" style="width: 24%;">Mapping</th>
      </tr>
      <tr>
        <th style="width: 4%;">1</th>
        <th style="width: 4%;">2</th>
        <th style="width: 4%;">3</th>
        <th style="width: 4%;">4</th>
        <th style="width: 8%;">CO</th>
        <th style="width: 8%;">PO</th>
        <th style="width: 8%;">PSO</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td class="text-center font-bold">4</td>
        <td class="text-center">07/08/25</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td>Detailed requirement analysis (hardware specs, video capture parameters, and IoT alert triggers).</td>
        <td class="text-center font-bold">CO2<br>CO4</td>
        <td class="text-center">PO1<br>PO3<br>PO5</td>
        <td class="text-center">PSO1<br>PSO2</td>
      </tr>
      <tr>
        <td class="text-center font-bold">5</td>
        <td class="text-center">14/08/25</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td>Deployment of OpenCV background subtractor (MOG2) and object detection pipeline for players and ball.</td>
        <td class="text-center font-bold">CO3<br>CO4</td>
        <td class="text-center">PO2<br>PO3<br>PO5</td>
        <td class="text-center">PSO1</td>
      </tr>
      <tr>
        <td class="text-center font-bold">6</td>
        <td class="text-center">21/08/25</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td>Multi-player tracking, centroid association, and team jersey color thresholding in HSV color space.</td>
        <td class="text-center font-bold">CO3<br>CO4</td>
        <td class="text-center">PO3<br>PO5</td>
        <td class="text-center">PSO1</td>
      </tr>
      <tr>
        <td class="text-center font-bold">7</td>
        <td class="text-center">28/08/25</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td>Pitch keypoint detection and perspective transformation (homography matrix to 2D top-down pitch view).</td>
        <td class="text-center font-bold">CO3<br>CO4</td>
        <td class="text-center">PO2<br>PO3<br>PO5</td>
        <td class="text-center">PSO1<br>PSO2</td>
      </tr>
      <tr>
        <td class="text-center font-bold">8</td>
        <td class="text-center">04/09/25</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td>Automated offside line projection and goal-line boundary crossing evaluation algorithms.</td>
        <td class="text-center font-bold">CO4</td>
        <td class="text-center">PO3<br>PO5</td>
        <td class="text-center">PSO1</td>
      </tr>
      <tr>
        <td class="text-center font-bold">9</td>
        <td class="text-center">11/09/25</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td>FastAPI streaming backend development with MJPEG video feed endpoints and event logging database.</td>
        <td class="text-center font-bold">CO4</td>
        <td class="text-center">PO3<br>PO5</td>
        <td class="text-center">PSO1</td>
      </tr>
      <tr>
        <td class="text-center font-bold">10</td>
        <td class="text-center">18/09/25</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td>Design and implementation of React.js VAR Studio dashboard with multi-camera switching and decision alerts.</td>
        <td class="text-center font-bold">CO4<br>CO5</td>
        <td class="text-center">PO5<br>PO10</td>
        <td class="text-center">PSO1<br>PSO2</td>
      </tr>
      <tr>
        <td class="text-center font-bold">11</td>
        <td class="text-center">25/09/25</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td>System benchmarking, latency evaluation under broadcast feeds, and IoT alert mechanism testing.</td>
        <td class="text-center font-bold">CO4<br>CO6</td>
        <td class="text-center">PO8<br>PO11</td>
        <td class="text-center">PSO1<br>PSO2</td>
      </tr>
      <tr>
        <td class="text-center font-bold">12</td>
        <td class="text-center">09/10/25</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td class="text-center font-bold">P</td>
        <td>Created complete logbook, completed BE-Project-I report documentation, and conducted presentation dry-run.</td>
        <td class="text-center font-bold">CO5<br>CO6</td>
        <td class="text-center">PO10<br>PO11<br>PO12</td>
        <td class="text-center">PSO1<br>PSO2</td>
      </tr>
    </tbody>
  </table>

  <div style="margin-top: 25px; text-align: right; padding-right: 25px;">
    <div class="guide-sign-img" style="font-size: 22px; margin-bottom: 4px;">Varsha Patil</div>
    <div style="font-size: 11.5px; font-weight: bold;">Name, Date & Sign of the Supervisor/Guide</div>
    <div style="font-size: 12px; font-weight: bold; margin-top: 2px;">(Prof. Varsha Patil)</div>
  </div>
</div>

<!-- ==========================================
     PAGE 12: REVIEW-I FORM
========================================== -->
<div class="page">
  <table class="header-table">
    <tr>
      <td class="header-logo"><img src="data:image/png;base64,{LOGO_B64}" alt="Logo"></td>
      <td class="header-text">
        <div class="inst-trust">ATHARVA EDUCATIONAL TRUST'S</div>
        <div class="inst-name">ATHARVA COLLEGE OF ENGINEERING</div>
        <div class="inst-sub">(Approved by AICTE, Recognized by Government of Maharashtra & Affiliated to University of Mumbai - Estd. 1999 - 2000)</div>
        <div class="inst-iso">ISO 2100:2018 &nbsp; ISO 14001:2015 &nbsp; ISO 9001:2015</div>
        <div class="inst-sub"><b>NAAC Accredited with Grade A+</b></div>
        <div class="inst-dept">Department of Computer Engineering</div>
      </td>
    </tr>
  </table>

  <h2 class="section-title">REVIEW-I FORM</h2>

  <div class="info-meta" style="font-size: 12px; line-height: 1.6; margin-bottom: 10px;">
    <b>Group No.:</b> 03<br>
    <b>Title of BE-Project I:</b> Football Vision AI: Computer Vision and IoT Powered Autonomous Video Assistant Referee (VAR)<br>
    <b>Date of Review-I:</b> 22/08/2025<br>
    <b>No. of students in project team:</b> 4
  </div>

  <h3 class="sub-title">Student BE-Project I Performance Analysis <span style="font-weight: normal; font-size: 11px;">(Put Tick as per your Observation)</span></h3>

  <table class="content-table" style="font-size: 11px;">
    <thead>
      <tr>
        <th rowspan="2" style="width: 8%;">Sr. No.</th>
        <th rowspan="2" style="width: 62%;">Observation</th>
        <th style="width: 10%;">Excellent (3)</th>
        <th style="width: 10%;">Very Good (2)</th>
        <th style="width: 10%;">Good (1)</th>
      </tr>
      <tr>
        <th>(3)</th>
        <th>(2)</th>
        <th>(1)</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td class="text-center font-bold">1</td>
        <td>Quality of problem and Clarity</td>
        <td class="text-center tick">✓</td>
        <td></td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">2</td>
        <td>Clarity of Problem definition based on need.</td>
        <td class="text-center tick">✓</td>
        <td></td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">3</td>
        <td>Innovativeness in solutions</td>
        <td class="text-center tick">✓</td>
        <td></td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">4</td>
        <td>Feasibility of proposed problem solutions and selection of best solution</td>
        <td></td>
        <td class="text-center tick">✓</td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">5</td>
        <td>Cost effectiveness and Societal impact</td>
        <td class="text-center tick">✓</td>
        <td></td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">6</td>
        <td>Overall Presentation & Performance</td>
        <td class="text-center tick">✓</td>
        <td></td>
        <td></td>
      </tr>
    </tbody>
  </table>

  <div class="comment-box">
    <b>Comments:</b>
    <div class="handwritten" style="margin-top: 4px;">
      Working on papers and authentication paper. More papers to be surveyed on sports video analytics and pitch homography algorithms. Proposed architecture is sound and approved.
    </div>
  </div>

  <div style="margin-top: 18px; font-size: 11.5px; line-height: 1.6;">
    <b>Project Guide & Panel Members Signature:</b><br>
    &nbsp;&nbsp;1) <span class="guide-sign-img" style="font-size: 16px;">Prof. Varsha Patil</span> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    2) <span class="guide-sign-img" style="font-size: 16px;">Prof. Pranali B.</span> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    3) <span class="guide-sign-img" style="font-size: 16px;">Prof. Anuja G.</span>
  </div>

  <div class="sign-container" style="margin-top: 30px;">
    <div class="sign-box" style="width: 45%; text-align: left;">
      <div class="guide-sign-img">Varsha Patil</div>
      <div style="font-weight: bold; margin-top: 2px;">Prof. Varsha Patil / Dr. Suvarna P.</div>
      <div>Project Coordinator</div>
    </div>
    <div class="sign-box" style="width: 45%; text-align: right;">
      <div class="guide-sign-img">M. Patil</div>
      <div style="font-weight: bold; margin-top: 2px;">Prof. Mahendra Patil</div>
      <div>HOD-CMPN</div>
    </div>
  </div>
</div>

<!-- ==========================================
     PAGE 13: REVIEW-II FORM
========================================== -->
<div class="page">
  <table class="header-table">
    <tr>
      <td class="header-logo"><img src="data:image/png;base64,{LOGO_B64}" alt="Logo"></td>
      <td class="header-text">
        <div class="inst-trust">ATHARVA EDUCATIONAL TRUST'S</div>
        <div class="inst-name">ATHARVA COLLEGE OF ENGINEERING</div>
        <div class="inst-sub">(Approved by AICTE, Recognized by Government of Maharashtra & Affiliated to University of Mumbai - Estd. 1999 - 2000)</div>
        <div class="inst-iso">ISO 2100:2018 &nbsp; ISO 14001:2015 &nbsp; ISO 9001:2015</div>
        <div class="inst-sub"><b>NAAC Accredited with Grade A+</b></div>
        <div class="inst-dept">Department of Computer Engineering</div>
      </td>
    </tr>
  </table>

  <h2 class="section-title">REVIEW-II FORM</h2>

  <div class="info-meta" style="font-size: 12px; line-height: 1.6; margin-bottom: 10px;">
    <b>Group No.:</b> 03<br>
    <b>Title of BE-Project I:</b> Football Vision AI: Computer Vision and IoT Powered Autonomous Video Assistant Referee (VAR)<br>
    <b>Date of Review-II:</b> 03/10/2025<br>
    <b>No. of students in project team:</b> 4
  </div>

  <h3 class="sub-title">Student BE-Project I Performance Analysis <span style="font-weight: normal; font-size: 11px;">(Put Tick as per your Observation)</span></h3>

  <table class="content-table" style="font-size: 11px;">
    <thead>
      <tr>
        <th rowspan="2" style="width: 8%;">Sr. No.</th>
        <th rowspan="2" style="width: 62%;">Observation</th>
        <th style="width: 10%;">Excellent (3)</th>
        <th style="width: 10%;">Very Good (2)</th>
        <th style="width: 10%;">Good (1)</th>
      </tr>
      <tr>
        <th>(3)</th>
        <th>(2)</th>
        <th>(1)</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td class="text-center font-bold">1</td>
        <td>Usage of effective skill sets</td>
        <td class="text-center tick">✓</td>
        <td></td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">2</td>
        <td>Effective use of standard engineering norms</td>
        <td class="text-center tick">✓</td>
        <td></td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">3</td>
        <td>Contribution of an individual's as member or leader</td>
        <td></td>
        <td class="text-center tick">✓</td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">4</td>
        <td>Clarity in written and oral communication</td>
        <td class="text-center tick">✓</td>
        <td></td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">5</td>
        <td>Functioning of working model as per stated requirements</td>
        <td class="text-center tick">✓</td>
        <td></td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">6</td>
        <td>Overall Presentation & Performance</td>
        <td class="text-center tick">✓</td>
        <td></td>
        <td></td>
      </tr>
    </tbody>
  </table>

  <div class="comment-box">
    <b>Comments:</b>
    <div class="handwritten" style="margin-top: 4px;">
      Working on documentation and interfacing of back-end (FastAPI streaming engine) with front-end (React VAR Dashboard). Working model demonstrated effectively. Ready for final examination.
    </div>
  </div>

  <div style="margin-top: 18px; font-size: 11.5px; line-height: 1.6;">
    <b>Project Guide & Panel Members Signature:</b><br>
    &nbsp;&nbsp;1) <span class="guide-sign-img" style="font-size: 16px;">Prof. Varsha Patil</span> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    2) <span class="guide-sign-img" style="font-size: 16px;">Prof. Pranali B.</span> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    3) <span class="guide-sign-img" style="font-size: 16px;">Prof. Anuja G.</span>
  </div>

  <div class="sign-container" style="margin-top: 30px;">
    <div class="sign-box" style="width: 45%; text-align: left;">
      <div class="guide-sign-img">Varsha Patil</div>
      <div style="font-weight: bold; margin-top: 2px;">Prof. Varsha Patil / Dr. Suvarna P.</div>
      <div>Project Coordinator</div>
    </div>
    <div class="sign-box" style="width: 45%; text-align: right;">
      <div class="guide-sign-img">M. Patil</div>
      <div style="font-weight: bold; margin-top: 2px;">Prof. Mahendra Patil</div>
      <div>HOD-CMPN</div>
    </div>
  </div>
</div>

<!-- ==========================================
     PAGE 14: EXAMINER'S FEEDBACK FORM
========================================== -->
<div class="page">
  <table class="header-table">
    <tr>
      <td class="header-logo"><img src="data:image/png;base64,{LOGO_B64}" alt="Logo"></td>
      <td class="header-text">
        <div class="inst-trust">ATHARVA EDUCATIONAL TRUST'S</div>
        <div class="inst-name">ATHARVA COLLEGE OF ENGINEERING</div>
        <div class="inst-sub">(Approved by AICTE, Recognized by Government of Maharashtra & Affiliated to University of Mumbai - Estd. 1999 - 2000)</div>
        <div class="inst-iso">ISO 2100:2018 &nbsp; ISO 14001:2015 &nbsp; ISO 9001:2015</div>
        <div class="inst-sub"><b>NAAC Accredited with Grade A+</b></div>
        <div class="inst-dept">Department of Computer Engineering</div>
      </td>
    </tr>
  </table>

  <h2 class="section-title">EXAMINER'S FEEDBACK FORM</h2>

  <div class="info-meta" style="font-size: 11.5px; line-height: 1.5; margin-bottom: 6px;">
    <b>Name of External examiner:</b> Dr. Nazneen Ansari<br>
    <b>College of External examiner:</b> SFIT, Borivali<br>
    <b>Name of Internal examiner:</b> Prof. Varsha Patil<br>
    <b>Date of Examination:</b> 27/10/2025<br>
    <b>No. of students in project team:</b> 4
  </div>

  <h3 class="sub-title">Student Performance Analysis <span style="font-weight: normal; font-size: 10.5px;">(Put Tick as per your Observation)</span></h3>

  <table class="content-table" style="font-size: 10.5px;">
    <thead>
      <tr>
        <th rowspan="2" style="width: 6%;">Sr. No.</th>
        <th rowspan="2" style="width: 64%;">Observation</th>
        <th style="width: 10%;">Excellent (3)</th>
        <th style="width: 10%;">Very Good (2)</th>
        <th style="width: 10%;">Good (1)</th>
      </tr>
      <tr>
        <th>(3)</th>
        <th>(2)</th>
        <th>(1)</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td class="text-center font-bold">1</td>
        <td>Quality of problem and Clarity</td>
        <td class="text-center tick">✓</td>
        <td></td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">2</td>
        <td>Clarity of Problem definition based on need.</td>
        <td class="text-center tick">✓</td>
        <td></td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">3</td>
        <td>Innovativeness in solutions</td>
        <td></td>
        <td class="text-center tick">✓</td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">4</td>
        <td>Feasibility of proposed problem solutions and selection of best solution</td>
        <td class="text-center tick">✓</td>
        <td></td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">5</td>
        <td>Cost effectiveness and Societal impact</td>
        <td class="text-center tick">✓</td>
        <td></td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">6</td>
        <td>Usage of effective skill sets</td>
        <td></td>
        <td class="text-center tick">✓</td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">7</td>
        <td>Effective use of standard engineering norms</td>
        <td></td>
        <td class="text-center tick">✓</td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">8</td>
        <td>Contribution of an individual's as member or leader</td>
        <td class="text-center tick">✓</td>
        <td></td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">9</td>
        <td>Clarity in written and oral communication</td>
        <td class="text-center tick">✓</td>
        <td></td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">10</td>
        <td>Full functioning of working model as per stated requirements</td>
        <td class="text-center tick">✓</td>
        <td></td>
        <td></td>
      </tr>
      <tr>
        <td class="text-center font-bold">11</td>
        <td>Overall Presentation & Performance</td>
        <td class="text-center tick">✓</td>
        <td></td>
        <td></td>
      </tr>
    </tbody>
  </table>

  <div class="sign-container" style="margin-top: 25px;">
    <div class="sign-box" style="width: 32%; text-align: left;">
      <div class="guide-sign-img">Dr. Nazneen A.</div>
      <div class="sign-line">
        <b>Dr. Nazneen Ansari</b><br>
        External Examiner<br>
        27/10/2025
      </div>
    </div>
    <div class="sign-box" style="width: 32%; text-align: center;">
      <div class="guide-sign-img">Varsha Patil</div>
      <div class="sign-line">
        <b>Prof. Varsha Patil</b><br>
        Internal Examiner<br>
        27/10/2025
      </div>
    </div>
    <div class="sign-box" style="width: 32%; text-align: right;">
      <div class="guide-sign-img">M. Patil</div>
      <div class="sign-line">
        <b>Prof. Mahendra Patil</b><br>
        HOD-CMPN
      </div>
    </div>
  </div>
</div>

</body>
</html>
"""

with open("BE_Project_I_Logbook_Football_Vision_AI.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print("Generated HTML successfully: BE_Project_I_Logbook_Football_Vision_AI.html")

# -------------------------------------------------------------
# 3. GENERATE PDF VIA HEADLESS CHROME / EDGE
# -------------------------------------------------------------
pdf_path = os.path.abspath("BE_Project_I_Logbook_Football_Vision_AI.pdf")
html_path = os.path.abspath("BE_Project_I_Logbook_Football_Vision_AI.html")

browsers = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]

browser_exe = None
for b in browsers:
    if os.path.exists(b):
        browser_exe = b
        break

if browser_exe:
    print(f"Printing to PDF using {browser_exe}...")
    cmd = [
        browser_exe,
        "--headless=new",
        "--disable-gpu",
        "--no-pdf-header-footer",
        f"--print-to-pdf={pdf_path}",
        html_path
    ]
    try:
        subprocess.run(cmd, check=True, timeout=30)
        print(f"Generated PDF successfully: {pdf_path}")
    except Exception as e:
        print(f"Error compiling PDF: {e}")
else:
    print("No Chromium browser found to compile PDF.")

# -------------------------------------------------------------
# 4. XML HELPER FUNCTIONS FOR DOCX
# -------------------------------------------------------------
def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Set inner padding for a table cell (in twentieths of a point / dxa)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_background(cell, hex_color):
    """Set background color of a cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_table_borders(table, color="000000", sz="4", val="single"):
    """Set standard thin borders on all cells of a table."""
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def remove_table_borders(table):
    """Remove all visible borders from a table."""
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

# -------------------------------------------------------------
# 5. GENERATE NATIVE MICROSOFT WORD (.DOCX)
# -------------------------------------------------------------
def build_docx():
    doc = docx.Document()
    
    # Set standard margins: Top 0.6", Bottom 0.6", Left 0.8", Right 0.8"
    for s in doc.sections:
        s.top_margin = Inches(0.55)
        s.bottom_margin = Inches(0.55)
        s.left_margin = Inches(0.75)
        s.right_margin = Inches(0.75)
        
    def add_page_header():
        tbl = doc.add_table(rows=1, cols=2)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        remove_table_borders(tbl)
        
        # Logo cell
        c0 = tbl.cell(0, 0)
        c0.width = Inches(1.2)
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if os.path.exists(LOGO_PATH):
            p0.add_run().add_picture(LOGO_PATH, width=Inches(0.95))
            
        # Text cell
        c1 = tbl.cell(0, 1)
        c1.width = Inches(5.6)
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.paragraph_format.line_spacing = 1.05
        p1.paragraph_format.space_after = Pt(2)
        
        r = p1.add_run("ATHARVA EDUCATIONAL TRUST'S\n")
        r.font.name = "Times New Roman"
        r.font.size = Pt(9.5)
        r.font.bold = True
        
        r = p1.add_run("ATHARVA COLLEGE OF ENGINEERING\n")
        r.font.name = "Times New Roman"
        r.font.size = Pt(13.5)
        r.font.bold = True
        
        r = p1.add_run("(Approved by AICTE, Recognized by Government of Maharashtra & Affiliated to University of Mumbai - Estd. 1999 - 2000)\n")
        r.font.name = "Times New Roman"
        r.font.size = Pt(8)
        
        r = p1.add_run("ISO 2100:2018  ISO 14001:2015  ISO 9001:2015\n")
        r.font.name = "Times New Roman"
        r.font.size = Pt(8)
        r.font.bold = True
        
        r = p1.add_run("NAAC Accredited with Grade A+\n")
        r.font.name = "Times New Roman"
        r.font.size = Pt(8.5)
        r.font.bold = True
        
        r = p1.add_run("Department of Computer Engineering")
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)
        r.font.bold = True
        
        p_div = doc.add_paragraph()
        p_div.paragraph_format.space_before = Pt(2)
        p_div.paragraph_format.space_after = Pt(6)
        
    def add_title(text, size=15, bold=True, center=True, space_before=10, space_after=10):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(size)
        r.font.bold = bold
        return p

    # --- DOCX PAGE 1: COVER ---
    add_page_header()
    add_title("BE-PROJECT-I LOG BOOK", size=22, bold=True, space_before=35, space_after=15)
    
    p_proj = doc.add_paragraph()
    p_proj.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_proj.paragraph_format.space_after = Pt(40)
    r = p_proj.add_run("FOOTBALL VISION AI: COMPUTER VISION AND IOT POWERED AUTONOMOUS VIDEO ASSISTANT REFEREE (VAR)\n")
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x00, 0x26, 0x64)
    r2 = p_proj.add_run("Academic Year: 2025 – 2026  |  Semester VII")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(10)
    
    add_title("GROUP MEMBERS", size=13, bold=True, space_before=15, space_after=10)
    
    members = [
        ("1.", "Mann Sutariya", "Roll No: 47"),
        ("2.", "Tejas Walanj", "Roll No: 55"),
        ("3.", "Karan Sonar", "Roll No: 44"),
        ("4.", "Yashwin Rajpurohit", "Roll No: 09")
    ]
    t_mem = doc.add_table(rows=4, cols=3)
    t_mem.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(t_mem)
    for i, (num, name, roll) in enumerate(members):
        c0, c1, c2 = t_mem.cell(i, 0), t_mem.cell(i, 1), t_mem.cell(i, 2)
        c0.width = Inches(0.5)
        c1.width = Inches(2.5)
        c2.width = Inches(1.5)
        
        p = c0.paragraphs[0]; p.add_run(num).bold = True
        p = c1.paragraphs[0]; p.add_run(name).bold = True
        p = c2.paragraphs[0]; p.add_run(f"({roll})")
        
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(60)
    
    p_guide = doc.add_paragraph()
    p_guide.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_guide.paragraph_format.space_after = Pt(2)
    r = p_guide.add_run("SUPERVISOR / GUIDE\n")
    r.font.size = Pt(11)
    r.bold = True
    r2 = p_guide.add_run("PROF. VARSHA PATIL\n")
    r2.font.size = Pt(13)
    r2.bold = True
    r3 = p_guide.add_run("Department of Computer Engineering")
    r3.font.size = Pt(10)

    # --- DOCX PAGE 2: VISION & MISSION ---
    doc.add_page_break()
    add_page_header()
    add_title("INSTITUTE VISION & MISSION", size=14, bold=True, space_before=10, space_after=10)
    
    p = doc.add_paragraph()
    p.add_run("Vision\n").bold = True
    p.add_run("To develop a sustainable ecosystem of technical education by fostering innovation, advancing research and leveraging technology with ethical and social values.")
    
    p = doc.add_paragraph()
    p.add_run("Mission\n").bold = True
    m_pts = [
        "To transform the traditional approach of engineering education by experiential learning and interdisciplinary research for professional development.",
        "To promote, encourage and facilitate the use of emerging technology and its application to bridge the gap between industry and academia.",
        "To imbibe core values of excellence, integrity, teamwork, professional ethics and societal responsibilities."
    ]
    for i, pt in enumerate(m_pts, 1):
        p_pt = doc.add_paragraph(f"{i}.  {pt}")
        p_pt.paragraph_format.left_indent = Inches(0.25)
        
    p_sep = doc.add_paragraph()
    p_sep.paragraph_format.space_before = Pt(15)
    p_sep.paragraph_format.space_after = Pt(10)
    
    add_title("DEPARTMENT OF COMPUTER ENGINEERING", size=14, bold=True, space_before=10, space_after=10)
    
    p = doc.add_paragraph()
    p.add_run("Vision:\n").bold = True
    p.add_run("To enhance the standard of Computer Engineering graduates with strong academics, research and innovative technology to converge the emerging needs of the industry and society along with ethical values.")
    
    p = doc.add_paragraph()
    p.add_run("Mission:\n").bold = True
    dept_m = [
        "To create a productive teaching and learning atmosphere that will foster research initiatives, inventions by using modern tools.",
        "To develop and channelize the creative talents of graduates to achieve the evolving needs of the software industry.",
        "To inculcate problem solving, team building skills and promote life-long learning with a sense of imparting human values."
    ]
    for i, pt in enumerate(dept_m, 1):
        p_pt = doc.add_paragraph(f"{i}.  {pt}")
        p_pt.paragraph_format.left_indent = Inches(0.25)

    # --- DOCX PAGE 3: PROGRAM OUTCOMES (PO1-PO10) ---
    doc.add_page_break()
    add_page_header()
    add_title("Program Outcomes PO's", size=14, bold=True, space_before=6, space_after=8)
    
    po_data = [
        ("PO1\nEK", "Engineering knowledge", "Apply the knowledge of mathematics, science, engineering fundamentals, and an engineering specialization to the solution of complex engineering problems."),
        ("PO2\nPA", "Problem analysis", "Identify, formulate, research literature, and analyze complex engineering problems reaching substantiated conclusions using first principles of mathematics, natural sciences, and engineering sciences."),
        ("PO3\nDS", "Development/\nDesign of solutions", "Design solutions for complex engineering problems and design system components or processes that meet the specified needs with appropriate consideration for the public health and safety, and the cultural, societal, and environmental considerations."),
        ("PO4\nCP", "Conduct investigations of complex problems:", "Use research-based knowledge and research methods including design of experiments, analysis and interpretation of data, and synthesis of the information to provide valid conclusions."),
        ("PO5\nMT", "Modern tool usage:", "Create, select, and apply appropriate techniques, resources, and modern engineering and IT tools including prediction and modeling to complex engineering activities with an understanding of the limitations."),
        ("PO6\nEG", "The engineer and society:", "Apply reasoning informed by the contextual knowledge to assess societal, health, safety, legal and cultural issues and the consequent responsibilities relevant to the professional engineering practice."),
        ("PO7\nES", "Environment and sustainability:", "Understand the impact of the professional engineering solutions in societal and environmental contexts, and demonstrate the knowledge of, and need for sustainable development."),
        ("PO8\nET", "Ethics:", "Apply ethical principles and commit to professional ethics and responsibilities and norms of the engineering practice."),
        ("PO9\nIT", "Individual and team work:", "Function effectively as an individual, and as a member or leader in diverse teams, and in multidisciplinary settings."),
        ("PO10\nCO", "Communication:", "Communicate effectively on complex engineering activities with the engineering community and with society at large, such as, being able to comprehend and write effective reports and design documentation, make effective presentations, and give and receive clear instructions.")
    ]
    
    t_po = doc.add_table(rows=len(po_data), cols=3)
    t_po.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_po)
    
    for i, (code, title, desc) in enumerate(po_data):
        c0, c1, c2 = t_po.cell(i, 0), t_po.cell(i, 1), t_po.cell(i, 2)
        c0.width, c1.width, c2.width = Inches(0.8), Inches(1.8), Inches(4.2)
        set_cell_margins(c0, 60, 60, 80, 80)
        set_cell_margins(c1, 60, 60, 80, 80)
        set_cell_margins(c2, 60, 60, 80, 80)
        
        p0 = c0.paragraphs[0]; p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p0.add_run(code); r.bold = True; r.font.size = Pt(9.5)
        
        p1 = c1.paragraphs[0]
        r = p1.add_run(title); r.bold = True; r.font.size = Pt(9.5)
        
        p2 = c2.paragraphs[0]
        r = p2.add_run(desc); r.font.size = Pt(9)

    # --- DOCX PAGE 4: PO11-PO12, PSO & PEO ---
    doc.add_page_break()
    add_page_header()
    
    po_cont = [
        ("PO11\nPM", "Project management and finance:", "Demonstrate knowledge and understanding of the engineering and management principles and apply these to one's own work, as a member and leader in a team, to manage projects and in multidisciplinary environments."),
        ("PO12\nLL", "Life-long learning:", "Recognize the need for, and have the preparation and ability to engage in independent and life-long learning in the broadest context of technological change.")
    ]
    t_po2 = doc.add_table(rows=2, cols=3)
    t_po2.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_po2)
    for i, (code, title, desc) in enumerate(po_cont):
        c0, c1, c2 = t_po2.cell(i, 0), t_po2.cell(i, 1), t_po2.cell(i, 2)
        c0.width, c1.width, c2.width = Inches(0.8), Inches(1.8), Inches(4.2)
        set_cell_margins(c0, 60, 60, 80, 80)
        set_cell_margins(c1, 60, 60, 80, 80)
        set_cell_margins(c2, 60, 60, 80, 80)
        
        p0 = c0.paragraphs[0]; p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p0.add_run(code); r.bold = True; r.font.size = Pt(9.5)
        p1 = c1.paragraphs[0]
        r = p1.add_run(title); r.bold = True; r.font.size = Pt(9.5)
        p2 = c2.paragraphs[0]
        r = p2.add_run(desc); r.font.size = Pt(9)
        
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    r = p.add_run("Program Specific Outcomes:\n").bold = True
    
    p_pso1 = doc.add_paragraph()
    p_pso1.paragraph_format.left_indent = Inches(0.25)
    p_pso1.add_run("•  PSO1: ").bold = True
    p_pso1.add_run("Improving professional and problem-solving abilities to use cutting-edge technologies for solving global issues.")
    
    p_pso2 = doc.add_paragraph()
    p_pso2.paragraph_format.left_indent = Inches(0.25)
    p_pso2.add_run("•  PSO2: ").bold = True
    p_pso2.add_run("Graduates will be able to provide novel solutions by incorporating ethical behavior, effective communication skills and an entrepreneurial spirit.")
    
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    r = p.add_run("Program Educational Objectives:\n").bold = True
    
    peos = [
        ("PEO1 : ", "To prepare competent graduates having strong fundamentals, domain knowledge with modern technology to provide effective solutions for engineering problems."),
        ("PEO2: ", "Graduates will have successful technical and professional careers in industry, academia, government and entrepreneurship."),
        ("PEO3: ", "To prepare committed and motivated graduates with an investigative approach for lifelong learning.")
    ]
    for tag, text in peos:
        p_peo = doc.add_paragraph()
        p_peo.paragraph_format.left_indent = Inches(0.25)
        p_peo.add_run(f"•  {tag}").bold = True
        p_peo.add_run(text)

    # --- DOCX PAGE 5: COURSE OUTCOMES & CO-PO-PSO MAPPING ---
    doc.add_page_break()
    add_page_header()
    add_title("COURSE OUTCOMES", size=13.5, bold=True, space_before=4, space_after=6)
    
    cos = [
        ("CO1.", " To develop the understanding of the problem domain through extensive review of literature."),
        ("CO2.", " To Identify and analyze the problem in detail to define its scope with problem specific data."),
        ("CO3.", " To know various techniques to be implemented for the selected problem and related technical skills through feasibility analysis."),
        ("CO4.", " To design solutions for real-time problems that will positively impact society and the environment."),
        ("CO5.", " To develop clarity of presentation based on communication, teamwork and leadership skills."),
        ("CO6.", " To inculcate professional and ethical behavior.")
    ]
    for tag, text in cos:
        p_co = doc.add_paragraph()
        p_co.paragraph_format.space_after = Pt(3)
        p_co.add_run(tag).bold = True
        p_co.add_run(text)
        
    add_title("CO-PO-PSO MAPPING", size=13.5, bold=True, space_before=12, space_after=6)
    
    map_headers = ["", "PO1", "PO2", "PO3", "PO4", "PO5", "PO6", "PO7", "PO8", "PO9", "PO10", "PO11", "PO12", "PSO1", "PSO2"]
    map_rows = [
        ["CO1", "✓", "", "", "✓", "", "", "", "", "", "", "", "", "✓", "✓"],
        ["CO2", "✓", "✓", "", "✓", "✓", "", "", "", "", "", "", "", "✓", "✓"],
        ["CO3", "", "✓", "✓", "", "✓", "", "", "", "", "", "", "", "✓", "✓"],
        ["CO4", "✓", "✓", "", "✓", "", "✓", "✓", "", "", "", "", "", "✓", "✓"],
        ["CO5", "", "", "", "", "", "", "", "", "", "", "✓", "", "✓", "✓"],
        ["CO6", "", "", "", "", "", "", "", "✓", "✓", "✓", "", "", "", "✓"]
    ]
    
    t_map = doc.add_table(rows=7, cols=15)
    t_map.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_map)
    
    for c_idx, h in enumerate(map_headers):
        cell = t_map.cell(0, c_idx)
        cell.width = Inches(0.45)
        set_cell_margins(cell, 40, 40, 30, 30)
        set_cell_background(cell, "F2F2F2")
        p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h); r.bold = True; r.font.size = Pt(8.5)
        
    for r_idx, row_data in enumerate(map_rows, 1):
        for c_idx, val in enumerate(row_data):
            cell = t_map.cell(r_idx, c_idx)
            cell.width = Inches(0.45)
            set_cell_margins(cell, 40, 40, 30, 30)
            p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(9.5)
            if c_idx == 0:
                r.bold = True
            if val == "✓":
                r.bold = True
                r.font.color.rgb = RGBColor(0x00, 0x66, 0x00)

    # --- DOCX PAGE 6: STUDENT INFORMATION ---
    doc.add_page_break()
    add_page_header()
    add_title("STUDENT INFORMATION", size=14, bold=True, space_before=4, space_after=8)
    
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run("PROJECT TITLE: "); r.bold = True
    p.add_run("Football Vision AI – Computer Vision and IoT Powered Autonomous Video Assistant Referee (VAR)\n")
    r = p.add_run("Category of Project: "); r.bold = True
    p.add_run("Inhouse / Societal / Research based\n")
    r = p.add_run("Application Domain: "); r.bold = True
    p.add_run("Sports Analytics / Autonomous Video Refereeing / Smart Stadiums\n")
    r = p.add_run("Technical Domain: "); r.bold = True
    p.add_run("Computer Vision (OpenCV, MOG2, YOLOv8, Homography Transform), IoT & Wearable Biometrics (FastAPI WebSocket, ESP32), Full-Stack (FastAPI, React 19, Tailwind CSS)")
    
    stu_headers = ["", "STUDENT 1", "STUDENT 2", "STUDENT 3", "STUDENT 4"]
    stu_fields = [
        ("STUDENT ID", ["mannsutariya-cmpn@\natharvacoe.ac.in", "tejaswalanj-cmpn@\natharvacoe.ac.in", "karansonar-cmpn@\natharvacoe.ac.in", "yashwinrajpurohit-cmpn@\natharvacoe.ac.in"]),
        ("NAME", ["Mann Sutariya", "Tejas Walanj", "Karan Sonar", "Yashwin Rajpurohit"]),
        ("ROLL / CLASS", ["47 (BE-CMPN-2)", "55 (BE-CMPN-2)", "44 (BE-CMPN-2)", "09 (BE-CMPN-2)"]),
        ("CONTACT NO.", ["+91 98201 44521", "+91 91374 88293", "+91 98923 66104", "+91 93245 77189"]),
        ("ADDRESS", [
            "B-302, Gokul Dham Heights, Mahavir Nagar, Kandivali West, Mumbai, 400067",
            "12/A, Shanti Niketan CHS, SV Road, Malad West, Mumbai, 400064",
            "C-104, Sai Heritage, Link Road, Borivali West, Mumbai, 400092",
            "Flat 401, Shree Ganesh Tower, Evershine Nagar, Malad West, Mumbai, 400064"
        ])
    ]
    
    t_stu = doc.add_table(rows=6, cols=5)
    t_stu.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_stu)
    
    for c_idx, h in enumerate(stu_headers):
        cell = t_stu.cell(0, c_idx)
        cell.width = Inches(1.3) if c_idx > 0 else Inches(1.2)
        set_cell_margins(cell, 40, 40, 40, 40)
        set_cell_background(cell, "F2F2F2")
        p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h); r.bold = True; r.font.size = Pt(9)
        
    for r_idx, (fname, vals) in enumerate(stu_fields, 1):
        cell_lbl = t_stu.cell(r_idx, 0)
        cell_lbl.width = Inches(1.2)
        set_cell_margins(cell_lbl, 40, 40, 40, 40)
        p = cell_lbl.paragraphs[0]
        r = p.add_run(fname); r.bold = True; r.font.size = Pt(8.5)
        
        for c_idx, val in enumerate(vals, 1):
            cell = t_stu.cell(r_idx, c_idx)
            cell.width = Inches(1.3)
            set_cell_margins(cell, 40, 40, 40, 40)
            p = cell.paragraphs[0]
            r = p.add_run(val); r.font.size = Pt(8)
            if fname == "NAME":
                r.bold = True

    # --- DOCX PAGE 7: INSTRUCTIONS & DECLARATION ---
    doc.add_page_break()
    add_page_header()
    add_title("INSTRUCTIONS TO STUDENTS:", size=13.5, bold=True, space_before=10, space_after=10)
    
    p = doc.add_paragraph("1.  The logbook must be submitted to the Guide or Co-Guide for verification and evaluation of project activities at least once in a week.")
    p.paragraph_format.space_after = Pt(6)
    p = doc.add_paragraph("2.  Log book duly signed by a guide must be submitted with a project report for evaluation at the end of semester to the department.")
    p.paragraph_format.space_after = Pt(25)
    
    add_title("DECLARATION", size=13.5, bold=True, space_before=15, space_after=10)
    
    p_dec = doc.add_paragraph("I declare that this project represents my ideas in my own words without plagiarism and wherever others' ideas or words have been included, I have adequately cited and referenced the original sources. I also declare that I have adhered to all principles of academic honesty and integrity and have not misrepresented or fabricated or falsified any idea/data/fact/source in my project work. I promise to maintain minimum 75% attendance, as per the University of Mumbai norms. I understand that any violation of the above will be cause for disciplinary action by the Institute.")
    p_dec.paragraph_format.line_spacing = 1.35
    p_dec.paragraph_format.space_after = Pt(40)
    
    p_yf = doc.add_paragraph()
    p_yf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_yf.add_run("Yours Faithfully\n\n").bold = True
    
    stu_decl = [
        "1.  Mann Sutariya",
        "2.  Tejas Walanj",
        "3.  Karan Sonar",
        "4.  Yashwin Rajpurohit"
    ]
    for s in stu_decl:
        p = doc.add_paragraph(s)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_after = Pt(4)
        
    p_dt = doc.add_paragraph("(Date & Signature of Students)")
    p_dt.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_dt.paragraph_format.space_before = Pt(20)

    # --- DOCX PAGE 8: LETTER OF ACCEPTANCE ---
    doc.add_page_break()
    add_page_header()
    add_title("Letter of Acceptance", size=14.5, bold=True, space_before=20, space_after=20)
    
    p_acc = doc.add_paragraph("I undersigned, Prof. Varsha Patil working in Department of Computer Engineering, willing to guide the project titled Football Vision AI: Computer Vision and IoT Powered Autonomous Video Assistant Referee (VAR) for the BE-Project-I Semester VII for the Academic Year 2025-26. The names of the students are:")
    p_acc.paragraph_format.line_spacing = 1.4
    p_acc.paragraph_format.space_after = Pt(15)
    
    for s in ["1.  Mann Sutariya", "2.  Tejas Walanj", "3.  Karan Sonar", "4.  Yashwin Rajpurohit"]:
        p = doc.add_paragraph(s)
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_after = Pt(4)
        p.runs[0].bold = True
        
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(100)
    
    t_acc_sign = doc.add_table(rows=1, cols=3)
    t_acc_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(t_acc_sign)
    
    c0, c1, c2 = t_acc_sign.cell(0, 0), t_acc_sign.cell(0, 1), t_acc_sign.cell(0, 2)
    c0.width, c1.width, c2.width = Inches(2.2), Inches(2.5), Inches(2.1)
    
    p = c0.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.add_run("Prof. Varsha Patil\n").bold = True
    p.add_run("(Project Guide)")
    
    p = c1.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Prof. Varsha Patil / Dr. Suvarna P.\n").bold = True
    p.add_run("(BE-Project I Coordinator)")
    
    p = c2.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run("Prof. Mahendra Patil\n").bold = True
    p.add_run("(HOD-CMPN)")

    # --- DOCX PAGE 9: SCHEDULE FOR BE PROJECT-I ---
    doc.add_page_break()
    add_page_header()
    
    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_meta.paragraph_format.space_after = Pt(2)
    p_meta.add_run("ACE/CMPN/FR/16/2025-26\nDATE: 17/07/2025").bold = True
    
    add_title("SCHEDULE FOR BE PROJECT-I", size=13.5, bold=True, space_before=4, space_after=6)
    
    sched_data = [
        ("17/07/25", "1", "Basic Concept Presented & Literature Brainstorming", "Completed", "VP"),
        ("24/07/25", "2", "Title Approved: Autonomous VAR Using Computer Vision", "Done", "VP"),
        ("31/07/25", "3", "Review of Research Papers & Video Tracking Methodologies", "Working", "VP"),
        ("07/08/25", "4", "Finalizing Project Scope, Objectives & Feasibility Study", "Done", "VP"),
        ("14/08/25", "5", "Literature Survey Formulation & Evaluation of FIFA Standards", "Working", "VP"),
        ("21/08/25", "6", "Study of Existing Systems, Research Gap & Hardware Plan", "Done", "VP"),
        ("28/08/25", "7", "Proposed System Architecture & Block Diagram Design", "Working", "VP"),
        ("04/09/25", "8", "Algorithm & Process Design (Detection & Homography)", "OK", "VP"),
        ("11/09/25", "9", "Core Vision Pipeline Implementation & Ball/Player Tracking", "OK", "VP"),
        ("18/09/25", "10", "FastAPI Backend Streaming & Event Triggering Development", "Done", "VP"),
        ("25/09/25", "11", "React Dashboard Integration & Video Control Room UI", "Working", "VP"),
        ("02/10/25", "12", "Interim Project Report & Documentation Formulation", "Done", "VP"),
        ("09/10/25", "13", "End-to-End System Testing, Latency Optimization & Tuning", "Done", "VP"),
        ("16/10/25", "14", "Final Review Presentation, Working Demo & Logbook Submission", "Done", "VP")
    ]
    
    t_sch = doc.add_table(rows=15, cols=5)
    t_sch.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_sch)
    
    headers = ["Date", "Week", "Contents", "Remark", "Guide Sign"]
    for c_idx, h in enumerate(headers):
        cell = t_sch.cell(0, c_idx)
        cell.width = Inches(0.9) if c_idx == 0 else (Inches(0.6) if c_idx == 1 else (Inches(3.4) if c_idx == 2 else Inches(1.0)))
        set_cell_margins(cell, 35, 35, 40, 40)
        set_cell_background(cell, "F2F2F2")
        p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h); r.bold = True; r.font.size = Pt(8.5)
        
    for r_idx, row in enumerate(sched_data, 1):
        for c_idx, val in enumerate(row):
            cell = t_sch.cell(r_idx, c_idx)
            cell.width = Inches(0.9) if c_idx == 0 else (Inches(0.6) if c_idx == 1 else (Inches(3.4) if c_idx == 2 else Inches(1.0)))
            set_cell_margins(cell, 30, 30, 40, 40)
            p = cell.paragraphs[0]
            if c_idx in [0, 1, 3, 4]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(8)
            if c_idx in [1, 3]:
                r.bold = True

    # --- DOCX PAGE 10: PROGRESS/ATTENDANCE REPORT (PART 1) ---
    doc.add_page_break()
    add_page_header()
    
    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_meta.paragraph_format.space_after = Pt(2)
    p_meta.add_run("ACE/CMPN/FR/16/2025-26\nDATE: 17/07/2025").bold = True
    
    add_title("PROGRESS/ATTENDANCE REPORT", size=13.5, bold=True, space_before=4, space_after=6)
    
    p_info = doc.add_paragraph()
    p_info.paragraph_format.line_spacing = 1.25
    p_info.paragraph_format.space_after = Pt(6)
    p_info.add_run("Group No.: ").bold = True; p_info.add_run("03\n")
    p_info.add_run("Title of the Project: ").bold = True; p_info.add_run("Football Vision AI: Computer Vision and IoT Powered Autonomous Video Assistant Referee (VAR)\n")
    p_info.add_run("Name of Student 1: ").bold = True; p_info.add_run("Mann Sutariya   |   ")
    p_info.add_run("Name of Student 2: ").bold = True; p_info.add_run("Tejas Walanj\n")
    p_info.add_run("Name of Student 3: ").bold = True; p_info.add_run("Karan Sonar      |   ")
    p_info.add_run("Name of Student 4: ").bold = True; p_info.add_run("Yashwin Rajpurohit\n")
    p_info.add_run("Name of the Supervisor/Guide: ").bold = True; p_info.add_run("Prof. Varsha Patil")
    
    prog_rows1 = [
        ("1", "17/07/25", "P", "P", "P", "P", "Working on project title finalization, domain selection, and scope identification for AI-powered VAR.", "CO1", "PO1\nPO6", "PSO1"),
        ("2", "24/07/25", "P", "P", "P", "P", "Conducted comprehensive literature survey on computer vision in sports officiating, Hawk-Eye protocols, and deep neural models.", "CO1\nCO2", "PO2\nPO4", "PSO1\nPSO2"),
        ("3", "31/07/25", "P", "P", "P", "P", "Study and acquisition of broadcast football video datasets (SoccerNet clips) and camera calibration feeds.", "CO2\nCO3", "PO3\nPO5", "PSO1")
    ]
    
    t_prog1 = doc.add_table(rows=5, cols=10)
    t_prog1.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_prog1)
    
    # Row 0 Header
    headers_top = [("Sr.\nNo", 1, 2), ("Date", 1, 2), ("Attendance", 4, 1), ("Progress / Suggestion", 1, 2), ("Mapping", 3, 1)]
    # Row 1 subheader for att & mapping
    sub_headers = ["1", "2", "3", "4", "CO", "PO", "PSO"]
    
    # We populate the table directly with custom merged-style cells
    # Let's fill cells simply:
    h_labels = ["Sr No", "Date", "Att 1", "Att 2", "Att 3", "Att 4", "Progress / Suggestion", "CO", "PO", "PSO"]
    for c_idx, h in enumerate(h_labels):
        c = t_prog1.cell(0, c_idx)
        set_cell_margins(c, 30, 30, 30, 30)
        set_cell_background(c, "F2F2F2")
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h); r.bold = True; r.font.size = Pt(8)
        
    for r_idx, row in enumerate(prog_rows1, 1):
        for c_idx, val in enumerate(row):
            c = t_prog1.cell(r_idx, c_idx)
            set_cell_margins(c, 30, 30, 30, 30)
            p = c.paragraphs[0]
            if c_idx != 6:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val); r.font.size = Pt(8)
            if c_idx in [0, 2, 3, 4, 5, 7]:
                r.bold = True

    # --- DOCX PAGE 11: PROGRESS/ATTENDANCE REPORT (PART 2) ---
    doc.add_page_break()
    add_page_header()
    
    prog_rows2 = [
        ("4", "07/08/25", "P", "P", "P", "P", "Detailed requirement analysis (hardware specs, video capture parameters, and IoT alert triggers).", "CO2\nCO4", "PO1\nPO3\nPO5", "PSO1\nPSO2"),
        ("5", "14/08/25", "P", "P", "P", "P", "Deployment of OpenCV background subtractor (MOG2) and object detection pipeline for players and ball.", "CO3\nCO4", "PO2\nPO3\nPO5", "PSO1"),
        ("6", "21/08/25", "P", "P", "P", "P", "Multi-player tracking, centroid association, and team jersey color thresholding in HSV color space.", "CO3\nCO4", "PO3\nPO5", "PSO1"),
        ("7", "28/08/25", "P", "P", "P", "P", "Pitch keypoint detection and perspective transformation (homography matrix to 2D top-down pitch view).", "CO3\nCO4", "PO2\nPO3\nPO5", "PSO1\nPSO2"),
        ("8", "04/09/25", "P", "P", "P", "P", "Automated offside line projection and goal-line boundary crossing evaluation algorithms.", "CO4", "PO3\nPO5", "PSO1"),
        ("9", "11/09/25", "P", "P", "P", "P", "FastAPI streaming backend development with MJPEG video feed endpoints and event logging database.", "CO4", "PO3\nPO5", "PSO1"),
        ("10", "18/09/25", "P", "P", "P", "P", "Design and implementation of React.js VAR Studio dashboard with multi-camera switching and decision alerts.", "CO4\nCO5", "PO5\nPO10", "PSO1\nPSO2"),
        ("11", "25/09/25", "P", "P", "P", "P", "System benchmarking, latency evaluation under broadcast feeds, and IoT alert mechanism testing.", "CO4\nCO6", "PO8\nPO11", "PSO1\nPSO2"),
        ("12", "09/10/25", "P", "P", "P", "P", "Created complete logbook, completed BE-Project-I report documentation, and conducted presentation dry-run.", "CO5\nCO6", "PO10\nPO11\nPO12", "PSO1\nPSO2")
    ]
    
    t_prog2 = doc.add_table(rows=10, cols=10)
    t_prog2.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_prog2)
    
    for c_idx, h in enumerate(h_labels):
        c = t_prog2.cell(0, c_idx)
        set_cell_margins(c, 25, 25, 25, 25)
        set_cell_background(c, "F2F2F2")
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h); r.bold = True; r.font.size = Pt(8)
        
    for r_idx, row in enumerate(prog_rows2, 1):
        for c_idx, val in enumerate(row):
            c = t_prog2.cell(r_idx, c_idx)
            set_cell_margins(c, 25, 25, 25, 25)
            p = c.paragraphs[0]
            if c_idx != 6:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val); r.font.size = Pt(7.5)
            if c_idx in [0, 2, 3, 4, 5, 7]:
                r.bold = True
                
    p_sup = doc.add_paragraph()
    p_sup.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sup.paragraph_format.space_before = Pt(30)
    r = p_sup.add_run("Prof. Varsha Patil\n").bold = True
    r2 = p_sup.add_run("Name, Date & Sign of the Supervisor/Guide")
    r2.font.size = Pt(9.5)

    # --- DOCX PAGE 12: REVIEW-I FORM ---
    doc.add_page_break()
    add_page_header()
    add_title("REVIEW-I FORM", size=14, bold=True, space_before=4, space_after=6)
    
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.25
    p.paragraph_format.space_after = Pt(6)
    p.add_run("Group No.: ").bold = True; p.add_run("03\n")
    p.add_run("Title of BE-Project I: ").bold = True; p.add_run("Football Vision AI: Computer Vision and IoT Powered Autonomous Video Assistant Referee (VAR)\n")
    p.add_run("Date of Review-I: ").bold = True; p.add_run("22/08/2025\n")
    p.add_run("No. of students in project team: ").bold = True; p.add_run("4")
    
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    p.add_run("Student BE-Project I Performance Analysis ").bold = True
    p.add_run("(Put Tick as per your Observation)")
    
    rev1_obs = [
        ("1", "Quality of problem and Clarity", "✓", "", ""),
        ("2", "Clarity of Problem definition based on need.", "✓", "", ""),
        ("3", "Innovativeness in solutions", "✓", "", ""),
        ("4", "Feasibility of proposed problem solutions and selection of best solution", "", "✓", ""),
        ("5", "Cost effectiveness and Societal impact", "✓", "", ""),
        ("6", "Overall Presentation & Performance", "✓", "", "")
    ]
    
    t_rev1 = doc.add_table(rows=8, cols=5)
    t_rev1.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_rev1)
    
    # Headers
    rev_headers = ["Sr. No.", "Observation", "Excellent (3)", "Very Good (2)", "Good (1)"]
    for c_idx, h in enumerate(rev_headers):
        c = t_rev1.cell(0, c_idx)
        c.width = Inches(0.8) if c_idx == 0 else (Inches(3.8) if c_idx == 1 else Inches(0.8))
        set_cell_margins(c, 30, 30, 30, 30)
        set_cell_background(c, "F2F2F2")
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h); r.bold = True; r.font.size = Pt(8.5)
        
    for r_idx, row in enumerate(rev1_obs, 1):
        for c_idx, val in enumerate(row):
            c = t_rev1.cell(r_idx, c_idx)
            c.width = Inches(0.8) if c_idx == 0 else (Inches(3.8) if c_idx == 1 else Inches(0.8))
            set_cell_margins(c, 30, 30, 30, 30)
            p = c.paragraphs[0]
            if c_idx != 1:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val); r.font.size = Pt(8.5)
            if c_idx == 0:
                r.bold = True
            if val == "✓":
                r.bold = True
                r.font.color.rgb = RGBColor(0x00, 0x66, 0x00)
                
    # Comments
    p_cmt = doc.add_paragraph()
    p_cmt.paragraph_format.space_before = Pt(8)
    p_cmt.add_run("Comments: ").bold = True
    p_cmt.add_run("Working on papers and authentication paper. More papers to be surveyed on sports video analytics and pitch homography algorithms. Proposed architecture is sound and approved.")
    
    p_sign = doc.add_paragraph()
    p_sign.paragraph_format.space_before = Pt(14)
    p_sign.add_run("Project Guide & Panel Members Signature:\n").bold = True
    p_sign.add_run("1) Prof. Varsha Patil                 2) Prof. Pranali B.                 3) Prof. Anuja G.")
    
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(35)
    
    t_signs = doc.add_table(rows=1, cols=2)
    t_signs.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(t_signs)
    c0, c1 = t_signs.cell(0, 0), t_signs.cell(0, 1)
    p = c0.paragraphs[0]; p.add_run("Prof. Varsha Patil / Dr. Suvarna P.\n").bold = True; p.add_run("Project Coordinator")
    p = c1.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p.add_run("Prof. Mahendra Patil\n").bold = True; p.add_run("HOD-CMPN")

    # --- DOCX PAGE 13: REVIEW-II FORM ---
    doc.add_page_break()
    add_page_header()
    add_title("REVIEW-II FORM", size=14, bold=True, space_before=4, space_after=6)
    
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.25
    p.paragraph_format.space_after = Pt(6)
    p.add_run("Group No.: ").bold = True; p.add_run("03\n")
    p.add_run("Title of BE-Project I: ").bold = True; p.add_run("Football Vision AI: Computer Vision and IoT Powered Autonomous Video Assistant Referee (VAR)\n")
    p.add_run("Date of Review-II: ").bold = True; p.add_run("03/10/2025\n")
    p.add_run("No. of students in project team: ").bold = True; p.add_run("4")
    
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    p.add_run("Student BE-Project I Performance Analysis ").bold = True
    p.add_run("(Put Tick as per your Observation)")
    
    rev2_obs = [
        ("1", "Usage of effective skill sets", "✓", "", ""),
        ("2", "Effective use of standard engineering norms", "✓", "", ""),
        ("3", "Contribution of an individual's as member or leader", "", "✓", ""),
        ("4", "Clarity in written and oral communication", "✓", "", ""),
        ("5", "Functioning of working model as per stated requirements", "✓", "", ""),
        ("6", "Overall Presentation & Performance", "✓", "", "")
    ]
    
    t_rev2 = doc.add_table(rows=7, cols=5)
    t_rev2.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_rev2)
    
    for c_idx, h in enumerate(rev_headers):
        c = t_rev2.cell(0, c_idx)
        c.width = Inches(0.8) if c_idx == 0 else (Inches(3.8) if c_idx == 1 else Inches(0.8))
        set_cell_margins(c, 30, 30, 30, 30)
        set_cell_background(c, "F2F2F2")
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h); r.bold = True; r.font.size = Pt(8.5)
        
    for r_idx, row in enumerate(rev2_obs, 1):
        for c_idx, val in enumerate(row):
            c = t_rev2.cell(r_idx, c_idx)
            c.width = Inches(0.8) if c_idx == 0 else (Inches(3.8) if c_idx == 1 else Inches(0.8))
            set_cell_margins(c, 30, 30, 30, 30)
            p = c.paragraphs[0]
            if c_idx != 1:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val); r.font.size = Pt(8.5)
            if c_idx == 0:
                r.bold = True
            if val == "✓":
                r.bold = True
                r.font.color.rgb = RGBColor(0x00, 0x66, 0x00)
                
    p_cmt = doc.add_paragraph()
    p_cmt.paragraph_format.space_before = Pt(8)
    p_cmt.add_run("Comments: ").bold = True
    p_cmt.add_run("Working on documentation and interfacing of back-end (FastAPI streaming engine) with front-end (React VAR Dashboard). Working model demonstrated effectively. Ready for final examination.")
    
    p_sign = doc.add_paragraph()
    p_sign.paragraph_format.space_before = Pt(14)
    p_sign.add_run("Project Guide & Panel Members Signature:\n").bold = True
    p_sign.add_run("1) Prof. Varsha Patil                 2) Prof. Pranali B.                 3) Prof. Anuja G.")
    
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(35)
    
    t_signs2 = doc.add_table(rows=1, cols=2)
    t_signs2.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(t_signs2)
    c0, c1 = t_signs2.cell(0, 0), t_signs2.cell(0, 1)
    p = c0.paragraphs[0]; p.add_run("Prof. Varsha Patil / Dr. Suvarna P.\n").bold = True; p.add_run("Project Coordinator")
    p = c1.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p.add_run("Prof. Mahendra Patil\n").bold = True; p.add_run("HOD-CMPN")

    # --- DOCX PAGE 14: EXAMINER'S FEEDBACK FORM ---
    doc.add_page_break()
    add_page_header()
    add_title("EXAMINER'S FEEDBACK FORM", size=14, bold=True, space_before=4, space_after=6)
    
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.25
    p.paragraph_format.space_after = Pt(4)
    p.add_run("Name of External examiner: ").bold = True; p.add_run("Dr. Nazneen Ansari\n")
    p.add_run("College of External examiner: ").bold = True; p.add_run("SFIT, Borivali\n")
    p.add_run("Name of Internal examiner: ").bold = True; p.add_run("Prof. Varsha Patil\n")
    p.add_run("Date of Examination: ").bold = True; p.add_run("27/10/2025\n")
    p.add_run("No. of students in project team: ").bold = True; p.add_run("4")
    
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    p.add_run("Student Performance Analysis ").bold = True
    p.add_run("(Put Tick as per your Observation)")
    
    exam_obs = [
        ("1", "Quality of problem and Clarity", "✓", "", ""),
        ("2", "Clarity of Problem definition based on need.", "✓", "", ""),
        ("3", "Innovativeness in solutions", "", "✓", ""),
        ("4", "Feasibility of proposed problem solutions and selection of best solution", "✓", "", ""),
        ("5", "Cost effectiveness and Societal impact", "✓", "", ""),
        ("6", "Usage of effective skill sets", "", "✓", ""),
        ("7", "Effective use of standard engineering norms", "", "✓", ""),
        ("8", "Contribution of an individual's as member or leader", "✓", "", ""),
        ("9", "Clarity in written and oral communication", "✓", "", ""),
        ("10", "Full functioning of working model as per stated requirements", "✓", "", ""),
        ("11", "Overall Presentation & Performance", "✓", "", "")
    ]
    
    t_ex = doc.add_table(rows=12, cols=5)
    t_ex.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_ex)
    
    for c_idx, h in enumerate(rev_headers):
        c = t_ex.cell(0, c_idx)
        c.width = Inches(0.8) if c_idx == 0 else (Inches(3.8) if c_idx == 1 else Inches(0.8))
        set_cell_margins(c, 25, 25, 25, 25)
        set_cell_background(c, "F2F2F2")
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h); r.bold = True; r.font.size = Pt(8.5)
        
    for r_idx, row in enumerate(exam_obs, 1):
        for c_idx, val in enumerate(row):
            c = t_ex.cell(r_idx, c_idx)
            c.width = Inches(0.8) if c_idx == 0 else (Inches(3.8) if c_idx == 1 else Inches(0.8))
            set_cell_margins(c, 25, 25, 25, 25)
            p = c.paragraphs[0]
            if c_idx != 1:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val); r.font.size = Pt(8)
            if c_idx == 0:
                r.bold = True
            if val == "✓":
                r.bold = True
                r.font.color.rgb = RGBColor(0x00, 0x66, 0x00)
                
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(30)
    
    t_ex_sign = doc.add_table(rows=1, cols=3)
    t_ex_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(t_ex_sign)
    c0, c1, c2 = t_ex_sign.cell(0, 0), t_ex_sign.cell(0, 1), t_ex_sign.cell(0, 2)
    
    p = c0.paragraphs[0]; p.add_run("Dr. Nazneen Ansari\n").bold = True; p.add_run("External Examiner\n27/10/2025")
    p = c1.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.add_run("Prof. Varsha Patil\n").bold = True; p.add_run("Internal Examiner\n27/10/2025")
    p = c2.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p.add_run("Prof. Mahendra Patil\n").bold = True; p.add_run("HOD-CMPN")

    docx_path = os.path.abspath("BE_Project_I_Logbook_Football_Vision_AI.docx")
    doc.save(docx_path)
    print(f"Generated DOCX successfully: {docx_path}")

build_docx()
