#!/usr/bin/env python3
"""
generate_complete_lab_record.py
Compiles the complete Amity University CSIT248 (Advanced Web Technology)
Laboratory Record for all 25 experiments.
Generates:
1. CSIT248_Lab_Record.html (Self-contained, printable HTML)
2. CSIT248_Lab_Record.pdf  (High-res PDF generated via headless Chrome)
3. CSIT248_Lab_Record.docx (Fully formatted Word Document)
"""

import os
import sys
import base64
import html
import subprocess
from PIL import Image

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "lab_record_assets")

STUDENT_NAME = "Atul Anand"
ENROLLMENT_NO = "A010150825001"
PROGRAMME = "BCA (Hons./Hons. with Research)"
SEMESTER = "3"
COURSE_CODE = "CSIT248"
SUBJECT_NAME = "Advanced Web Technology"
UNIVERSITY = "AMITY UNIVERSITY"
CAMPUS = "UTTAR PRADESH"

EXPERIMENTS = [
    {
        "id": "01",
        "num_display": "1",
        "title": "Demonstration of Basic HTML Tags, Fonts, Images, and Lists.",
        "aim": "To create a web page using basic HTML tags such as Body, Horizontal Rule, Font, Image, Ordered List, and Unordered List, and to demonstrate different font styles and list types.",
        "date_performed": "21-07-2026",
        "date_submitted": "04-08-2026",
        "src_path": "/Users/atulanand/Desktop/My-files/Adv-Web-Tech-Lab-Class/First-Lab/Today_assing-21-07-2026/index.html",
        "img_path": os.path.join(ASSETS_DIR, "exp01_output.png")
    },
    {
        "id": "02",
        "num_display": "2",
        "title": "Creating a Simple Burger Café Web Page Using HTML.",
        "aim": "To design a simple restaurant web page using basic HTML elements such as headings, images, lists, paragraphs, and hyperlinks to display menu details, contact information, and social media links.",
        "date_performed": "28-07-2026",
        "date_submitted": "04-08-2026",
        "src_path": "/Users/atulanand/Desktop/My-files/Adv-Web-Tech-Lab-Class/Second-Lab/Today-Lab-28-07/index.html",
        "img_path": os.path.join(ASSETS_DIR, "exp02_output.png")
    },
    {
        "id": "03",
        "num_display": "3",
        "title": "Creating a Web Page with Headings, Hyperlinks, Images, and Lists.",
        "aim": "To create a simple HTML web page that displays an image, headings, formatted text, unordered lists, and hyperlinks to demonstrate navigation using HTML.",
        "date_performed": "28-07-2026",
        "date_submitted": "04-08-2026",
        "src_path": "/Users/atulanand/Desktop/My-files/Adv-Web-Tech-Lab-Class/Second-Lab/Today-Lab-28-07/gfg.html",
        "img_path": os.path.join(ASSETS_DIR, "exp03_output.png")
    },
    {
        "id": "04",
        "num_display": "4",
        "title": "Student Registration Form with Rich Form Controls.",
        "aim": "To design a comprehensive student registration web form utilizing diverse input controls (text, radio, checkboxes, date, select dropdowns, file uploads, textareas) and structured fieldsets.",
        "date_performed": "04-08-2026",
        "date_submitted": "11-08-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-1/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp04_output.png")
    },
    {
        "id": "05",
        "num_display": "5",
        "title": "Personal Introduction Web Page Showcasing Bio, Timeline, and Skills.",
        "aim": "To design a web page to introduce yourself, showcasing personal bio, academic timeline, technical competencies, definition lists, marquee banner, and contact details.",
        "date_performed": "04-08-2026",
        "date_submitted": "11-08-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-2/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp05_output.png")
    },
    {
        "id": "06",
        "num_display": "6",
        "title": "Weekly Academic Timetable Matrix with Colspan and Rowspan.",
        "aim": "To construct a weekly semester timetable matrix using HTML table elements, cell padding, cell spacing, and column/row spans to represent overlapping lecture slots and laboratory practicals.",
        "date_performed": "11-08-2026",
        "date_submitted": "18-08-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-3/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp06_output.png")
    },
    {
        "id": "07",
        "num_display": "7",
        "title": "Client-Side Image Mapping with Poly and Rect Hotspots.",
        "aim": "To implement interactive client-side image mapping using HTML <map> and <area> tags with circular, polygonal, and rectangular hotspot coordinates linked to target resources.",
        "date_performed": "11-08-2026",
        "date_submitted": "18-08-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-4/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp07_output.png")
    },
    {
        "id": "08",
        "num_display": "8",
        "title": "Nested Ordered and Unordered Categorized Lists.",
        "aim": "To design an HTML web page demonstrating complex hierarchical nested ordered lists, unordered lists with varied bullet types (disc, circle, square), and styled form button controls.",
        "date_performed": "18-08-2026",
        "date_submitted": "25-08-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-5/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp08_output.png")
    },
    {
        "id": "09",
        "num_display": "9",
        "title": "Computer Hardware Table Gallery with Borders and Cellpadding.",
        "aim": "To construct a computer hardware image gallery using HTML table structure, specifying borders, cell padding, cell spacing, captions, and aligned thumbnail descriptions.",
        "date_performed": "18-08-2026",
        "date_submitted": "25-08-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-6/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp09_output.png")
    },
    {
        "id": "10",
        "num_display": "10",
        "title": "Client-Side History Navigation using JavaScript window.history.",
        "aim": "To demonstrate client-side web browser history manipulation utilizing JavaScript window.history methods (back(), forward(), go(), and history.length).",
        "date_performed": "25-08-2026",
        "date_submitted": "01-09-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-7/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp10_output.png")
    },
    {
        "id": "11",
        "num_display": "11",
        "title": "Comprehensive Front-Level Form Validation with Regular Expressions.",
        "aim": "To implement robust client-side front-level form validation in JavaScript using regular expressions for student registration details, validating email, phone, passwords, and required fields before submission.",
        "date_performed": "25-08-2026",
        "date_submitted": "01-09-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-8/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp11_output.png")
    },
    {
        "id": "12",
        "num_display": "12",
        "title": "External and Inline CSS Design Rules Demonstration.",
        "aim": "To demonstrate Cascading Style Sheets (CSS) priority rules, class selectors, ID selectors, inline styling, and external stylesheet integration on a course evaluation and feedback form.",
        "date_performed": "01-09-2026",
        "date_submitted": "08-09-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-9/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp12_output.png")
    },
    {
        "id": "13",
        "num_display": "13",
        "title": "DHTML Dynamic Records Management Form.",
        "aim": "To build a dynamic HTML (DHTML) data management interface capable of adding, editing, validating, and deleting student records in the DOM dynamically without page refresh.",
        "date_performed": "01-09-2026",
        "date_submitted": "08-09-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-10/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp13_output.png")
    },
    {
        "id": "14",
        "num_display": "14",
        "title": "Computer Peripherals Shop Catalog with Live Filters.",
        "aim": "To design an e-commerce hardware peripherals product catalog utilizing HTML tables, responsive layout, styled badges, and client-side category filtering.",
        "date_performed": "08-09-2026",
        "date_submitted": "15-09-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-11/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp14_output.png")
    },
    {
        "id": "15",
        "num_display": "15",
        "title": "Home Town Heritage Showcase (Patna, Bihar) with Inline CSS.",
        "aim": "To design an attractive web page showcasing hometown heritage, culture, landmarks, and cuisine (Patna, Bihar) utilizing 100% inline CSS properties, custom typography, and responsive layouts.",
        "date_performed": "08-09-2026",
        "date_submitted": "15-09-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-12/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp15_output.png")
    },
    {
        "id": "16",
        "num_display": "16",
        "title": "Student XML Database & Server-Side Lookup System.",
        "aim": "To create an XML document of 10 students of SEM VI IT storing 5 subject marks, total marks, and percentage, and write a program that accepts an Enrollment No. and queries the XML file to retrieve scorecard data.",
        "date_performed": "15-09-2026",
        "date_submitted": "22-09-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-13/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp16_output.png"),
        "extra_files": [
            ("students.xml", os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-13/students.xml")),
            ("student_lookup.php", os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-13/student_lookup.php"))
        ]
    },
    {
        "id": "17",
        "num_display": "17",
        "title": "Interactive Login Portal with Real-Time Validation.",
        "aim": "To design an institutional student/faculty login portal implementing real-time client-side credential verification, show/hide password toggle, and entropy strength meter.",
        "date_performed": "15-09-2026",
        "date_submitted": "22-09-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-14/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp17_output.png")
    },
    {
        "id": "18",
        "num_display": "18",
        "title": "Web Technology Quiz Engine and Dynamic Score Calculation.",
        "aim": "To build an interactive web technology quiz system featuring multiple choice questions, timer countdown, dynamic answer evaluation, and immediate percentage scorecard display.",
        "date_performed": "22-09-2026",
        "date_submitted": "29-09-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-15/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp18_output.png")
    },
    {
        "id": "19",
        "num_display": "19",
        "title": "Proportional HTML Frames Layout Using Percentage Dimensions.",
        "aim": "To create a multi-window web layout using HTML <frameset> and <frame> tags partitioned into proportional percentage rows and columns with hyperlinked target navigation.",
        "date_performed": "22-09-2026",
        "date_submitted": "29-09-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-16/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp19_output.png")
    },
    {
        "id": "20",
        "num_display": "20",
        "title": "Responsive Menu Bar, DIV Partitions and IFRAME Sandboxing.",
        "aim": "To construct a modern responsive web portal featuring a top horizontal menu bar, multi-column <div> layout partitions, mobile hamburger toggle, and sandboxed <iframe> content loaders.",
        "date_performed": "29-09-2026",
        "date_submitted": "06-10-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-17/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp20_output.png")
    },
    {
        "id": "21",
        "num_display": "21",
        "title": "CSS Pseudo-Classes Interactive Studio & Sandbox.",
        "aim": "To demonstrate interactive user-action and structural CSS pseudo-classes (:hover, :active, :focus, :first-child, :nth-child, :disabled) through a real-time interactive styling sandbox.",
        "date_performed": "29-09-2026",
        "date_submitted": "06-10-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-18/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp21_output.png")
    },
    {
        "id": "22",
        "num_display": "22",
        "title": "Add New Record in CUSTOMER Database using PHP and MySQL.",
        "aim": "To write and execute a secure PHP script to insert new customer records into a MySQL database utilizing prepared statements, input sanitization, and SQL injection defense.",
        "date_performed": "06-10-2026",
        "date_submitted": "13-10-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-19/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp22_output.png"),
        "extra_files": [
            ("customer_schema.sql", os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-19/customer_schema.sql")),
            ("add_customer.php", os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-19/add_customer.php"))
        ]
    },
    {
        "id": "23",
        "num_display": "23",
        "title": "Implementation of include( ) and require( ) Mechanics in PHP.",
        "aim": "To design a modular PHP architecture demonstrating the functionality, differences, and error handling behaviors (E_WARNING vs E_COMPILE_ERROR) of include(), require(), include_once(), and require_once().",
        "date_performed": "06-10-2026",
        "date_submitted": "13-10-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-20/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp23_output.png"),
        "extra_files": [
            ("include_require_demo.php", os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-20/include_require_demo.php"))
        ]
    },
    {
        "id": "24",
        "num_display": "24",
        "title": "Demonstration of Array, String and Numeric Functions in PHP.",
        "aim": "To write a PHP script to demonstrate and execute native array functions (array_push, sort), string functions (strlen, strpos, str_replace), and numeric functions (round, sqrt, rand).",
        "date_performed": "06-10-2026",
        "date_submitted": "13-10-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-21/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp24_output.png"),
        "extra_files": [
            ("functions_demo.php", os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-21/functions_demo.php"))
        ]
    },
    {
        "id": "25",
        "num_display": "25",
        "title": "Implementation of MySQL DML Commands (DELETE, ORDER BY, UPDATE).",
        "aim": "To implement an interactive database manager in PHP executing atomic MySQL DML operations: DELETE to remove records, ORDER BY for multi-column sorting, and UPDATE with prepared statements.",
        "date_performed": "06-10-2026",
        "date_submitted": "13-10-2026",
        "src_path": os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-22/index.html"),
        "img_path": os.path.join(ASSETS_DIR, "exp25_output.png"),
        "extra_files": [
            ("schema.sql", os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-22/schema.sql")),
            ("mysql_crud_demo.php", os.path.join(BASE_DIR, "HTML-LAB-WORK/experiment-22/mysql_crud_demo.php"))
        ]
    }
]

def to_base64_data_uri(filepath):
    if not os.path.exists(filepath):
        return ""
    with open(filepath, "rb") as f:
        data = f.read()
    ext = os.path.splitext(filepath)[1].lower().replace(".", "")
    if ext == "jpg":
        ext = "jpeg"
    return f"data:image/{ext};base64,{base64.b64encode(data).decode('utf-8')}"

def read_file_safe(filepath):
    if not os.path.exists(filepath):
        return f"<!-- Source code file not found: {filepath} -->"
    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        return f.read()

def generate_html():
    print("Generating CSIT248_Lab_Record.html...")
    hdr_data = to_base64_data_uri(os.path.join(ASSETS_DIR, "amity_header.png"))
    shd_data = to_base64_data_uri(os.path.join(ASSETS_DIR, "amity_shield.png"))

    html_parts = []
    html_parts.append("""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CSIT248 - Advanced Web Technology Lab Record - Atul Anand</title>
<style>
  @page {
    size: A4;
    margin: 18mm 20mm 18mm 20mm;
  }
  * {
    box-sizing: border-box;
  }
  body {
    margin: 0;
    padding: 0;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    color: #111827;
    background: #ffffff;
    font-size: 11pt;
    line-height: 1.5;
  }
  
  @media screen {
    body {
      background: #f1f5f9;
      padding: 30px 0;
    }
    .page-sheet {
      background: #ffffff;
      max-width: 850px;
      margin: 0 auto 30px auto;
      padding: 45px 55px;
      border-radius: 8px;
      box-shadow: 0 4px 25px rgba(0, 0, 0, 0.08);
      box-sizing: border-box;
    }
    .cover-sheet {
      min-height: 1050px;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
    }
  }

  @media print {
    body {
      background: #ffffff;
      padding: 0;
    }
    .page-sheet {
      max-width: 100%;
      margin: 0;
      padding: 0;
      box-shadow: none;
      border-radius: 0;
    }
    .page-break {
      page-break-before: always;
      break-before: page;
    }
  }

  /* Print Button Floating Bar */
  .no-print-bar {
    position: fixed;
    top: 15px;
    right: 20px;
    z-index: 9999;
    display: flex;
    gap: 10px;
    background: rgba(15, 23, 42, 0.88);
    backdrop-filter: blur(8px);
    padding: 8px 14px;
    border-radius: 30px;
    box-shadow: 0 4px 15px rgba(0,0,0,0.25);
  }
  .no-print-bar button {
    background: #00c7b7;
    color: #042022;
    border: none;
    font-weight: 700;
    font-size: 13px;
    padding: 8px 16px;
    border-radius: 20px;
    cursor: pointer;
    display: flex;
    align-items: center;
    gap: 6px;
    transition: all 0.2s;
  }
  .no-print-bar button:hover {
    background: #14f0de;
    transform: translateY(-1px);
  }
  .no-print-bar a {
    color: #cbd5e1;
    text-decoration: none;
    font-size: 13px;
    display: flex;
    align-items: center;
    padding: 0 8px;
    font-weight: 500;
  }
  @media print {
    .no-print-bar {
      display: none !important;
    }
  }

  /* COVER PAGE */
  .cover-container {
    height: 94vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: space-between;
    text-align: center;
    padding: 5mm 0;
  }
  .cover-header-logo {
    width: 100%;
    max-width: 480px;
    margin-top: 5mm;
  }
  .cover-shield-logo {
    width: 165px;
    height: auto;
    margin: 8mm 0;
  }
  .cover-course-code {
    font-size: 20pt;
    font-weight: 800;
    letter-spacing: 2px;
    margin-bottom: 2mm;
    color: #000000;
  }
  .cover-subject {
    font-size: 18pt;
    font-weight: 700;
    color: #000000;
    letter-spacing: 0.5px;
    margin-bottom: 8mm;
  }
  .cover-submitted-by {
    font-size: 13.5pt;
    font-weight: 500;
    color: #374151;
    margin-bottom: 6mm;
  }
  .cover-details-card {
    text-align: left;
    display: inline-block;
    margin-bottom: 12mm;
    font-size: 12pt;
    line-height: 1.85;
  }
  .cover-details-card table {
    border-collapse: collapse;
  }
  .cover-details-card td {
    padding: 3px 8px;
  }
  .cover-details-card td.lbl {
    font-weight: 700;
    color: #000000;
    width: 160px;
  }
  .cover-details-card td.val {
    color: #1f2937;
  }

  /* INDEX PAGE */
  .index-container {
    padding-top: 5mm;
  }
  .index-title {
    font-size: 26pt;
    font-weight: 800;
    margin-bottom: 6mm;
    color: #000000;
  }
  .index-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 8.5pt;
  }
  .index-table th {
    border: 1px solid #000000;
    padding: 7px 6px;
    font-weight: 700;
    text-align: left;
    background: #f9fafb;
    color: #000000;
  }
  .index-table td {
    border: 1px solid #374151;
    padding: 5px 6px;
    vertical-align: middle;
    line-height: 1.3;
  }
  .index-table tr:nth-child(even) {
    background-color: #fafbfc;
  }

  /* EXPERIMENT PAGES */
  .exp-page {
    padding-top: 4mm;
  }
  .exp-num-head {
    font-size: 14pt;
    font-weight: 800;
    color: #000000;
    margin-bottom: 3.5mm;
  }
  .exp-field-row {
    font-size: 11pt;
    margin-bottom: 3.5mm;
    line-height: 1.45;
  }
  .exp-label {
    font-weight: 800;
    color: #000000;
  }
  .exp-title-text {
    font-weight: 700;
    color: #b91c1c; /* Red color matching PDF */
  }
  .exp-aim-text {
    font-weight: 600;
    color: #1d4ed8; /* Blue color matching PDF */
  }
  .exp-code-heading {
    font-size: 11pt;
    font-weight: 800;
    color: #000000;
    margin: 4mm 0 2mm 0;
  }
  .exp-code-block {
    font-family: Consolas, "Courier New", Courier, monospace;
    font-size: 8.2pt;
    line-height: 1.35;
    white-space: pre-wrap;
    word-break: break-all;
    color: #111827;
    background: #ffffff;
    padding: 0;
    margin: 0 0 4mm 0;
  }
  .exp-output-heading {
    font-size: 11pt;
    font-weight: 800;
    color: #000000;
    margin: 4mm 0 3mm 0;
  }
  .exp-output-box {
    text-align: center;
    margin-top: 2mm;
    margin-bottom: 6mm;
  }
  .exp-output-box img {
    max-width: 95%;
    max-height: 480px;
    border: 1px solid #94a3b8;
    border-radius: 4px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.1);
  }
  .extra-file-heading {
    font-size: 10pt;
    font-weight: 700;
    color: #4338ca;
    margin: 3mm 0 1.5mm 0;
  }
</style>
</head>
<body>

<div class="no-print-bar">
  <button onclick="window.print()">
    🖨️ Print / Save as PDF
  </button>
  <a href="#index">📑 Index</a>
</div>

<!-- COVER PAGE -->
<div class="page-sheet cover-sheet">
  <div class="cover-container">
    <div>
      <img src=\"""" + hdr_data + """\" alt="Amity University Uttar Pradesh" class="cover-header-logo">
    </div>
    <div>
      <img src=\"""" + shd_data + """\" alt="Amity University Shield" class="cover-shield-logo">
      <div class="cover-course-code">""" + COURSE_CODE + """</div>
      <div class="cover-subject">""" + SUBJECT_NAME + """</div>
    </div>
    <div>
      <div class="cover-submitted-by">Submitted by</div>
      <div class="cover-details-card">
        <table>
          <tr>
            <td class="lbl">Name:</td>
            <td class="val">""" + STUDENT_NAME + """</td>
          </tr>
          <tr>
            <td class="lbl">Enrollment No.:</td>
            <td class="val">""" + ENROLLMENT_NO + """</td>
          </tr>
          <tr>
            <td class="lbl">Programme:</td>
            <td class="val">""" + PROGRAMME + """</td>
          </tr>
          <tr>
            <td class="lbl">Semester:</td>
            <td class="val">""" + SEMESTER + """</td>
          </tr>
        </table>
      </div>
    </div>
  </div>
</div>

<div class="page-break"></div>

<!-- INDEX PAGE -->
<div class="page-sheet index-sheet" id="index">
  <div class="index-container">
    <div class="index-title">Index</div>
    <table class="index-table">
      <thead>
        <tr>
          <th style="width: 7%; text-align: center;">S. No.</th>
          <th style="width: 11%; text-align: center;">Experiment No.</th>
          <th style="width: 44%;">Title of Experiment</th>
          <th style="width: 14%; text-align: center;">Date Performed</th>
          <th style="width: 14%; text-align: center;">Date Submitted</th>
          <th style="width: 10%; text-align: center;">Faculty Signature</th>
        </tr>
      </thead>
      <tbody>
""")

    # Add Index Rows
    for idx, exp in enumerate(EXPERIMENTS, 1):
        html_parts.append(f"""        <tr>
          <td style="text-align: center;">{idx}</td>
          <td style="text-align: center;">{exp['id']}</td>
          <td>{html.escape(exp['title'])}</td>
          <td style="text-align: center;">{exp['date_performed']}</td>
          <td style="text-align: center;">{exp['date_submitted']}</td>
          <td></td>
        </tr>""")

    html_parts.append("""      </tbody>
    </table>
  </div>
</div>
""")

    # Add All 25 Experiments
    for exp in EXPERIMENTS:
        src_code = read_file_safe(exp["src_path"])
        img_b64 = to_base64_data_uri(exp["img_path"])

        extra_code_html = ""
        if "extra_files" in exp:
            for extra_name, extra_path in exp["extra_files"]:
                extra_content = read_file_safe(extra_path)
                extra_code_html += f"""
  <div class="extra-file-heading">Supplementary Source File: {html.escape(extra_name)}</div>
  <pre class="exp-code-block">{html.escape(extra_content)}</pre>
"""

        html_parts.append(f"""
<div class="page-break"></div>
<div class="page-sheet exp-sheet" id="exp-{exp['id']}">
  <div class="exp-page">
    <div class="exp-num-head">Experiment No.: {exp['num_display']}</div>
    
    <div class="exp-field-row">
      <span class="exp-label">Title : </span><span class="exp-title-text">{html.escape(exp['title'])}</span>
    </div>

    <div class="exp-field-row">
      <span class="exp-label">Aim : </span><span class="exp-aim-text">{html.escape(exp['aim'])}</span>
    </div>

    <div class="exp-code-heading">Source Code :</div>
    <pre class="exp-code-block">{html.escape(src_code)}</pre>
    {extra_code_html}

    <div class="exp-output-heading">Output :</div>
    <div class="exp-output-box">
      <img src="{img_b64}" alt="Output of Experiment {exp['num_display']}">
    </div>
  </div>
</div>
""")

    html_parts.append("""
</body>
</html>
""")

    out_file = os.path.join(BASE_DIR, "CSIT248_Lab_Record.html")
    with open(out_file, "w", encoding="utf-8") as f:
        f.write("\n".join(html_parts))
    print(f"Generated HTML successfully: {out_file} ({os.path.getsize(out_file)} bytes)")
    return out_file

def generate_pdf(html_path):
    print("Generating CSIT248_Lab_Record.pdf via Chrome headless...")
    pdf_path = os.path.join(BASE_DIR, "CSIT248_Lab_Record.pdf")
    chrome = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
    cmd = [
        chrome,
        "--headless",
        "--disable-gpu",
        "--no-pdf-header-footer",
        f"--print-to-pdf={pdf_path}",
        f"file://{html_path}"
    ]
    res = subprocess.run(cmd, capture_output=True)
    if os.path.exists(pdf_path):
        size_mb = os.path.getsize(pdf_path) / (1024 * 1024)
        print(f"SUCCESS: Generated PDF at {pdf_path} ({size_mb:.2f} MB)")
    else:
        print(f"ERROR generating PDF: {res.stderr.decode('utf-8', errors='ignore')}")

def set_cell_border(cell, **kwargs):
    """
    Set cell borders for python-docx table cell.
    kwargs: top, bottom, left, right with dict(sz=12, val='single', color='000000')
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        r'<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        r'<w:top w:val="single" w:sz="4" w:space="0" w:color="333333"/>'
        r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="333333"/>'
        r'<w:left w:val="single" w:sz="4" w:space="0" w:color="333333"/>'
        r'<w:right w:val="single" w:sz="4" w:space="0" w:color="333333"/>'
        r'</w:tcBorders>'
    )
    tcPr.append(tcBorders)

def generate_docx():
    print("Generating CSIT248_Lab_Record.docx via python-docx...")
    docx_path = os.path.join(BASE_DIR, "CSIT248_Lab_Record.docx")
    doc = docx.Document()

    # Set Margins (0.75 in top/bottom, 0.8 in left/right)
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # 1. COVER PAGE
    hdr_img = os.path.join(ASSETS_DIR, "amity_header.png")
    shd_img = os.path.join(ASSETS_DIR, "amity_shield.png")

    if os.path.exists(hdr_img):
        p_hdr = doc.add_paragraph()
        p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_hdr.add_run().add_picture(hdr_img, width=Inches(5.0))

    p_sp1 = doc.add_paragraph()
    p_sp1.paragraph_format.space_before = Pt(25)

    if os.path.exists(shd_img):
        p_shd = doc.add_paragraph()
        p_shd.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_shd.add_run().add_picture(shd_img, width=Inches(1.8))

    p_code = doc.add_paragraph()
    p_code.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_code = p_code.add_run(COURSE_CODE)
    r_code.bold = True
    r_code.font.size = Pt(20)
    r_code.font.name = "Arial"
    r_code.font.color.rgb = RGBColor(0, 0, 0)

    p_subj = doc.add_paragraph()
    p_subj.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_subj = p_subj.add_run(SUBJECT_NAME)
    r_subj.bold = True
    r_subj.font.size = Pt(18)
    r_subj.font.name = "Arial"
    r_subj.font.color.rgb = RGBColor(0, 0, 0)

    p_sp2 = doc.add_paragraph()
    p_sp2.paragraph_format.space_before = Pt(35)

    p_subby = doc.add_paragraph()
    p_subby.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_subby = p_subby.add_run("Submitted by")
    r_subby.font.size = Pt(13)
    r_subby.font.name = "Arial"
    r_subby.font.color.rgb = RGBColor(55, 65, 81)

    p_det = doc.add_paragraph()
    p_det.alignment = WD_ALIGN_PARAGRAPH.CENTER
    det_lines = [
        ("Name:", STUDENT_NAME),
        ("Enrollment No.:", ENROLLMENT_NO),
        ("Programme:", PROGRAMME),
        ("Semester:", SEMESTER)
    ]
    p_det_tbl = doc.add_table(rows=4, cols=2)
    p_det_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (k, v) in enumerate(det_lines):
        r = p_det_tbl.rows[i]
        c0, c1 = r.cells[0], r.cells[1]
        c0.width = Inches(1.8)
        c1.width = Inches(3.2)
        p0 = c0.paragraphs[0]
        r0 = p0.add_run(k)
        r0.bold = True
        r0.font.size = Pt(11.5)
        r0.font.name = "Arial"
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(v)
        r1.font.size = Pt(11.5)
        r1.font.name = "Arial"

    doc.add_page_break()

    # 2. INDEX PAGE
    p_idx = doc.add_paragraph()
    r_idx = p_idx.add_run("Index")
    r_idx.bold = True
    r_idx.font.size = Pt(24)
    r_idx.font.name = "Arial"

    tbl = doc.add_table(rows=1, cols=6)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False

    hdr_titles = ["S. No.", "Experiment No.", "Title of Experiment", "Date Performed", "Date Submitted", "Faculty Signature"]
    widths = [Inches(0.6), Inches(1.1), Inches(2.9), Inches(1.0), Inches(1.0), Inches(0.9)]

    hdr_cells = tbl.rows[0].cells
    for i, (cell, title, w) in enumerate(zip(hdr_cells, hdr_titles, widths)):
        cell.width = w
        set_cell_border(cell)
        p = cell.paragraphs[0]
        r = p.add_run(title)
        r.bold = True
        r.font.size = Pt(8.5)
        r.font.name = "Arial"
        if i in [0, 1, 3, 4, 5]:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for idx, exp in enumerate(EXPERIMENTS, 1):
        row_cells = tbl.add_row().cells
        for cell, w in zip(row_cells, widths):
            cell.width = w
            set_cell_border(cell)

        # S.No
        p0 = row_cells[0].paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r0 = p0.add_run(str(idx))
        r0.font.size = Pt(8.5)

        # Exp No
        p1 = row_cells[1].paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p1.add_run(exp["id"])
        r1.font.size = Pt(8.5)

        # Title
        p2 = row_cells[2].paragraphs[0]
        r2 = p2.add_run(exp["title"])
        r2.font.size = Pt(8.5)

        # Date Performed
        p3 = row_cells[3].paragraphs[0]
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r3 = p3.add_run(exp["date_performed"])
        r3.font.size = Pt(8.0)

        # Date Submitted
        p4 = row_cells[4].paragraphs[0]
        p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r4 = p4.add_run(exp["date_submitted"])
        r4.font.size = Pt(8.0)

    doc.add_page_break()

    # 3. EXPERIMENT DETAILS
    for exp in EXPERIMENTS:
        src_code = read_file_safe(exp["src_path"])

        p_num = doc.add_paragraph()
        r_num = p_num.add_run(f"Experiment No.: {exp['num_display']}")
        r_num.bold = True
        r_num.font.size = Pt(13)
        r_num.font.name = "Arial"
        p_num.paragraph_format.space_after = Pt(4)

        # Title (Red text)
        p_title = doc.add_paragraph()
        r_tlbl = p_title.add_run("Title : ")
        r_tlbl.bold = True
        r_tlbl.font.size = Pt(11)
        r_tlbl.font.name = "Arial"
        r_tval = p_title.add_run(exp["title"])
        r_tval.bold = True
        r_tval.font.size = Pt(11)
        r_tval.font.name = "Arial"
        r_tval.font.color.rgb = RGBColor(185, 28, 28) # Red
        p_title.paragraph_format.space_after = Pt(4)

        # Aim (Blue text)
        p_aim = doc.add_paragraph()
        r_albl = p_aim.add_run("Aim : ")
        r_albl.bold = True
        r_albl.font.size = Pt(11)
        r_albl.font.name = "Arial"
        r_aval = p_aim.add_run(exp["aim"])
        r_aval.bold = True
        r_aval.font.size = Pt(11)
        r_aval.font.name = "Arial"
        r_aval.font.color.rgb = RGBColor(29, 78, 216) # Blue
        p_aim.paragraph_format.space_after = Pt(8)

        # Source Code Heading
        p_sc = doc.add_paragraph()
        r_sc = p_sc.add_run("Source Code :")
        r_sc.bold = True
        r_sc.font.size = Pt(11)
        r_sc.font.name = "Arial"
        p_sc.paragraph_format.space_after = Pt(4)

        # Source Code text
        p_code = doc.add_paragraph()
        p_code.paragraph_format.line_spacing = 1.05
        p_code.paragraph_format.space_after = Pt(8)
        r_code = p_code.add_run(src_code)
        r_code.font.name = "Courier New"
        r_code.font.size = Pt(8.0)

        # Extra files if any (XML, PHP)
        if "extra_files" in exp:
            for extra_name, extra_path in exp["extra_files"]:
                p_exh = doc.add_paragraph()
                r_exh = p_exh.add_run(f"Supplementary Source File ({extra_name}) :")
                r_exh.bold = True
                r_exh.font.size = Pt(10)
                r_exh.font.color.rgb = RGBColor(67, 56, 202)
                p_exh.paragraph_format.space_after = Pt(2)

                p_exc = doc.add_paragraph()
                p_exc.paragraph_format.line_spacing = 1.05
                p_exc.paragraph_format.space_after = Pt(8)
                r_exc = p_exc.add_run(read_file_safe(extra_path))
                r_exc.font.name = "Courier New"
                r_exc.font.size = Pt(8.0)

        # Output Heading
        p_out = doc.add_paragraph()
        r_out = p_out.add_run("Output :")
        r_out.bold = True
        r_out.font.size = Pt(11)
        r_out.font.name = "Arial"
        p_out.paragraph_format.space_after = Pt(6)

        # Output Screenshot
        if os.path.exists(exp["img_path"]):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.add_run().add_picture(exp["img_path"], width=Inches(5.6))
            p_img.paragraph_format.space_after = Pt(14)

        doc.add_page_break()

    doc.save(docx_path)
    size_mb = os.path.getsize(docx_path) / (1024 * 1024)
    print(f"SUCCESS: Generated DOCX at {docx_path} ({size_mb:.2f} MB)")

if __name__ == "__main__":
    html_file = generate_html()
    generate_pdf(html_file)
    generate_docx()
    print("\n🎉 ALL GENERATION COMPLETE!")
