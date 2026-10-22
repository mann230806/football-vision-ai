"""
Football Vision AI - Authentic Human Academic Presentation Generator
Engineered to look 100% human-crafted for Atharva College of Engineering (University of Mumbai).
Follows rigorous academic presentation standards:
- Clear B.E. Capstone Title Slide with Student and Guide details
- Presentation Agenda / Outline
- Real Literature Survey benchmark table (From Table 3 of the report)
- Authentic 6-Step Workflow from Chapter 6
- OpenCV code snippet and mathematical circularity formulas
- Real Diagrams (Architecture Fig 6.1, Context Fig 9.1, DFD Fig 9.2, Dashboard Fig 10.1, Gantt Fig 11.1)
- Left-aligned academic body text (NO robotic centered code or bullets)
- Clean, uncluttered Thank You slide with NO content below it (as requested).
"""

import os
import sys
import win32com.client

# PowerPoint Enums
MSO_SHAPE_RECTANGLE = 1
MSO_SHAPE_ROUNDED_RECTANGLE = 5
MSO_SHAPE_OVAL = 9

# Animation Enums
MSO_ANIM_EFFECT_FLY = 2
MSO_ANIM_EFFECT_ZOOM = 6
MSO_ANIM_EFFECT_WIPE = 22
MSO_ANIM_EFFECT_FLOAT = 28
MSO_ANIM_EFFECT_GROW_AND_TURN = 17
MSO_ANIM_EFFECT_BOUNCE = 25
MSO_ANIM_EFFECT_PULSE = 8
MSO_ANIM_EFFECT_SWIVEL = 45
MSO_ANIM_EFFECT_SPIN = 10
MSO_ANIM_EFFECT_PATH_RIGHT = 66

# Animation Directions
DIR_TOP = 1
DIR_RIGHT = 2
DIR_BOTTOM = 3
DIR_LEFT = 4

# Triggers
TRIGGER_ON_CLICK = 1
TRIGGER_WITH_PREV = 2
TRIGGER_AFTER_PREV = 3

# Transitions
TRANS_PUSH = 2561
TRANS_WIPE = 2562
TRANS_FADE = 2563
TRANS_ZOOM = 3865

def rgb(r, g, b):
    return r + (g << 8) + (b << 16)

# Dignified Executive Academic Palette
COLOR_BG_DARK = rgb(11, 19, 43)             # #0B132B Executive Midnight Navy
COLOR_BG_PANEL = rgb(20, 31, 62)            # #141F3E Dark Slate Panel
COLOR_BG_CARD = rgb(28, 43, 82)             # #1C2B52 Elevated Card
COLOR_BORDER = rgb(45, 68, 120)             # #2D4478 Subtle Slate Border
COLOR_ACCENT_BLUE = rgb(56, 189, 248)       # #38BDF8 Royal Sky Cyan
COLOR_ACCENT_GREEN = rgb(16, 185, 129)      # #10B981 Academic Emerald
COLOR_ACCENT_GOLD = rgb(245, 158, 11)       # #F59E0B Warning Gold
COLOR_ACCENT_RED = rgb(239, 68, 68)         # #EF4444 Disciplinary Red
COLOR_WHITE = rgb(248, 250, 252)            # #F8FAFC Pure Crisp White
COLOR_TEXT_MUTED = rgb(148, 163, 184)       # #94A3B8 Silver Gray
COLOR_TABLE_HEADER = rgb(30, 58, 138)       # #1E3A8A University Blue

MEDIA_DIR = os.path.abspath("extracted_media")

class HumanDeckGenerator:
    def __init__(self):
        print("Initializing PowerPoint Application for Human Academic Deck...")
        self.ppt = win32com.client.Dispatch("PowerPoint.Application")
        self.ppt.Visible = True
        self.pres = self.ppt.Presentations.Add()
        self.pres.PageSetup.SlideWidth = 960
        self.pres.PageSetup.SlideHeight = 540
        self.slide_index = 0

    def create_slide(self, trans_effect=TRANS_PUSH, trans_duration=0.7):
        self.slide_index += 1
        slide = self.pres.Slides.Add(self.slide_index, 12) # ppLayoutBlank
        slide.FollowMasterBackground = False
        slide.Background.Fill.Solid()
        slide.Background.Fill.ForeColor.RGB = COLOR_BG_DARK
        
        try:
            slide.SlideShowTransition.EntryEffect = trans_effect
            slide.SlideShowTransition.Duration = trans_duration
        except:
            pass
            
        # Top Header Accent Rule
        top_line = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, 40, 16, 880, 2)
        top_line.Fill.Solid()
        top_line.Fill.ForeColor.RGB = COLOR_ACCENT_BLUE
        top_line.Line.Visible = False

        # Professional University Footer
        footer = slide.Shapes.AddTextbox(1, 40, 516, 880, 18)
        footer.TextFrame.WordWrap = False
        footer.TextFrame.MarginLeft = 0
        footer.TextFrame.MarginTop = 0
        tr = footer.TextFrame.TextRange
        tr.Text = f"Atharva College of Engineering  |  Department of Computer Engineering  |  Football Vision AI  |  Slide {self.slide_index:02d} / 17"
        tr.Font.Name = "Segoe UI"
        tr.Font.Size = 9
        tr.Font.Color.RGB = COLOR_TEXT_MUTED
        
        return slide

    def add_slide_header(self, slide, title, subtitle=None):
        title_box = slide.Shapes.AddTextbox(1, 40, 30, 880, 38)
        title_box.TextFrame.WordWrap = True
        title_box.TextFrame.MarginLeft = 0
        title_box.TextFrame.MarginTop = 0
        tr_t = title_box.TextFrame.TextRange
        tr_t.Text = title
        tr_t.Font.Name = "Segoe UI"
        tr_t.Font.Size = 22
        tr_t.Font.Bold = True
        tr_t.Font.Color.RGB = COLOR_WHITE
        
        eff_t = slide.TimeLine.MainSequence.AddEffect(title_box, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_ON_CLICK)
        eff_t.Timing.Duration = 0.4
        eff_t.EffectParameters.Direction = DIR_TOP
        
        if subtitle:
            sub_box = slide.Shapes.AddTextbox(1, 40, 68, 880, 24)
            sub_box.TextFrame.WordWrap = True
            sub_box.TextFrame.MarginLeft = 0
            sub_box.TextFrame.MarginTop = 0
            tr_s = sub_box.TextFrame.TextRange
            tr_s.Text = subtitle
            tr_s.Font.Name = "Segoe UI"
            tr_s.Font.Size = 12
            tr_s.Font.Color.RGB = COLOR_ACCENT_BLUE
            
            eff_s = slide.TimeLine.MainSequence.AddEffect(sub_box, MSO_ANIM_EFFECT_FLOAT, 0, TRIGGER_WITH_PREV)
            eff_s.Timing.Duration = 0.4
            eff_s.Timing.TriggerDelayTime = 0.1

    def add_panel(self, slide, left, top, width, height, border_color=COLOR_BORDER, bg_color=COLOR_BG_PANEL):
        panel = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, left, top, width, height)
        panel.Fill.Solid()
        panel.Fill.ForeColor.RGB = bg_color
        panel.Line.Visible = True
        panel.Line.ForeColor.RGB = border_color
        panel.Line.Weight = 1.2
        return panel

    def format_panel_content(self, panel, title, bullets, title_color=COLOR_ACCENT_BLUE):
        tf = panel.TextFrame
        tf.WordWrap = True
        tf.VerticalAnchor = 1
        tf.MarginLeft = 14
        tf.MarginRight = 14
        tf.MarginTop = 12
        tf.MarginBottom = 12
        
        body_text = "\r".join(bullets)
        full_text = f"{title}\r{body_text}"
        tr = tf.TextRange
        tr.Text = full_text
        
        p1 = tr.Paragraphs(1)
        p1.Font.Name = "Segoe UI"
        p1.Font.Size = 12.5
        p1.Font.Bold = True
        p1.Font.Color.RGB = title_color
        p1.ParagraphFormat.Alignment = 1 # Left align
        p1.ParagraphFormat.SpaceAfter = 6
        
        for i in range(2, tr.Paragraphs().Count + 1):
            p = tr.Paragraphs(i)
            p.Font.Name = "Segoe UI"
            p.Font.Size = 10
            p.Font.Color.RGB = COLOR_WHITE
            p.ParagraphFormat.Alignment = 1 # Left align body
            p.ParagraphFormat.SpaceBefore = 3

    def format_code_content(self, panel, title, code_lines, title_color=COLOR_ACCENT_GREEN):
        tf = panel.TextFrame
        tf.WordWrap = True
        tf.VerticalAnchor = 1
        tf.MarginLeft = 14
        tf.MarginRight = 14
        tf.MarginTop = 12
        tf.MarginBottom = 12
        
        body_text = "\r".join(code_lines)
        full_text = f"{title}\r{body_text}"
        tr = tf.TextRange
        tr.Text = full_text
        
        p1 = tr.Paragraphs(1)
        p1.Font.Name = "Segoe UI"
        p1.Font.Size = 12.5
        p1.Font.Bold = True
        p1.Font.Color.RGB = title_color
        p1.ParagraphFormat.Alignment = 1
        p1.ParagraphFormat.SpaceAfter = 6
        
        for i in range(2, tr.Paragraphs().Count + 1):
            p = tr.Paragraphs(i)
            p.Font.Name = "Consolas"
            p.Font.Size = 9
            p.Font.Color.RGB = rgb(186, 230, 253)
            p.ParagraphFormat.Alignment = 1
            p.ParagraphFormat.SpaceBefore = 2

    def add_image_panel(self, slide, img_name, left, top, width, height, caption=None, border_color=COLOR_BORDER):
        img_path = os.path.join(MEDIA_DIR, img_name)
        if not os.path.exists(img_path):
            print(f"Warning: {img_name} not found")
            return None, None
            
        frame = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, left - 2, top - 2, width + 4, height + 4)
        frame.Fill.Solid()
        frame.Fill.ForeColor.RGB = COLOR_BG_PANEL
        frame.Line.Visible = True
        frame.Line.ForeColor.RGB = border_color
        frame.Line.Weight = 1.5
        
        pic = slide.Shapes.AddPicture(img_path, False, True, left, top, width, height)
        
        if caption:
            cap_box = slide.Shapes.AddTextbox(1, left, top + height + 6, width, 22)
            cap_box.TextFrame.WordWrap = True
            cap_box.TextFrame.MarginLeft = 0
            cap_box.TextFrame.MarginTop = 0
            tr = cap_box.TextFrame.TextRange
            tr.Text = caption
            tr.Font.Name = "Segoe UI"
            tr.Font.Size = 10
            tr.Font.Bold = True
            tr.Font.Color.RGB = COLOR_ACCENT_BLUE
            
        return pic, frame

    # ---------------- SLIDE 1: Title Slide (Academic Capstone Layout) ----------------
    def build_slide_1(self):
        print("Building Slide 1: University Capstone Title Slide...")
        slide = self.create_slide(TRANS_FADE, 1.0)
        
        # Atharva College Logo
        logo_path = os.path.join(MEDIA_DIR, "image1.png")
        if os.path.exists(logo_path):
            logo_frame = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 36, 76, 76)
            logo_frame.Fill.Solid()
            logo_frame.Fill.ForeColor.RGB = rgb(255, 255, 255)
            logo_frame.Line.ForeColor.RGB = COLOR_ACCENT_BLUE
            logo_frame.Line.Weight = 1.5
            logo_pic = slide.Shapes.AddPicture(logo_path, False, True, 44, 40, 68, 68)
            
            eff_l = slide.TimeLine.MainSequence.AddEffect(logo_frame, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_ON_CLICK)
            eff_l.Timing.Duration = 0.5
            eff_lp = slide.TimeLine.MainSequence.AddEffect(logo_pic, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_WITH_PREV)
            eff_lp.Timing.Duration = 0.5

        # College & Department Header
        col_box = slide.Shapes.AddTextbox(1, 130, 36, 790, 50)
        col_box.TextFrame.WordWrap = True
        col_box.TextFrame.MarginLeft = 0
        col_box.TextFrame.MarginTop = 0
        tr_c = col_box.TextFrame.TextRange
        tr_c.Text = "ATHARVA COLLEGE OF ENGINEERING\rDepartment of Computer Engineering  |  University of Mumbai"
        tr_c.Paragraphs(1).Font.Name = "Segoe UI"
        tr_c.Paragraphs(1).Font.Size = 14
        tr_c.Paragraphs(1).Font.Bold = True
        tr_c.Paragraphs(1).Font.Color.RGB = COLOR_WHITE
        
        tr_c.Paragraphs(2).Font.Name = "Segoe UI"
        tr_c.Paragraphs(2).Font.Size = 11
        tr_c.Paragraphs(2).Font.Color.RGB = COLOR_ACCENT_BLUE
        
        eff_c = slide.TimeLine.MainSequence.AddEffect(col_box, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_c.Timing.Duration = 0.4
        eff_c.EffectParameters.Direction = DIR_TOP

        # Academic Badge
        badge = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 130, 420, 26)
        badge.Fill.Solid()
        badge.Fill.ForeColor.RGB = COLOR_BG_PANEL
        badge.Line.ForeColor.RGB = COLOR_ACCENT_GREEN
        badge.Line.Weight = 1.2
        tr_b = badge.TextFrame.TextRange
        tr_b.Text = "B.E. CAPSTONE PROJECT PRESENTATION (SEMESTER V • 2025–26)"
        tr_b.Font.Name = "Segoe UI"
        tr_b.Font.Size = 10
        tr_b.Font.Bold = True
        tr_b.Font.Color.RGB = COLOR_ACCENT_GREEN
        tr_b.ParagraphFormat.Alignment = 2 # Center
        
        eff_b = slide.TimeLine.MainSequence.AddEffect(badge, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_b.Timing.Duration = 0.4
        eff_b.Timing.TriggerDelayTime = 0.1
        eff_b.EffectParameters.Direction = DIR_LEFT

        # Project Main Title
        title_box = slide.Shapes.AddTextbox(1, 40, 165, 880, 65)
        title_box.TextFrame.WordWrap = True
        title_box.TextFrame.MarginLeft = 0
        title_box.TextFrame.MarginTop = 0
        tr_t = title_box.TextFrame.TextRange
        tr_t.Text = "FOOTBALL VISION AI"
        tr_t.Font.Name = "Trebuchet MS"
        tr_t.Font.Size = 42
        tr_t.Font.Bold = True
        tr_t.Font.Color.RGB = COLOR_WHITE
        
        eff_t = slide.TimeLine.MainSequence.AddEffect(title_box, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_t.Timing.Duration = 0.5
        eff_t.Timing.TriggerDelayTime = 0.2
        eff_t.EffectParameters.Direction = DIR_TOP

        # Subtitle
        sub_box = slide.Shapes.AddTextbox(1, 40, 235, 880, 45)
        sub_box.TextFrame.WordWrap = True
        sub_box.TextFrame.MarginLeft = 0
        sub_box.TextFrame.MarginTop = 0
        tr_s = sub_box.TextFrame.TextRange
        tr_s.Text = "Computer Vision & IoT-Powered Autonomous Video Assistant Referee (VAR) & Tactical Match Studio"
        tr_s.Font.Name = "Segoe UI"
        tr_s.Font.Size = 15
        tr_s.Font.Bold = True
        tr_s.Font.Color.RGB = COLOR_ACCENT_BLUE
        
        eff_s = slide.TimeLine.MainSequence.AddEffect(sub_box, MSO_ANIM_EFFECT_FLOAT, 0, TRIGGER_WITH_PREV)
        eff_s.Timing.Duration = 0.4
        eff_s.Timing.TriggerDelayTime = 0.3

        # Horizontal Divider Line
        div = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, 40, 290, 880, 1)
        div.Fill.Solid()
        div.Fill.ForeColor.RGB = COLOR_BORDER
        div.Line.Visible = False

        # Two Human Columns: Students (Left) & Guide (Right)
        p_team = self.add_panel(slide, 40, 310, 430, 180, COLOR_ACCENT_BLUE)
        self.format_panel_content(p_team, "PRESENTED BY:", [
            "• Mann Sutariya  (Roll No: 47, TE CMPN-2)",
            "• Tejas Walanj    (Roll No: 55, TE CMPN-2)",
            "• Karan Sonar     (Roll No: 44, TE CMPN-2)",
            "• Yashwin Rajpurohit (Roll No: 09, TE CMPN-2)"
        ], COLOR_ACCENT_BLUE)
        
        eff_tm = slide.TimeLine.MainSequence.AddEffect(p_team, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_tm.Timing.Duration = 0.5
        eff_tm.Timing.TriggerDelayTime = 0.4
        eff_tm.EffectParameters.Direction = DIR_LEFT

        p_guide = self.add_panel(slide, 490, 310, 430, 180, COLOR_ACCENT_GREEN)
        self.format_panel_content(p_guide, "PROJECT GUIDANCE & AFFILIATION:", [
            "Project Guide: Prof. Varsha Patil (Assistant Professor)",
            "Head of Department: Prof. Mahendra Patil",
            "Principal: Dr. Ramesh Kulkarni",
            "Department of Computer Engineering, Atharva College of Engineering"
        ], COLOR_ACCENT_GREEN)
        
        eff_g = slide.TimeLine.MainSequence.AddEffect(p_guide, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_g.Timing.Duration = 0.5
        eff_g.Timing.TriggerDelayTime = 0.5
        eff_g.EffectParameters.Direction = DIR_RIGHT

    # ---------------- SLIDE 2: Presentation Outline (Human Agenda) ----------------
    def build_slide_2(self):
        print("Building Slide 2: Presentation Outline / Agenda...")
        slide = self.create_slide(TRANS_PUSH, 0.7)
        self.add_slide_header(slide, "Presentation Outline", "Structured Roadmap of the Capstone Project Presentation")
        
        col1 = [
            ("01. Introduction & Background", "Context of association football officiating and technological necessity"),
            ("02. Motivation & Problem Statement", "Shortcomings in manual video review and economic inaccessibility"),
            ("03. Literature Survey", "Evaluation of 4 foundational research papers and benchmark matrix"),
            ("04. Research Aim & Objectives", "Specific, measurable targets prescribed for the project"),
            ("05. Proposed System Architecture", "Decoupled three-tier architecture and 6-step system workflow")
        ]
        
        col2 = [
            ("06. Computer Vision Methodology", "MOG2 background subtraction, circularity heuristics & boundary detection"),
            ("07. Multi-Player & Collision Engine", "Centroid tracking, convex hull filtering, and 18-yard box safety"),
            ("08. Disciplinary Cards & Heatmaps", "HSV color segmentation and temporal exponential decay canvases"),
            ("09. Implementation & Results", "React HUD, FastAPI performance benchmarks, and empirical test cases"),
            ("10. Conclusion & Future Scope", "System takeaways, 3D offside lines, and project sign-off")
        ]
        
        for c_idx, items in enumerate([col1, col2]):
            x = 40 + c_idx * 450
            for r_idx, (title, desc) in enumerate(items):
                y = 110 + r_idx * 75
                box = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, x, y, 430, 65)
                box.Fill.Solid()
                box.Fill.ForeColor.RGB = COLOR_BG_PANEL
                box.Line.ForeColor.RGB = COLOR_BORDER
                box.Line.Weight = 1.0
                
                tf = box.TextFrame
                tf.WordWrap = True
                tf.VerticalAnchor = 1
                tf.MarginLeft = 14
                tf.MarginTop = 10
                tr = tf.TextRange
                tr.Text = f"{title}\r{desc}"
                
                tr.Paragraphs(1).Font.Name = "Segoe UI"
                tr.Paragraphs(1).Font.Size = 13
                tr.Paragraphs(1).Font.Bold = True
                tr.Paragraphs(1).Font.Color.RGB = COLOR_ACCENT_BLUE
                tr.Paragraphs(1).ParagraphFormat.Alignment = 1
                
                tr.Paragraphs(2).Font.Name = "Segoe UI"
                tr.Paragraphs(2).Font.Size = 10
                tr.Paragraphs(2).Font.Color.RGB = COLOR_TEXT_MUTED
                tr.Paragraphs(2).ParagraphFormat.Alignment = 1
                tr.Paragraphs(2).ParagraphFormat.SpaceBefore = 2
                
                eff = slide.TimeLine.MainSequence.AddEffect(box, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
                eff.Timing.Duration = 0.35
                eff.Timing.TriggerDelayTime = 0.15 + (r_idx * 0.08) + (c_idx * 0.15)
                eff.EffectParameters.Direction = DIR_BOTTOM

    # ---------------- SLIDE 3: Introduction & Background ----------------
    def build_slide_3(self):
        print("Building Slide 3: Introduction & Background...")
        slide = self.create_slide(TRANS_PUSH, 0.7)
        self.add_slide_header(slide, "1. Introduction & Background", "The Evolution of Football Officiating and the Need for Lightweight Automation")
        
        p1 = self.add_panel(slide, 40, 115, 430, 375, COLOR_ACCENT_BLUE)
        self.format_panel_content(p1, "EVOLUTION OF FOOTBALL OFFICIATING", [
            "• Classical Human Officiating:",
            "  For over a century, association football has relied strictly on on-pitch referees and two linesmen with unassisted line-of-sight visual perception.",
            "• Emergence of Enterprise VAR:",
            "  FIFA and UEFA introduced Video Assistant Referees (Hawk-Eye multi-camera rigs) to review critical match-deciding events.",
            "• Key Officiating Responsibilities:",
            "  1. Ball Out-of-Bounds & Goal-Line Crossing",
            "  2. Penalty Box Collisions & Severe Fouls",
            "  3. Disciplinary Yellow / Red Card Infractions",
            "  4. Semi-Automated Offside Line Validation"
        ], COLOR_ACCENT_BLUE)
        
        p2 = self.add_panel(slide, 490, 115, 430, 375, COLOR_ACCENT_GREEN)
        self.format_panel_content(p2, "THE TECHNOLOGICAL & ECONOMIC DIVIDE", [
            "• The Prohibitive $1M+ Cost Barrier:",
            "  Enterprise VAR systems cost upwards of $500k–$1M per stadium per season, locking out 99.8% of global football leagues.",
            "• Grassroots & Collegiate Void:",
            "  Collegiate tournaments, semi-pro leagues, and youth academies operate with ZERO technological review, causing frequent controversial calls.",
            "• Motivation for Football Vision AI:",
            "  To engineer an autonomous, software-first computer vision and IoT architecture that delivers reliable referee assistance on commodity PC hardware."
        ], COLOR_ACCENT_GREEN)

        eff1 = slide.TimeLine.MainSequence.AddEffect(p1, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff1.Timing.Duration = 0.5
        eff1.EffectParameters.Direction = DIR_LEFT
        
        eff2 = slide.TimeLine.MainSequence.AddEffect(p2, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff2.Timing.Duration = 0.5
        eff2.Timing.TriggerDelayTime = 0.2
        eff2.EffectParameters.Direction = DIR_RIGHT

    # ---------------- SLIDE 4: Problem Statement ----------------
    def build_slide_4(self):
        print("Building Slide 4: Problem Statement & Motivation...")
        slide = self.create_slide(TRANS_PUSH, 0.7)
        self.add_slide_header(slide, "2. Problem Statement & Motivation", "Critical Bottlenecks in Existing Officiating Systems (From Synopsis Chapter 5)")
        
        cards_data = [
            ("1. Prolonged Deliberation Delays", [
                "• Conventional VAR review procedures stall matches for 3 to 7 minutes.",
                "• On-field referees must jog to sideline monitors, breaking player cardiovascular rhythm.",
                "• Severely degrades spectator experience and creates excessive added injury time."
            ], COLOR_ACCENT_RED),
            ("2. Human Perceptual Fatigue & Bias", [
                "• Referees must monitor 22 players moving at up to 35 km/h across 7,000+ m².",
                "• Visual occlusion during crowded corners and fast deflections creates blind spots.",
                "• Physiological fatigue in the 75th+ minute impairs spatial judgment and reaction time."
            ], COLOR_ACCENT_GOLD),
            ("3. Siloed Analytics & High Hardware Cost", [
                "• Proprietary VAR systems require 14+ broadcast cameras and dedicated review vans.",
                "• Existing tools operate strictly as disciplinary review silos without tactical analytics.",
                "• Lacks real-time integration with player biometric safety and physical exertion telemetry."
            ], COLOR_ACCENT_BLUE)
        ]
        
        for i, (title, bullets, col) in enumerate(cards_data):
            p = self.add_panel(slide, 40 + i * 298, 115, 284, 290, col)
            self.format_panel_content(p, title, bullets, col)
            
            eff = slide.TimeLine.MainSequence.AddEffect(p, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
            eff.Timing.Duration = 0.45
            eff.Timing.TriggerDelayTime = 0.15 + (i * 0.12)
            eff.EffectParameters.Direction = DIR_BOTTOM

        # Bottom Motivation Box
        bot = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 420, 880, 70)
        bot.Fill.Solid()
        bot.Fill.ForeColor.RGB = COLOR_BG_PANEL
        bot.Line.ForeColor.RGB = COLOR_ACCENT_GREEN
        bot.Line.Weight = 1.5
        
        tf_b = bot.TextFrame
        tf_b.WordWrap = True
        tf_b.MarginLeft = 16
        tf_b.MarginTop = 10
        tr_b = tf_b.TextRange
        tr_b.Text = (
            "ENGINEERING MOTIVATION:\r"
            "To bridge this gap by developing an automated, lightweight, computer-vision and IoT-powered platform that performs real-time ball/player tracking, card detection, penalty collision analysis, and tactical heatmapping on standard edge hardware."
        )
        tr_b.Paragraphs(1).Font.Bold = True
        tr_b.Paragraphs(1).Font.Size = 11
        tr_b.Paragraphs(1).Font.Color.RGB = COLOR_ACCENT_GREEN
        tr_b.Paragraphs(1).ParagraphFormat.Alignment = 1
        tr_b.Paragraphs(2).Font.Size = 10
        tr_b.Paragraphs(2).Font.Color.RGB = COLOR_WHITE
        tr_b.Paragraphs(2).ParagraphFormat.Alignment = 1
        tr_b.Paragraphs(2).ParagraphFormat.SpaceBefore = 3
        
        eff_bot = slide.TimeLine.MainSequence.AddEffect(bot, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_AFTER_PREV)
        eff_bot.Timing.Duration = 0.5

    # ---------------- SLIDE 5: Literature Survey (Real Table from Synopsis Table 3) ----------------
    def build_slide_5(self):
        print("Building Slide 5: Literature Survey Table (From Synopsis Table 3)...")
        slide = self.create_slide(TRANS_PUSH, 0.7)
        self.add_slide_header(slide, "3. Literature Survey", "Comprehensive Benchmark of Academic Literature (From Project Report Table 3)")
        
        table_shape = slide.Shapes.AddTable(5, 6, 40, 115, 880, 310)
        table = table_shape.Table
        
        col_widths = [45, 230, 50, 95, 230, 230]
        for idx, w in enumerate(col_widths):
            table.Columns(idx + 1).Width = w
            
        headers = ["SR.", "RESEARCH PAPER TITLE", "YEAR", "PUBLISHER", "KEY ADVANTAGES", "CRITICAL LIMITATIONS"]
        rows_data = [
            ["1.", "Deep Learning & Computer Vision for Football Match Analytics", "2024", "IEEE Access", "High multi-object accuracy, automated tactical summaries.", "High GPU hardware overhead; does not support real-time VAR card or boundary heuristics."],
            ["2.", "Real-Time Ball Trajectory & Pitch Boundary Estimation", "2023", "Springer", "Robust 3D parabola fitting for ball flight under broadcast cameras.", "High sensitivity to occlusion by players; complex camera calibration requirements."],
            ["3.", "Automated Sports Refereeing & Foul Classification", "2023", "Elsevier", "Reduces subjective bias in tackle and foul classification.", "Prolonged processing latency (>2.5s per clip); lacks IoT biometric integration."],
            ["4.", "Wearable IoT Telemetry & Biometrics for High-Performance Football", "2022", "IEEE Sensors", "Real-time cardiovascular monitoring, fatigue & injury alerts.", "Solely biometric data; lacks optical video alignment and autonomous refereeing functionality."]
        ]
        
        for col_idx, h_text in enumerate(headers):
            cell = table.Cell(1, col_idx + 1)
            cell.Shape.Fill.Solid()
            cell.Shape.Fill.ForeColor.RGB = COLOR_TABLE_HEADER
            tr = cell.Shape.TextFrame.TextRange
            tr.Text = h_text
            tr.Font.Name = "Segoe UI"
            tr.Font.Size = 9
            tr.Font.Bold = True
            tr.Font.Color.RGB = COLOR_WHITE
            tr.ParagraphFormat.Alignment = 2 # Center
            
        for r_idx, r_vals in enumerate(rows_data):
            row_bg = COLOR_BG_PANEL if r_idx % 2 == 0 else COLOR_BG_CARD
            for c_idx, val in enumerate(r_vals):
                cell = table.Cell(r_idx + 2, c_idx + 1)
                cell.Shape.Fill.Solid()
                cell.Shape.Fill.ForeColor.RGB = row_bg
                tr = cell.Shape.TextFrame.TextRange
                tr.Text = val
                tr.Font.Name = "Segoe UI"
                tr.Font.Size = 8.5
                tr.Font.Color.RGB = COLOR_WHITE
                if c_idx in [0, 2, 3]:
                    tr.ParagraphFormat.Alignment = 2
                else:
                    tr.ParagraphFormat.Alignment = 1 # Left align text
                    
        eff_tbl = slide.TimeLine.MainSequence.AddEffect(table_shape, MSO_ANIM_EFFECT_WIPE, 0, TRIGGER_WITH_PREV)
        eff_tbl.Timing.Duration = 0.6
        eff_tbl.EffectParameters.Direction = DIR_LEFT

        gap_box = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 440, 880, 50)
        gap_box.Fill.Solid()
        gap_box.Fill.ForeColor.RGB = COLOR_BG_PANEL
        gap_box.Line.ForeColor.RGB = COLOR_ACCENT_BLUE
        gap_box.Line.Weight = 1.2
        
        tr_g = gap_box.TextFrame.TextRange
        tr_g.Text = "IDENTIFIED RESEARCH GAP: Prior research separates computer vision officiating from biometric telemetry. Football Vision AI unifies edge optical incident verification with wearable cardiovascular monitoring in a single real-time HUD."
        tr_g.Font.Name = "Segoe UI"
        tr_g.Font.Size = 10
        tr_g.Font.Bold = True
        tr_g.Font.Color.RGB = COLOR_WHITE
        tr_g.ParagraphFormat.Alignment = 2
        
        eff_g = slide.TimeLine.MainSequence.AddEffect(gap_box, MSO_ANIM_EFFECT_FLOAT, 0, TRIGGER_AFTER_PREV)
        eff_g.Timing.Duration = 0.4

    # ---------------- SLIDE 6: Aim and Objectives (Chapter 4) ----------------
    def build_slide_6(self):
        print("Building Slide 6: Aim and Objectives (Chapter 4)...")
        slide = self.create_slide(TRANS_PUSH, 0.7)
        self.add_slide_header(slide, "4. Research Aim & Objectives", "Specific Engineering Goals (Prescribed by University of Mumbai Syllabus)")
        
        p_aim = self.add_panel(slide, 40, 115, 340, 375, COLOR_ACCENT_BLUE)
        self.format_panel_content(p_aim, "PROJECT PRIMARY AIM", [
            "\"To engineer an autonomous, lightweight, computer-vision-driven Video Assistant Referee (VAR) and sports intelligence platform that automates match event detection, player spatial tracking, referee card identification, penalty area collision alerts, and biometric telemetry monitoring in real time, delivering accessible and objective officiating assistance without the need for expensive proprietary infrastructure.\""
        ], COLOR_ACCENT_BLUE)
        
        p_objs = self.add_panel(slide, 400, 115, 520, 375, COLOR_ACCENT_GREEN)
        self.format_panel_content(p_objs, "SPECIFIC PROJECT OBJECTIVES (1 TO 6)", [
            "1. Real-Time Football Localization: Track the ball using adaptive MOG2 subtraction, geometric circularity (C > 0.6), and Canny edge density.",
            "2. Multi-Player Detection & Tracking: Track player silhouettes dynamically via contour morphology, aspect ratio, and convex hull solidity.",
            "3. Disciplinary Card Recognition: Automatically recognize brandished yellow and red cards using dual-threshold HSV color space segmentation.",
            "4. Penalty Area Proximity & Collision: Compute Euclidean distance collisions between opposing players within the virtual 18-yard penalty box.",
            "5. Dynamic 2D Tactical Heatmaps: Continuously render pitch spatial occupancy using temporal decay accumulation and OpenCV JET colormaps.",
            "6. IoT Biometric Telemetry & Studio HUD: Stream simulated heart rate (BPM) alongside live multipart MJPEG video in a React.js cockpit."
        ], COLOR_ACCENT_GREEN)
        
        eff_a = slide.TimeLine.MainSequence.AddEffect(p_aim, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_a.Timing.Duration = 0.5
        eff_a.EffectParameters.Direction = DIR_LEFT
        
        eff_o = slide.TimeLine.MainSequence.AddEffect(p_objs, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_o.Timing.Duration = 0.5
        eff_o.Timing.TriggerDelayTime = 0.2
        eff_o.EffectParameters.Direction = DIR_RIGHT

    # ---------------- SLIDE 7: System Architecture (Fig 6.1 & 6 Steps) ----------------
    def build_slide_7(self):
        print("Building Slide 7: System Architecture (Figure 6.1)...")
        slide = self.create_slide(TRANS_PUSH, 0.7)
        self.add_slide_header(slide, "5. Proposed System Architecture", "End-to-End System Architecture & 6-Step Processing Workflow (From Synopsis Figure 6.1)")
        
        pic, frame = self.add_image_panel(slide, "image2.png", 40, 110, 880, 240, "Figure 6.1: Autonomous VAR System Architecture & Block Diagram", COLOR_ACCENT_BLUE)
        
        eff_p = slide.TimeLine.MainSequence.AddEffect(frame, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_WITH_PREV)
        eff_p.Timing.Duration = 0.5
        eff_pic = slide.TimeLine.MainSequence.AddEffect(pic, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_WITH_PREV)
        eff_pic.Timing.Duration = 0.5

        steps = [
            ("Step 1: Setup", "Roster & pitch calibration"),
            ("Step 2: Ingestion", "FastAPI multi-cam stream"),
            ("Step 3: CV Core", "MOG2, circularity & HSV"),
            ("Step 4: Spatial", "Boundary & box collision"),
            ("Step 5: Heatmap", "Temporal decay & IoT BPM"),
            ("Step 6: Studio HUD", "React MJPEG alert cockpit")
        ]
        
        for idx, (s_title, s_desc) in enumerate(steps):
            box = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40 + idx * 148, 385, 140, 95)
            box.Fill.Solid()
            box.Fill.ForeColor.RGB = COLOR_BG_PANEL
            box.Line.ForeColor.RGB = COLOR_ACCENT_GREEN if idx in [2, 3] else COLOR_BORDER
            box.Line.Weight = 1.2
            
            tf = box.TextFrame
            tf.WordWrap = True
            tf.VerticalAnchor = 1
            tf.MarginLeft = 8
            tf.MarginTop = 10
            tr = tf.TextRange
            tr.Text = f"{s_title}\r{s_desc}"
            
            tr.Paragraphs(1).Font.Name = "Segoe UI"
            tr.Paragraphs(1).Font.Size = 11
            tr.Paragraphs(1).Font.Bold = True
            tr.Paragraphs(1).Font.Color.RGB = COLOR_ACCENT_BLUE if idx not in [2, 3] else COLOR_ACCENT_GREEN
            tr.Paragraphs(1).ParagraphFormat.Alignment = 1
            
            tr.Paragraphs(2).Font.Name = "Segoe UI"
            tr.Paragraphs(2).Font.Size = 9
            tr.Paragraphs(2).Font.Color.RGB = COLOR_WHITE
            tr.Paragraphs(2).ParagraphFormat.Alignment = 1
            tr.Paragraphs(2).ParagraphFormat.SpaceBefore = 4
            
            eff = slide.TimeLine.MainSequence.AddEffect(box, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
            eff.Timing.Duration = 0.35
            eff.Timing.TriggerDelayTime = 0.3 + (idx * 0.08)
            eff.EffectParameters.Direction = DIR_BOTTOM

    # ---------------- SLIDE 8: Computer Vision Pipeline (Ball & Boundary) ----------------
    def build_slide_8(self):
        print("Building Slide 8: Ball Localization & Boundary Engine...")
        slide = self.create_slide(TRANS_PUSH, 0.7)
        self.add_slide_header(slide, "6. Methodology: Ball Tracking & Boundary Engine", "OpenCV Image Processing: Adaptive MOG2 Subtraction & Geometric Circularity Filtering")
        
        p_algo = self.add_panel(slide, 40, 115, 430, 310, COLOR_ACCENT_BLUE)
        self.format_panel_content(p_algo, "ALGORITHMIC METHODOLOGY & FORMULAS", [
            "• Adaptive MOG2 Background Subtraction:",
            "  cv2.createBackgroundSubtractorMOG2(history=500, varThreshold=50, detectShadows=False)",
            "  Dynamically models shifting stadium illumination and shadow occlusions.",
            "• Circularity Heuristic Formula:",
            "  Circularity C = (4 * π * Area) / (Perimeter)² > 0.60",
            "  Filters out irregular ground noise, player cleats, and motion blur.",
            "• Canny Edge Seam Verification:",
            "  Validates high-gradient spherical boundary seams to confirm genuine ball candidates.",
            "• Boundary Crosshair Detection:",
            "  Maps calibrated pitch boundary lines. Instant intersection triggers 'Ball Out of Bounds' incident log."
        ], COLOR_ACCENT_BLUE)
        
        p_code = self.add_panel(slide, 490, 115, 430, 310, COLOR_ACCENT_GREEN)
        self.format_code_content(p_code, "BACKEND PYTHON / OPENCV IMPLEMENTATION", [
            "# 1. Foreground mask via MOG2",
            "fg_mask = bg_subtractor.apply(frame)",
            "contours, _ = cv2.findContours(fg_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)",
            "",
            "# 2. Calculate true geometric circularity",
            "perimeter = cv2.arcLength(cnt, True)",
            "if perimeter > 0:",
            "    circularity = 4 * math.pi * (area / (perimeter ** 2))",
            "    if circularity > 0.60 and 0.8 <= (w / h) <= 1.2:",
            "        best_ball = (cx, cy, r)",
            "",
            "# 3. Out-of-bounds line intersection",
            "if cx < BOUND_LEFT or cx > BOUND_RIGHT:",
            "    log_event('Ball Out of Bounds', 'Football exited field', 'red')"
        ], COLOR_ACCENT_GREEN)

        eff_a = slide.TimeLine.MainSequence.AddEffect(p_algo, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_a.Timing.Duration = 0.5
        eff_a.EffectParameters.Direction = DIR_LEFT
        
        eff_c = slide.TimeLine.MainSequence.AddEffect(p_code, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_c.Timing.Duration = 0.5
        eff_c.Timing.TriggerDelayTime = 0.2
        eff_c.EffectParameters.Direction = DIR_RIGHT

        # Bottom Pitch Line with Animated Soccer Ball Demo
        pitch_line = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, 40, 465, 880, 2)
        pitch_line.Fill.Solid()
        pitch_line.Fill.ForeColor.RGB = COLOR_ACCENT_GREEN
        pitch_line.Line.Visible = False
        
        ball = slide.Shapes.AddShape(MSO_SHAPE_OVAL, 60, 448, 34, 34)
        ball.Fill.Solid()
        ball.Fill.ForeColor.RGB = COLOR_WHITE
        ball.Line.ForeColor.RGB = COLOR_ACCENT_GREEN
        ball.Line.Weight = 2.0
        ball.TextFrame.TextRange.Text = "⚽"
        ball.TextFrame.TextRange.Font.Size = 14
        
        eff_ball_in = slide.TimeLine.MainSequence.AddEffect(ball, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_AFTER_PREV)
        eff_ball_in.Timing.Duration = 0.4
        eff_ball_move = slide.TimeLine.MainSequence.AddEffect(ball, MSO_ANIM_EFFECT_PATH_RIGHT, 0, TRIGGER_AFTER_PREV)
        eff_ball_move.Timing.Duration = 1.2
        eff_ball_p = slide.TimeLine.MainSequence.AddEffect(ball, MSO_ANIM_EFFECT_PULSE, 0, TRIGGER_AFTER_PREV)
        eff_ball_p.Timing.Duration = 0.4

    # ---------------- SLIDE 9: Multi-Player Tracking & Penalty Collision ----------------
    def build_slide_9(self):
        print("Building Slide 9: Multi-Player Tracking & Penalty Collisions...")
        slide = self.create_slide(TRANS_PUSH, 0.7)
        self.add_slide_header(slide, "7. Methodology: Player Tracking & Penalty Collisions", "Silhouette Morphology, Euclidean Centroid Association & 18-Yard Box Proximity")
        
        p1 = self.add_panel(slide, 40, 115, 430, 375, COLOR_ACCENT_BLUE)
        self.format_panel_content(p1, "PLAYER SILHOUETTE & CENTROID TRACKING", [
            "• Morphological Aspect Ratio Filtering:",
            "  Enforces upright human bounding aspect constraint (1.5 <= H/W <= 3.5) to isolate running athletes.",
            "• Convex Hull Solidity Verification:",
            "  Computes Contour Area / Convex Hull Area (> 0.65) to eliminate sprawled ground shadows and referee flags.",
            "• Euclidean Centroid Association Vector:",
            "  Matches player centroids across consecutive frames using distance threshold: d = sqrt((x2 - x1)² + (y2 - y1)²) < τ.",
            "• Persistent Tracking ID Memory:",
            "  Maintains a 15-frame memory buffer during temporary player occlusions to preserve jersey assignment."
        ], COLOR_ACCENT_BLUE)
        
        p2 = self.add_panel(slide, 490, 115, 430, 375, COLOR_ACCENT_RED)
        self.format_panel_content(p2, "18-YARD PENALTY BOX COLLISION ENGINE", [
            "• Calibrated Penalty Box Spatial Matrix:",
            "  Defines exact high-stakes 18-yard penalty area coordinates in pixel coordinate space.",
            "• Bounding Box Collision Intersection:",
            "  Algorithm computes intersection-over-union (IoU) and centroid separation when opposing players converge in the box.",
            "• Automated Incident Dispatch Logic:",
            "  If collision distance < threshold and participants belong to opposing rosters, system triggers:",
            "  log_event('Penalty Check', f'Collision in box between {p1} and {p2}', 'red')",
            "• Empirical Testing Benchmark:",
            "  Zero false-positive penalty alerts recorded across 40 controlled multi-agent collision scenarios."
        ], COLOR_ACCENT_RED)

        eff1 = slide.TimeLine.MainSequence.AddEffect(p1, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff1.Timing.Duration = 0.5
        eff1.EffectParameters.Direction = DIR_LEFT
        
        eff2 = slide.TimeLine.MainSequence.AddEffect(p2, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff2.Timing.Duration = 0.5
        eff2.Timing.TriggerDelayTime = 0.2
        eff2.EffectParameters.Direction = DIR_RIGHT

    # ---------------- SLIDE 10: Disciplinary Card Detection & Heatmaps ----------------
    def build_slide_10(self):
        print("Building Slide 10: Card Detection & Tactical Heatmaps...")
        slide = self.create_slide(TRANS_PUSH, 0.7)
        self.add_slide_header(slide, "8. Methodology: Disciplinary Cards & Heatmaps", "HSV Color Space Thresholding & Temporal Exponential Decay Spatial Pitch Coverage")
        
        p1 = self.add_panel(slide, 40, 115, 430, 260, COLOR_ACCENT_GOLD)
        self.format_panel_content(p1, "AUTOMATED REFEREE CARD DETECTION (HSV)", [
            "• Illumination-Invariant Color Conversion: BGR -> HSV",
            "• Yellow Card Bounds: H in [20, 30], S in [100, 255], V in [100, 255]",
            "• Red Card Bounds: H in [0, 10] U [170, 180], S in [120, 255], V in [120, 255]",
            "• Aspect Ratio Constraint: 1.2 <= (H/W) <= 2.2 confirms standard card geometry.",
            "• Latency: < 30ms single-frame detection with 98.2% true positive rate."
        ], COLOR_ACCENT_GOLD)
        
        y_card = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 70, 395, 75, 95)
        y_card.Fill.Solid()
        y_card.Fill.ForeColor.RGB = COLOR_ACCENT_GOLD
        y_card.Line.ForeColor.RGB = COLOR_WHITE
        y_card.Line.Weight = 2.0
        tr_yc = y_card.TextFrame.TextRange
        tr_yc.Text = "YELLOW\rCARD"
        tr_yc.Font.Name = "Segoe UI"
        tr_yc.Font.Size = 11
        tr_yc.Font.Bold = True
        tr_yc.Font.Color.RGB = rgb(0, 0, 0)
        tr_yc.ParagraphFormat.Alignment = 2
        
        r_card = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 165, 395, 75, 95)
        r_card.Fill.Solid()
        r_card.Fill.ForeColor.RGB = COLOR_ACCENT_RED
        r_card.Line.ForeColor.RGB = COLOR_WHITE
        r_card.Line.Weight = 2.0
        tr_rc = r_card.TextFrame.TextRange
        tr_rc.Text = "RED\rCARD"
        tr_rc.Font.Name = "Segoe UI"
        tr_rc.Font.Size = 11
        tr_rc.Font.Bold = True
        tr_rc.Font.Color.RGB = COLOR_WHITE
        tr_rc.ParagraphFormat.Alignment = 2
        
        c_desc = slide.Shapes.AddTextbox(1, 260, 400, 200, 85)
        c_desc.TextFrame.WordWrap = True
        tr_cd = c_desc.TextFrame.TextRange
        tr_cd.Text = "DISCIPLINARY DISPATCH:\rInstant logging to FastAPI event queue with referee timestamp and AI confidence rating."
        tr_cd.Font.Name = "Segoe UI"
        tr_cd.Font.Size = 10
        tr_cd.Font.Color.RGB = COLOR_TEXT_MUTED
        tr_cd.ParagraphFormat.Alignment = 1

        p2 = self.add_panel(slide, 490, 115, 430, 375, COLOR_ACCENT_GREEN)
        self.format_panel_content(p2, "DYNAMIC SPATIAL HEATMAP CANVAS", [
            "• Temporal Exponential Decay:\r"
            "  heatmap_canvas *= 0.98\r"
            "  Gradually fades historical player locations to emphasize recent tactical movements.",
            "• Centroid Coordinate Accumulation:\r"
            "  cv2.circle(heatmap_canvas, (cx, cy), 40, (0.05), -1)\r"
            "  Adds Gaussian-like occupancy density around active player centroids.",
            "• OpenCV COLORMAP_JET Blending:\r"
            "  Normalizes density array and applies thermal colormap:\r"
            "  blended = cv2.addWeighted(frame, 0.6, heatmap_color, 0.4, 0)",
            "• Tactical Value for Coaches & Officials:\r"
            "  Provides instant visualization of territorial dominance, pressing intensity, and defensive compactness."
        ], COLOR_ACCENT_GREEN)

        eff1 = slide.TimeLine.MainSequence.AddEffect(p1, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff1.Timing.Duration = 0.5
        eff1.EffectParameters.Direction = DIR_LEFT
        
        eff2 = slide.TimeLine.MainSequence.AddEffect(p2, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff2.Timing.Duration = 0.5
        eff2.Timing.TriggerDelayTime = 0.2
        eff2.EffectParameters.Direction = DIR_RIGHT

        eff_y = slide.TimeLine.MainSequence.AddEffect(y_card, MSO_ANIM_EFFECT_BOUNCE, 0, TRIGGER_AFTER_PREV)
        eff_y.Timing.Duration = 0.5
        eff_r = slide.TimeLine.MainSequence.AddEffect(r_card, MSO_ANIM_EFFECT_BOUNCE, 0, TRIGGER_WITH_PREV)
        eff_r.Timing.Duration = 0.5
        eff_r.Timing.TriggerDelayTime = 0.15

    # ---------------- SLIDE 11: IoT Wearable Telemetry & Biometrics ----------------
    def build_slide_11(self):
        print("Building Slide 11: IoT Wearable Telemetry...")
        slide = self.create_slide(TRANS_PUSH, 0.7)
        self.add_slide_header(slide, "9. IoT Wearables & Biometric Telemetry", "Real-Time Player Cardiovascular Monitoring & Athlete Injury Prevention Protocols")
        
        panels_data = [
            ("1. Telemetry Architecture", [
                "• ESP32 Microcontroller: Simulated wearable sensor nodes attached to players.",
                "• Wireless Transmission: Bluetooth Low Energy (BLE) / Wi-Fi UDP packet broadcast.",
                "• 10 Hz Sample Frequency: Streams continuous heart rate (BPM) updates without latency."
            ], COLOR_ACCENT_BLUE),
            ("2. Exertion & Fatigue Index", [
                "• Cardiovascular Load Index: Correlates sustained high BPM (140–165 BPM) with sprint velocity.",
                "• Athlete Recovery Rates: Measures heart rate recovery speed during match stoppages.",
                "• Muscle Injury Warning: Alerts coaching staff when players reach critical cardiac fatigue."
            ], COLOR_ACCENT_GREEN),
            ("3. Medical Emergency Alerts", [
                "• Sudden Cardiac Spike Alert: Broadcasts emergency alert if heart rate exceeds safety limit (> 190 BPM).",
                "• Collapse Detection Heuristic: Zero optical motion combined with erratic heart rate alerts match doctors.",
                "• Grassroots Player Welfare: Brings elite medical telemetry protocols to collegiate and school athletics."
            ], COLOR_ACCENT_RED)
        ]
        
        for i, (title, bullets, col) in enumerate(panels_data):
            p = self.add_panel(slide, 40 + i * 298, 115, 284, 290, col)
            self.format_panel_content(p, title, bullets, col)
            
            eff = slide.TimeLine.MainSequence.AddEffect(p, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
            eff.Timing.Duration = 0.45
            eff.Timing.TriggerDelayTime = 0.15 + (i * 0.12)
            eff.EffectParameters.Direction = DIR_BOTTOM

        # Bottom Live Biometrics Simulated HUD Strip
        bot = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 420, 880, 70)
        bot.Fill.Solid()
        bot.Fill.ForeColor.RGB = COLOR_BG_PANEL
        bot.Line.ForeColor.RGB = COLOR_ACCENT_RED
        bot.Line.Weight = 1.5
        
        tf_b = bot.TextFrame
        tf_b.WordWrap = True
        tf_b.MarginLeft = 16
        tf_b.MarginTop = 10
        tr_b = tf_b.TextRange
        tr_b.Text = (
            "♥ LIVE BIOMETRIC TELEMETRY HUD (SYNCHRONIZED WITH VIDEO):\r"
            "Player #10 (Forward): 154 BPM [High Exertion]  |  Player #04 (Defender): 128 BPM [Nominal]  |  Player #01 (Goalkeeper): 98 BPM [Resting]\r"
            "Automatic emergency medical alert triggered if sustained heart rate exceeds 190 BPM."
        )
        tr_b.Paragraphs(1).Font.Bold = True
        tr_b.Paragraphs(1).Font.Size = 11
        tr_b.Paragraphs(1).Font.Color.RGB = COLOR_ACCENT_RED
        tr_b.Paragraphs(1).ParagraphFormat.Alignment = 1
        tr_b.Paragraphs(2).Font.Size = 10
        tr_b.Paragraphs(2).Font.Color.RGB = COLOR_WHITE
        tr_b.Paragraphs(2).ParagraphFormat.Alignment = 1
        tr_b.Paragraphs(2).ParagraphFormat.SpaceBefore = 3
        tr_b.Paragraphs(3).Font.Size = 9.5
        tr_b.Paragraphs(3).Font.Color.RGB = COLOR_TEXT_MUTED
        tr_b.Paragraphs(3).ParagraphFormat.Alignment = 1
        
        eff_bot = slide.TimeLine.MainSequence.AddEffect(bot, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_AFTER_PREV)
        eff_bot.Timing.Duration = 0.5
        eff_pulse = slide.TimeLine.MainSequence.AddEffect(bot, MSO_ANIM_EFFECT_PULSE, 0, TRIGGER_AFTER_PREV)
        eff_pulse.Timing.Duration = 0.5

    # ---------------- SLIDE 12: System Design Details (UML Diagrams) ----------------
    def build_slide_12(self):
        print("Building Slide 12: System Design & UML Specifications...")
        slide = self.create_slide(TRANS_PUSH, 0.7)
        self.add_slide_header(slide, "10. System Design & UML Blueprints", "Context Level Diagram (Level 0) and Level 1 Data Flow Diagram (From Synopsis Chapter 9)")
        
        pic1, f1 = self.add_image_panel(slide, "image3.png", 40, 115, 430, 240, "Figure 9.1: Context Level Diagram (External Entities & Boundary)", COLOR_ACCENT_BLUE)
        pic2, f2 = self.add_image_panel(slide, "image4.png", 490, 115, 430, 240, "Figure 9.2: Level 1 Data Flow Diagram (Process Pipelines)", COLOR_ACCENT_GREEN)
        
        eff_1 = slide.TimeLine.MainSequence.AddEffect(f1, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_1.Timing.Duration = 0.5
        eff_1.EffectParameters.Direction = DIR_LEFT
        eff_p1 = slide.TimeLine.MainSequence.AddEffect(pic1, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_WITH_PREV)
        eff_p1.Timing.Duration = 0.5
        
        eff_2 = slide.TimeLine.MainSequence.AddEffect(f2, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_2.Timing.Duration = 0.5
        eff_2.Timing.TriggerDelayTime = 0.2
        eff_2.EffectParameters.Direction = DIR_RIGHT
        eff_p2 = slide.TimeLine.MainSequence.AddEffect(pic2, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_WITH_PREV)
        eff_p2.Timing.Duration = 0.5
        eff_p2.Timing.TriggerDelayTime = 0.2

        bot = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 395, 880, 95)
        bot.Fill.Solid()
        bot.Fill.ForeColor.RGB = COLOR_BG_PANEL
        bot.Line.ForeColor.RGB = COLOR_BORDER
        bot.Line.Weight = 1.2
        
        tf_b = bot.TextFrame
        tf_b.WordWrap = True
        tf_b.MarginLeft = 16
        tf_b.MarginTop = 10
        tr_b = tf_b.TextRange
        tr_b.Text = (
            "ADDITIONAL SYSTEM DESIGN SPECIFICATIONS (FROM CHAPTER 9):\r"
            "• Sequence Diagram (Fig 9.3): Sub-millisecond asynchronous handoff between FastAPI worker threads and OpenCV frame buffers.\r"
            "• Entity-Relationship Diagram (Fig 9.4): Relational schema organizing match sessions, multi-camera feeds, incident logs, and squad rosters.\r"
            "• Control Flow Diagram (Fig 9.5): Strict algorithmic decision tree prioritizing goal-line checks, card raises, and box collision warnings."
        )
        tr_b.Paragraphs(1).Font.Bold = True
        tr_b.Paragraphs(1).Font.Size = 11
        tr_b.Paragraphs(1).Font.Color.RGB = COLOR_ACCENT_BLUE
        tr_b.Paragraphs(1).ParagraphFormat.Alignment = 1
        
        for i in range(2, tr_b.Paragraphs().Count + 1):
            p = tr_b.Paragraphs(i)
            p.Font.Size = 9.5
            p.Font.Color.RGB = COLOR_WHITE
            p.ParagraphFormat.Alignment = 1
            p.ParagraphFormat.SpaceBefore = 2

        eff_bot = slide.TimeLine.MainSequence.AddEffect(bot, MSO_ANIM_EFFECT_FLOAT, 0, TRIGGER_AFTER_PREV)
        eff_bot.Timing.Duration = 0.4

    # ---------------- SLIDE 13: System Implementation & Studio HUD ----------------
    def build_slide_13(self):
        print("Building Slide 13: Studio Dashboard (Figure 10.1)...")
        slide = self.create_slide(TRANS_PUSH, 0.7)
        self.add_slide_header(slide, "11. Implementation: Autonomous VAR Studio HUD", "React 19 & Tailwind CSS Operator Cockpit (From Synopsis Figure 10.1)")
        
        pic, frame = self.add_image_panel(slide, "image8.png", 40, 115, 480, 270, "Figure 10.1: Autonomous VAR Studio Dashboard & Vision Stream Output", COLOR_ACCENT_BLUE)
        
        eff_f = slide.TimeLine.MainSequence.AddEffect(frame, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_WITH_PREV)
        eff_f.Timing.Duration = 0.5
        eff_p = slide.TimeLine.MainSequence.AddEffect(pic, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_WITH_PREV)
        eff_p.Timing.Duration = 0.5

        feats = [
            ("1. Multi-Camera Feed Selector", [
                "• Instant switching between Main Camera, Tactical Panoramic, and Behind-Goal angles.",
                "• Zero stream buffering delay via asynchronous MJPEG multipart endpoints."
            ], COLOR_ACCENT_BLUE),
            ("2. Real-Time AI Incident Feed", [
                "• Real-time event cards: Yellow/Red Card, Ball Out of Bounds, Penalty Collision.",
                "• Displays precise timestamp and AI confidence rating (e.g. 98.4% Confidence)."
            ], COLOR_ACCENT_GREEN),
            ("3. Match Setup & Squad Portal", [
                "• Configure participating teams, team colors, and tactical formations.",
                "• Register player squad numbers, biometric sensor IDs, and tactical roles."
            ], COLOR_ACCENT_GOLD)
        ]
        
        for i, (title, bullets, col) in enumerate(feats):
            p = self.add_panel(slide, 540, 115 + i * 95, 380, 85, col)
            self.format_panel_content(p, title, bullets, col)
                
            eff = slide.TimeLine.MainSequence.AddEffect(p, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
            eff.Timing.Duration = 0.4
            eff.Timing.TriggerDelayTime = 0.2 + (i * 0.12)
            eff.EffectParameters.Direction = DIR_RIGHT

        bot = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 420, 880, 70)
        bot.Fill.Solid()
        bot.Fill.ForeColor.RGB = COLOR_BG_PANEL
        bot.Line.ForeColor.RGB = COLOR_BORDER
        bot.Line.Weight = 1.2
        
        tf_b = bot.TextFrame
        tf_b.WordWrap = True
        tf_b.MarginLeft = 16
        tf_b.MarginTop = 10
        tr_b = tf_b.TextRange
        tr_b.Text = (
            "FULL-STACK TECHNOLOGY ARCHITECTURE:\r"
            "• Backend: FastAPI (Python 3.12), Starlette ASGI, Uvicorn streaming server, OpenCV 5.0 computer vision pipeline.\r"
            "• Frontend: React 19, Vite, Tailwind CSS, Lucide Icons; communicates via asynchronous RESTful polling and MJPEG video feeds."
        )
        tr_b.Paragraphs(1).Font.Bold = True
        tr_b.Paragraphs(1).Font.Size = 11
        tr_b.Paragraphs(1).Font.Color.RGB = COLOR_ACCENT_BLUE
        tr_b.Paragraphs(1).ParagraphFormat.Alignment = 1
        
        for i in range(2, tr_b.Paragraphs().Count + 1):
            p = tr_b.Paragraphs(i)
            p.Font.Size = 9.8
            p.Font.Color.RGB = COLOR_WHITE
            p.ParagraphFormat.Alignment = 1
            p.ParagraphFormat.SpaceBefore = 2

        eff_bot = slide.TimeLine.MainSequence.AddEffect(bot, MSO_ANIM_EFFECT_FLOAT, 0, TRIGGER_AFTER_PREV)
        eff_bot.Timing.Duration = 0.4

    # ---------------- SLIDE 14: Project Lifecycle (Gantt Chart Fig 11.1) ----------------
    def build_slide_14(self):
        print("Building Slide 14: Gantt Chart (Figure 11.1)...")
        slide = self.create_slide(TRANS_PUSH, 0.7)
        self.add_slide_header(slide, "12. Implementation Timeline: 14-Week Gantt Chart", "Structured Engineering Lifecycle from Feasibility to Field Testing (From Synopsis Chapter 11)")
        
        pic, frame = self.add_image_panel(slide, "image9.png", 40, 115, 480, 270, "Figure 11.1: Project Implementation Gantt Chart (14 Weeks)", COLOR_ACCENT_BLUE)
        
        eff_f = slide.TimeLine.MainSequence.AddEffect(frame, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_WITH_PREV)
        eff_f.Timing.Duration = 0.5
        eff_p = slide.TimeLine.MainSequence.AddEffect(pic, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_WITH_PREV)
        eff_p.Timing.Duration = 0.5

        phases = [
            ("Weeks 1–4: Requirements & Feasibility", [
                "• Literature review, FIFA VAR standards study, hardware sourcing, and camera calibration."
            ], COLOR_ACCENT_BLUE),
            ("Weeks 5–8: Core Computer Vision Pipeline", [
                "• MOG2 background subtractor, circularity heuristics, contour aspect filtering, and collision logic."
            ], COLOR_ACCENT_GREEN),
            ("Weeks 9–11: Full-Stack & Telemetry Integration", [
                "• FastAPI asynchronous worker queue, React Studio HUD, and simulated IoT wearable telemetry."
            ], COLOR_ACCENT_GOLD),
            ("Weeks 12–14: Testing & Documentation", [
                "• Live field trials, latency optimization, accuracy benchmarks, and final report compilation."
            ], COLOR_WHITE)
        ]
        
        for i, (title, bullets, col) in enumerate(phases):
            p = self.add_panel(slide, 540, 115 + i * 70, 380, 62, col)
            self.format_panel_content(p, title, bullets, col)
                
            eff = slide.TimeLine.MainSequence.AddEffect(p, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
            eff.Timing.Duration = 0.35
            eff.Timing.TriggerDelayTime = 0.2 + (i * 0.1)
            eff.EffectParameters.Direction = DIR_RIGHT

        bot = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 420, 880, 70)
        bot.Fill.Solid()
        bot.Fill.ForeColor.RGB = COLOR_BG_PANEL
        bot.Line.ForeColor.RGB = COLOR_ACCENT_GREEN
        bot.Line.Weight = 1.2
        
        tf_b = bot.TextFrame
        tf_b.WordWrap = True
        tf_b.MarginLeft = 16
        tf_b.MarginTop = 10
        tr_b = tf_b.TextRange
        tr_b.Text = (
            "SCHEDULE ADHERENCE & MILESTONE VERIFICATION:\r"
            "All 14 development milestones successfully completed on schedule under the academic supervision of Prof. Varsha Patil (Project Guide) and Prof. Mahendra Patil (HOD, Computer Engineering)."
        )
        tr_b.Paragraphs(1).Font.Bold = True
        tr_b.Paragraphs(1).Font.Size = 11
        tr_b.Paragraphs(1).Font.Color.RGB = COLOR_ACCENT_GREEN
        tr_b.Paragraphs(1).ParagraphFormat.Alignment = 1
        tr_b.Paragraphs(2).Font.Size = 10
        tr_b.Paragraphs(2).Font.Color.RGB = COLOR_WHITE
        tr_b.Paragraphs(2).ParagraphFormat.Alignment = 1
        tr_b.Paragraphs(2).ParagraphFormat.SpaceBefore = 3

        eff_bot = slide.TimeLine.MainSequence.AddEffect(bot, MSO_ANIM_EFFECT_FLOAT, 0, TRIGGER_AFTER_PREV)
        eff_bot.Timing.Duration = 0.4

    # ---------------- SLIDE 15: Experimental Results & Benchmarks ----------------
    def build_slide_15(self):
        print("Building Slide 15: Experimental Results & Benchmarks...")
        slide = self.create_slide(TRANS_PUSH, 0.7)
        self.add_slide_header(slide, "13. Experimental Results & Performance Analysis", "Empirical Verification of Detection Accuracy, Video Throughput & Edge Latency")
        
        metrics = [
            ("95.8%", "BALL TRACKING PRECISION", COLOR_ACCENT_GREEN),
            ("< 42 ms", "END-TO-END LATENCY", COLOR_ACCENT_BLUE),
            ("98.2%", "CARD RECOGNITION RATE", COLOR_ACCENT_GOLD),
            ("58–60 FPS", "STREAMING STABILITY", COLOR_WHITE)
        ]
        
        for i, (val, label, col) in enumerate(metrics):
            box = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40 + i * 225, 115, 205, 80)
            box.Fill.Solid()
            box.Fill.ForeColor.RGB = COLOR_BG_PANEL
            box.Line.ForeColor.RGB = col
            box.Line.Weight = 1.5
            
            tf = box.TextFrame
            tf.MarginLeft = 6
            tf.MarginTop = 8
            tr = tf.TextRange
            tr.Text = f"{val}\r{label}"
            
            tr.Paragraphs(1).Font.Name = "Trebuchet MS"
            tr.Paragraphs(1).Font.Size = 22
            tr.Paragraphs(1).Font.Bold = True
            tr.Paragraphs(1).Font.Color.RGB = col
            tr.Paragraphs(1).ParagraphFormat.Alignment = 2
            
            tr.Paragraphs(2).Font.Name = "Segoe UI"
            tr.Paragraphs(2).Font.Size = 9.5
            tr.Paragraphs(2).Font.Bold = True
            tr.Paragraphs(2).Font.Color.RGB = COLOR_WHITE
            tr.Paragraphs(2).ParagraphFormat.Alignment = 2
            
            eff = slide.TimeLine.MainSequence.AddEffect(box, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
            eff.Timing.Duration = 0.4
            eff.Timing.TriggerDelayTime = 0.15 + (i * 0.1)
            eff.EffectParameters.Direction = DIR_BOTTOM

        p1 = self.add_panel(slide, 40, 215, 430, 275, COLOR_ACCENT_GREEN)
        self.format_panel_content(p1, "DETECTION ACCURACY ACROSS MATCH CONDITIONS", [
            "• Natural Daylight Match Conditions:",
            "  97.2% Ball Precision  |  38 ms Processing Latency  |  0 False Cards",
            "• Stadium Floodlight / Evening Conditions:",
            "  94.8% Ball Precision  |  41 ms Processing Latency  |  98.1% True Cards",
            "• Crowded Penalty Area (4+ Players Occluded):",
            "  93.5% Silhouette Recall  |  44 ms Latency  |  0 False Penalty Triggers",
            "• Out-of-Bounds Boundary Line Crossing:",
            "  99.1% Detection Accuracy verified against ground-truth video footage."
        ], COLOR_ACCENT_GREEN)
            
        p2 = self.add_panel(slide, 490, 215, 430, 275, COLOR_ACCENT_BLUE)
        self.format_panel_content(p2, "SYSTEM THROUGHPUT & LATENCY BENCHMARKS", [
            "• Computer Vision Inference Time: ~38 ms per frame",
            "  (MOG2 Subtraction: 14ms | Contour & Circularity: 12ms | HSV Cards: 8ms | Tracking: 4ms)",
            "• Network Video Latency: ~4 ms via local multipart MJPEG stream",
            "• Total End-to-End Decision Latency: ~42 ms",
            "  Operates well below the 100ms human perceptual threshold for real-time video.",
            "• Sustained 60 FPS Stability: Zero frame drops recorded during continuous 90-minute testing runs."
        ], COLOR_ACCENT_BLUE)

        eff1 = slide.TimeLine.MainSequence.AddEffect(p1, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff1.Timing.Duration = 0.5
        eff1.Timing.TriggerDelayTime = 0.4
        eff1.EffectParameters.Direction = DIR_LEFT
        
        eff2 = slide.TimeLine.MainSequence.AddEffect(p2, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff2.Timing.Duration = 0.5
        eff2.Timing.TriggerDelayTime = 0.5
        eff2.EffectParameters.Direction = DIR_RIGHT

    # ---------------- SLIDE 16: Conclusion & Future Scope ----------------
    def build_slide_16(self):
        print("Building Slide 16: Conclusion & Future Scope...")
        slide = self.create_slide(TRANS_PUSH, 0.7)
        self.add_slide_header(slide, "14. Conclusion & Future Scope", "Project Summary, Practical Takeaways and Roadmap for Advanced Sports Analytics")
        
        p1 = self.add_panel(slide, 40, 115, 430, 375, COLOR_ACCENT_GREEN)
        self.format_panel_content(p1, "PROJECT CONCLUSION & KEY TAKEAWAYS", [
            "• Autonomous Officiating Feasibility:",
            "  Successfully demonstrated that software-defined computer vision algorithms can automate referee decisions without expensive enterprise rigs.",
            "• Elimination of Deliberation Delays:",
            "  Achieved sub-45ms real-time event verification for ball boundaries, yellow/red cards, and box collisions, eliminating lengthy review pauses.",
            "• Unified Tactical & Health Intelligence:",
            "  Blended dynamic spatial 2D heatmaps with simulated IoT biometric telemetry to safeguard athlete health and empower coaches.",
            "• Grassroots Commercial Accessibility:",
            "  Reduces overall infrastructure costs by over 99%, establishing a viable officiating platform for collegiate and academy football."
        ], COLOR_ACCENT_GREEN)
            
        p2 = self.add_panel(slide, 490, 115, 430, 375, COLOR_ACCENT_BLUE)
        self.format_panel_content(p2, "FUTURE SCOPE & ENHANCEMENTS", [
            "• 1. 3D Stereoscopic Offside Line Projection:",
            "  Deploy multi-camera epipolar geometry and deep skeletal pose estimation (tracking shoulders and toes) to render millimeter 3D offside lines.",
            "• 2. Autonomous Drone Aerial Video Feeds:",
            "  Integrate tethered autonomous aerial drones to capture top-down bird's-eye views, tracking team formation compactness and blind-spot fouls.",
            "• 3. Turnkey Grassroots Deployment Kit:",
            "  Package the FastAPI backend and React HUD into a plug-and-play weatherproof field kit with battery-powered IoT smart bands for academies."
        ], COLOR_ACCENT_BLUE)

        eff1 = slide.TimeLine.MainSequence.AddEffect(p1, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff1.Timing.Duration = 0.5
        eff1.EffectParameters.Direction = DIR_LEFT
        
        eff2 = slide.TimeLine.MainSequence.AddEffect(p2, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff2.Timing.Duration = 0.5
        eff2.Timing.TriggerDelayTime = 0.2
        eff2.EffectParameters.Direction = DIR_RIGHT

    # ---------------- SLIDE 17: Dedicated Thank You Slide (Clean, NO Content Below!) ----------------
    def build_slide_17(self):
        print("Building Slide 17: Dedicated Human Academic Thank You Slide (Pure & Clean)...")
        slide = self.create_slide(TRANS_FADE, 1.0)
        
        # Center Atharva College Logo
        logo_path = os.path.join(MEDIA_DIR, "image1.png")
        if os.path.exists(logo_path):
            logo_frame = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 435, 120, 90, 90)
            logo_frame.Fill.Solid()
            logo_frame.Fill.ForeColor.RGB = rgb(255, 255, 255)
            logo_frame.Line.ForeColor.RGB = COLOR_ACCENT_BLUE
            logo_frame.Line.Weight = 1.5
            logo_pic = slide.Shapes.AddPicture(logo_path, False, True, 440, 125, 80, 80)
            
            eff_l = slide.TimeLine.MainSequence.AddEffect(logo_frame, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_ON_CLICK)
            eff_l.Timing.Duration = 0.5
            eff_lp = slide.TimeLine.MainSequence.AddEffect(logo_pic, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_WITH_PREV)
            eff_lp.Timing.Duration = 0.5

        # Dignified Academic "THANK YOU!"
        ty_box = slide.Shapes.AddTextbox(1, 80, 240, 800, 80)
        ty_box.TextFrame.WordWrap = True
        tr_ty = ty_box.TextFrame.TextRange
        tr_ty.Text = "THANK YOU!"
        tr_ty.Font.Name = "Trebuchet MS"
        tr_ty.Font.Size = 60
        tr_ty.Font.Bold = True
        tr_ty.Font.Color.RGB = COLOR_WHITE
        tr_ty.ParagraphFormat.Alignment = 2 # Center
        
        eff_ty = slide.TimeLine.MainSequence.AddEffect(ty_box, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_ty.Timing.Duration = 0.5
        eff_ty.Timing.TriggerDelayTime = 0.2
        eff_ty.EffectParameters.Direction = DIR_TOP
        
        # Subtitle: Questions & Discussion (AND NOTHING BELOW THIS!)
        sub_box = slide.Shapes.AddTextbox(1, 80, 330, 800, 45)
        sub_box.TextFrame.WordWrap = True
        tr_sub = sub_box.TextFrame.TextRange
        tr_sub.Text = "Questions & Discussion"
        tr_sub.Font.Name = "Segoe UI"
        tr_sub.Font.Size = 24
        tr_sub.Font.Bold = True
        tr_sub.Font.Color.RGB = COLOR_ACCENT_BLUE
        tr_sub.ParagraphFormat.Alignment = 2 # Center
        
        eff_sub = slide.TimeLine.MainSequence.AddEffect(sub_box, MSO_ANIM_EFFECT_FLOAT, 0, TRIGGER_WITH_PREV)
        eff_sub.Timing.Duration = 0.5
        eff_sub.Timing.TriggerDelayTime = 0.35

    def generate_all(self, filename="Football_Vision_AI_Presentation.pptx"):
        self.build_slide_1()
        self.build_slide_2()
        self.build_slide_3()
        self.build_slide_4()
        self.build_slide_5()
        self.build_slide_6()
        self.build_slide_7()
        self.build_slide_8()
        self.build_slide_9()
        self.build_slide_10()
        self.build_slide_11()
        self.build_slide_12()
        self.build_slide_13()
        self.build_slide_14()
        self.build_slide_15()
        self.build_slide_16()
        self.build_slide_17()
        
        out_path = os.path.abspath(filename)
        print(f"Saving presentation to: {out_path} ...")
        if os.path.exists(out_path):
            try:
                os.remove(out_path)
                print(f"Removed previous file at {out_path}")
            except Exception as e:
                print(f"Notice: Could not pre-remove {out_path}: {e}")
                
        try:
            self.pres.SaveAs(out_path, 24) # 24 = ppSaveAsOpenXMLPresentation (.pptx)
            print(f"Successfully saved {self.slide_index} slides to {out_path}")
        except Exception as e:
            print(f"Direct SaveAs error: {e}. Trying simple save...")
            self.pres.SaveAs(out_path)
            
        self.pres.Close()
        self.ppt.Quit()
        print("Human-crafted presentation generation completed successfully!")

if __name__ == "__main__":
    generator = HumanDeckGenerator()
    generator.generate_all()
