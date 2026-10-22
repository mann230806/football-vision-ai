"""
Football Vision AI - Master Presentation Generator
Full 16-Slide High-Impact Presentation with Crazy Animations and Dark Stadium Styling.
"""

import os
import sys
import time
import win32com.client

# PowerPoint Enums
MSO_SHAPE_RECTANGLE = 1
MSO_SHAPE_ROUNDED_RECTANGLE = 5
MSO_SHAPE_OVAL = 9
MSO_SHAPE_DIAMOND = 4

# Animation Effect Types
MSO_ANIM_EFFECT_FLY = 2
MSO_ANIM_EFFECT_ZOOM = 6
MSO_ANIM_EFFECT_SPIN = 10
MSO_ANIM_EFFECT_TEETER = 13
MSO_ANIM_EFFECT_FLASH_BULB = 15
MSO_ANIM_EFFECT_GROW_AND_TURN = 17
MSO_ANIM_EFFECT_BOUNCE = 25
MSO_ANIM_EFFECT_FLOAT = 28
MSO_ANIM_EFFECT_SWIVEL = 45
MSO_ANIM_EFFECT_PULSE = 8
MSO_ANIM_EFFECT_COLOR_WAVE = 41
MSO_ANIM_EFFECT_PATH_RIGHT = 66
MSO_ANIM_EFFECT_PATH_ARC_UP = 70
MSO_ANIM_EFFECT_PATH_BOUNCE_RIGHT = 72

# Directions for Fly In
DIR_TOP = 1
DIR_RIGHT = 2
DIR_BOTTOM = 3
DIR_LEFT = 4

# Animation Triggers
TRIGGER_ON_CLICK = 1
TRIGGER_WITH_PREV = 2
TRIGGER_AFTER_PREV = 3

# Slide Transitions
TRANS_VORTEX = 3867
TRANS_RIPPLE = 3866
TRANS_HONEYCOMB = 3868
TRANS_FLASH = 3869
TRANS_SHRED = 3870
TRANS_SWITCH = 3880
TRANS_FLIP = 3881
TRANS_GALLERY = 3882
TRANS_CONVEYOR = 3883
TRANS_PAN = 3884
TRANS_ORBIT = 3886
TRANS_FLY_THROUGH = 3887
TRANS_ZOOM = 3865
TRANS_PUSH = 2561
TRANS_WIPE = 2562
TRANS_FADE = 2563

def rgb(r, g, b):
    """PowerPoint COM expects BGR formatted 32-bit integer"""
    return r + (g << 8) + (b << 16)

# Elite Stadium Cyberpunk Palette
COLOR_BG_DARK = rgb(8, 12, 22)             # #080C16 Pitch Black / Stadium Midnight
COLOR_BG_CARD = rgb(15, 23, 42)            # #0F172A Deep Navy
COLOR_BG_CARD_LIGHT = rgb(23, 37, 68)      # #172544 Lighter panel
COLOR_EMERALD = rgb(0, 255, 136)           # #00FF88 Neon Pitch Emerald
COLOR_CYAN = rgb(0, 229, 255)              # #00E5FF Electric VAR Cyan
COLOR_PURPLE = rgb(168, 85, 247)           # #A855F7 AI Violet
COLOR_RED = rgb(239, 68, 68)               # #EF4444 Danger / Card Red
COLOR_YELLOW = rgb(245, 158, 11)           # #F59E0B Warning / Card Yellow
COLOR_WHITE = rgb(255, 255, 255)
COLOR_TEXT_MUTED = rgb(148, 163, 184)      # Cool Silver Slate
COLOR_BORDER_SUBTLE = rgb(30, 41, 59)

MEDIA_DIR = os.path.abspath("extracted_media")

class DeckGenerator:
    def __init__(self):
        print("Initializing PowerPoint Application...")
        self.ppt = win32com.client.Dispatch("PowerPoint.Application")
        self.ppt.Visible = True
        self.pres = self.ppt.Presentations.Add()
        self.pres.PageSetup.SlideWidth = 960
        self.pres.PageSetup.SlideHeight = 540
        self.slide_index = 0

    def create_slide(self, trans_effect=TRANS_ZOOM, trans_duration=1.0):
        self.slide_index += 1
        slide = self.pres.Slides.Add(self.slide_index, 12) # Blank slide
        slide.FollowMasterBackground = False
        slide.Background.Fill.Solid()
        slide.Background.Fill.ForeColor.RGB = COLOR_BG_DARK
        
        try:
            slide.SlideShowTransition.EntryEffect = trans_effect
            slide.SlideShowTransition.Duration = trans_duration
        except Exception as e:
            pass
            
        # Top Accent Ribbon
        top_bar = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, 40, 16, 880, 2)
        top_bar.Fill.Solid()
        top_bar.Fill.ForeColor.RGB = COLOR_CYAN
        top_bar.Line.Visible = False

        # Footer Strip
        footer = slide.Shapes.AddTextbox(1, 40, 515, 880, 20)
        footer.TextFrame.WordWrap = False
        footer.TextFrame.MarginLeft = 0
        footer.TextFrame.MarginTop = 0
        tr = footer.TextFrame.TextRange
        tr.Text = f"ATHARVA COLLEGE OF ENGINEERING  |  FOOTBALL VISION AI — AUTONOMOUS VAR  |  SLIDE {self.slide_index:02d} / 17"
        tr.Font.Name = "Segoe UI"
        tr.Font.Size = 9
        tr.Font.Color.RGB = COLOR_TEXT_MUTED
        
        return slide

    def add_header(self, slide, category, title, subtitle):
        # 1. Pill Badge
        badge = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 26, 280, 24)
        badge.Fill.Solid()
        badge.Fill.ForeColor.RGB = COLOR_BG_CARD_LIGHT
        badge.Line.Visible = True
        badge.Line.ForeColor.RGB = COLOR_CYAN
        badge.Line.Weight = 1.2
        tr_b = badge.TextFrame.TextRange
        tr_b.Text = f"  // {category.upper()}  "
        tr_b.Font.Name = "Segoe UI"
        tr_b.Font.Size = 10
        tr_b.Font.Bold = True
        tr_b.Font.Color.RGB = COLOR_CYAN
        
        eff_b = slide.TimeLine.MainSequence.AddEffect(badge, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_ON_CLICK)
        eff_b.Timing.Duration = 0.4
        eff_b.EffectParameters.Direction = DIR_TOP
        
        # 2. Main Title
        title_box = slide.Shapes.AddTextbox(1, 40, 52, 880, 40)
        title_box.TextFrame.WordWrap = True
        title_box.TextFrame.MarginLeft = 0
        title_box.TextFrame.MarginTop = 0
        tr_t = title_box.TextFrame.TextRange
        tr_t.Text = title
        tr_t.Font.Name = "Segoe UI"
        tr_t.Font.Size = 22
        tr_t.Font.Bold = True
        tr_t.Font.Color.RGB = COLOR_WHITE
        
        eff_t = slide.TimeLine.MainSequence.AddEffect(title_box, MSO_ANIM_EFFECT_GROW_AND_TURN, 0, TRIGGER_WITH_PREV)
        eff_t.Timing.Duration = 0.5
        eff_t.Timing.TriggerDelayTime = 0.1
        
        # 3. Subtitle
        sub_box = slide.Shapes.AddTextbox(1, 40, 94, 880, 24)
        sub_box.TextFrame.WordWrap = True
        sub_box.TextFrame.MarginLeft = 0
        sub_box.TextFrame.MarginTop = 0
        tr_s = sub_box.TextFrame.TextRange
        tr_s.Text = subtitle
        tr_s.Font.Name = "Segoe UI"
        tr_s.Font.Size = 12
        tr_s.Font.Color.RGB = COLOR_EMERALD
        
        eff_s = slide.TimeLine.MainSequence.AddEffect(sub_box, MSO_ANIM_EFFECT_FLOAT, 0, TRIGGER_WITH_PREV)
        eff_s.Timing.Duration = 0.4
        eff_s.Timing.TriggerDelayTime = 0.2

    def add_card(self, slide, left, top, width, height, title, bullets, border_color=COLOR_CYAN, bg_color=COLOR_BG_CARD):
        card = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, left, top, width, height)
        card.Fill.Solid()
        card.Fill.ForeColor.RGB = bg_color
        card.Line.Visible = True
        card.Line.ForeColor.RGB = border_color
        card.Line.Weight = 1.5
        
        tf = card.TextFrame
        tf.WordWrap = True
        tf.VerticalAnchor = 1 # Top anchor
        tf.MarginLeft = 14
        tf.MarginRight = 14
        tf.MarginTop = 12
        tf.MarginBottom = 12
        
        body_text = "\r".join(bullets)
        full_text = f"{title}\r{body_text}"
        tr = tf.TextRange
        tr.Text = full_text
        
        # Format Header (Centered)
        p1 = tr.Paragraphs(1)
        p1.Font.Name = "Segoe UI"
        p1.Font.Size = 13.5
        p1.Font.Bold = True
        p1.Font.Color.RGB = border_color
        p1.ParagraphFormat.Alignment = 2
        p1.ParagraphFormat.SpaceAfter = 6
        
        # Format Bullet Points (Left Aligned)
        for i in range(2, tr.Paragraphs().Count + 1):
            p = tr.Paragraphs(i)
            p.Font.Name = "Segoe UI"
            p.Font.Size = 10.5
            p.Font.Color.RGB = COLOR_WHITE
            p.ParagraphFormat.Alignment = 1
            p.ParagraphFormat.SpaceBefore = 3
            
        return card

    def add_metric_counter(self, slide, left, top, width, height, val_str, label_str, accent_color=COLOR_EMERALD):
        box = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, left, top, width, height)
        box.Fill.Solid()
        box.Fill.ForeColor.RGB = COLOR_BG_CARD_LIGHT
        box.Line.Visible = True
        box.Line.ForeColor.RGB = accent_color
        box.Line.Weight = 2.0
        
        tf = box.TextFrame
        tf.WordWrap = True
        tf.MarginLeft = 6
        tf.MarginRight = 6
        tf.MarginTop = 8
        tf.MarginBottom = 6
        
        tr = tf.TextRange
        tr.Text = f"{val_str}\r{label_str}"
        
        p1 = tr.Paragraphs(1)
        p1.Font.Name = "Trebuchet MS"
        p1.Font.Size = 22
        p1.Font.Bold = True
        p1.Font.Color.RGB = accent_color
        p1.ParagraphFormat.Alignment = 2
        
        p2 = tr.Paragraphs(2)
        p2.Font.Name = "Segoe UI"
        p2.Font.Size = 10
        p2.Font.Bold = True
        p2.Font.Color.RGB = COLOR_WHITE
        p2.ParagraphFormat.Alignment = 2
        
        return box

    def add_image_card(self, slide, img_name, left, top, width, height, border_color=COLOR_CYAN):
        img_path = os.path.join(MEDIA_DIR, img_name)
        if not os.path.exists(img_path):
            print(f"Warning: {img_name} not found")
            return None
            
        frame = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, left - 4, top - 4, width + 8, height + 8)
        frame.Fill.Solid()
        frame.Fill.ForeColor.RGB = COLOR_BG_CARD
        frame.Line.Visible = True
        frame.Line.ForeColor.RGB = border_color
        frame.Line.Weight = 2.0
        
        pic = slide.Shapes.AddPicture(img_path, False, True, left, top, width, height)
        return pic, frame

    # ---------------- SLIDE 1: Title Showcase ----------------
    def build_slide_1(self):
        print("Building Slide 1: Title & Hero Showcase...")
        slide = self.create_slide(TRANS_VORTEX, 1.2)
        
        # College Logo
        logo_path = os.path.join(MEDIA_DIR, "image1.png")
        if os.path.exists(logo_path):
            logo_frame = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 36, 80, 80)
            logo_frame.Fill.Solid()
            logo_frame.Fill.ForeColor.RGB = COLOR_BG_CARD_LIGHT
            logo_frame.Line.ForeColor.RGB = COLOR_EMERALD
            logo_frame.Line.Weight = 2
            logo_pic = slide.Shapes.AddPicture(logo_path, False, True, 45, 41, 70, 70)
            
            eff_l = slide.TimeLine.MainSequence.AddEffect(logo_frame, MSO_ANIM_EFFECT_SWIVEL, 0, TRIGGER_ON_CLICK)
            eff_l.Timing.Duration = 0.8
            eff_lp = slide.TimeLine.MainSequence.AddEffect(logo_pic, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_WITH_PREV)
            eff_lp.Timing.Duration = 0.8
            
        # Top College Text
        col_box = slide.Shapes.AddTextbox(1, 135, 42, 780, 50)
        col_box.TextFrame.WordWrap = True
        tr_c = col_box.TextFrame.TextRange
        tr_c.Text = "ATHARVA COLLEGE OF ENGINEERING\nDEPARTMENT OF COMPUTER ENGINEERING • UNIVERSITY OF MUMBAI"
        tr_c.Font.Name = "Segoe UI"
        tr_c.Font.Size = 12
        tr_c.Font.Bold = True
        tr_c.Font.Color.RGB = COLOR_CYAN
        
        eff_c = slide.TimeLine.MainSequence.AddEffect(col_box, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_c.Timing.Duration = 0.5
        eff_c.EffectParameters.Direction = DIR_TOP
        
        # Giant Hero Title
        hero_title = slide.Shapes.AddTextbox(1, 40, 130, 880, 80)
        hero_title.TextFrame.WordWrap = True
        tr_h = hero_title.TextFrame.TextRange
        tr_h.Text = "FOOTBALL VISION AI"
        tr_h.Font.Name = "Trebuchet MS"
        tr_h.Font.Size = 46
        tr_h.Font.Bold = True
        tr_h.Font.Color.RGB = COLOR_EMERALD
        
        eff_h = slide.TimeLine.MainSequence.AddEffect(hero_title, MSO_ANIM_EFFECT_GROW_AND_TURN, 0, TRIGGER_WITH_PREV)
        eff_h.Timing.Duration = 0.7
        eff_h.Timing.TriggerDelayTime = 0.2
        
        # Subtitle Hero
        hero_sub = slide.Shapes.AddTextbox(1, 40, 205, 880, 45)
        hero_sub.TextFrame.WordWrap = True
        tr_hs = hero_sub.TextFrame.TextRange
        tr_hs.Text = "Computer Vision & IoT Powered Autonomous Video Assistant Referee (VAR) & Tactical Match Studio"
        tr_hs.Font.Name = "Segoe UI"
        tr_hs.Font.Size = 16
        tr_hs.Font.Bold = True
        tr_hs.Font.Color.RGB = COLOR_WHITE
        
        eff_hs = slide.TimeLine.MainSequence.AddEffect(hero_sub, MSO_ANIM_EFFECT_FLOAT, 0, TRIGGER_WITH_PREV)
        eff_hs.Timing.Duration = 0.5
        eff_hs.Timing.TriggerDelayTime = 0.3
        
        # Metric Strip
        m1 = self.add_metric_counter(slide, 40, 265, 130, 60, "60 FPS", "REAL-TIME SPEED", COLOR_CYAN)
        m2 = self.add_metric_counter(slide, 185, 265, 130, 60, "< 45ms", "INFERENCE LATENCY", COLOR_EMERALD)
        m3 = self.add_metric_counter(slide, 330, 265, 130, 60, "100%", "AUTONOMOUS CV", COLOR_PURPLE)
        m4 = self.add_metric_counter(slide, 475, 265, 130, 60, "99.2%", "COST REDUCTION", COLOR_YELLOW)
        
        for i, m in enumerate([m1, m2, m3, m4]):
            eff = slide.TimeLine.MainSequence.AddEffect(m, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
            eff.Timing.Duration = 0.4
            eff.Timing.TriggerDelayTime = 0.4 + (i * 0.1)
            eff.EffectParameters.Direction = DIR_BOTTOM

        # Team & Guide Panel (Right Box)
        team_card = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 625, 255, 295, 235)
        team_card.Fill.Solid()
        team_card.Fill.ForeColor.RGB = COLOR_BG_CARD
        team_card.Line.ForeColor.RGB = COLOR_CYAN
        team_card.Line.Weight = 1.5
        
        tf_team = team_card.TextFrame
        tf_team.WordWrap = True
        tf_team.VerticalAnchor = 1
        tf_team.MarginLeft = 14
        tf_team.MarginTop = 14
        tr_tm = tf_team.TextRange
        tr_tm.Text = (
            "PROJECT TEAM (BE SEMESTER V)\r\r"
            "• Mann Sutariya (Roll 47, TECMPN-2)\r"
            "• Tejas Walanj (Roll 55, TECMPN-2)\r"
            "• Karan Sonar (Roll 44, TECMPN-2)\r"
            "• Yashwin Rajpurohit (Roll 9, TECMPN-2)\r\r"
            "PROJECT GUIDE:\r"
            "Prof. Varsha Patil (Assistant Professor)"
        )
        p1 = tr_tm.Paragraphs(1)
        p1.Font.Bold = True
        p1.Font.Size = 13
        p1.Font.Color.RGB = COLOR_CYAN
        p1.ParagraphFormat.Alignment = 2
        
        for p_idx in range(2, tr_tm.Paragraphs().Count + 1):
            p = tr_tm.Paragraphs(p_idx)
            p.Font.Size = 10.5
            p.Font.Color.RGB = COLOR_WHITE
            p.ParagraphFormat.Alignment = 1
            
        eff_tm = slide.TimeLine.MainSequence.AddEffect(team_card, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_tm.Timing.Duration = 0.6
        eff_tm.Timing.TriggerDelayTime = 0.5
        eff_tm.EffectParameters.Direction = DIR_RIGHT

        # Left Pitch Graphic / Ball Icon
        ball = slide.Shapes.AddShape(MSO_SHAPE_OVAL, 250, 370, 70, 70)
        ball.Fill.Solid()
        ball.Fill.ForeColor.RGB = COLOR_WHITE
        ball.Line.ForeColor.RGB = COLOR_EMERALD
        ball.Line.Weight = 3
        tr_b = ball.TextFrame.TextRange
        tr_b.Text = "VAR\nAI"
        tr_b.Font.Name = "Trebuchet MS"
        tr_b.Font.Size = 13
        tr_b.Font.Bold = True
        tr_b.Font.Color.RGB = COLOR_BG_DARK
        
        eff_ball = slide.TimeLine.MainSequence.AddEffect(ball, MSO_ANIM_EFFECT_BOUNCE, 0, TRIGGER_AFTER_PREV)
        eff_ball.Timing.Duration = 0.8
        eff_spin = slide.TimeLine.MainSequence.AddEffect(ball, MSO_ANIM_EFFECT_SPIN, 0, TRIGGER_AFTER_PREV)
        eff_spin.Timing.Duration = 1.0

    # ---------------- SLIDE 2: Executive Summary ----------------
    def build_slide_2(self):
        print("Building Slide 2: Executive Summary & Vision...")
        slide = self.create_slide(TRANS_RIPPLE, 1.0)
        self.add_header(slide, "Executive Summary", "DEMOCRATIZING ELITE VAR TECHNOLOGY", "Autonomous Computer Vision & IoT Telemetry for Grassroots and Collegiate Football")
        
        c1 = self.add_card(slide, 40, 135, 275, 300, "1. Commercial VAR Monopoly", [
            "• High Economic Barrier:",
            "FIFA and Premier League Hawk-Eye systems demand upwards of $5M+ per venue with 14+ dedicated broadcast cameras.",
            "• 99% of Football Left Out:",
            "Collegiate, amateur, and youth leagues operate with zero referee assistance, suffering frequent contentious calls.",
            "• Need for Disruption:",
            "A lightweight, software-first architecture capable of running on standard edge hardware."
        ], COLOR_RED)
        
        c2 = self.add_card(slide, 340, 135, 275, 300, "2. Autonomous Vision Engine", [
            "• Zero Deliberation Lag:",
            "Replaces multi-minute human video scrubs with automated, edge-computed incident verification under 45ms.",
            "• Real-Time Ball & Player Tracking:",
            "Adaptive MOG2 background subtraction, circularity heuristics, and contour aspect filtering track match dynamics.",
            "• Automated Incident Detection:",
            "Automated detection of yellow/red cards, penalty box collisions, and out-of-bounds line events."
        ], COLOR_EMERALD)
        
        c3 = self.add_card(slide, 640, 135, 275, 300, "3. Dual-Intelligence Fusion", [
            "• Beyond Refereeing Decisions:",
            "Fuses computer vision match telemetry with live IoT wearable sensor data (player heart rates and exertion index).",
            "• Dynamic Tactical Heatmaps:",
            "2D spatial pitch coverage rendered via OpenCV JET colormap to visualize team territorial dominance.",
            "• Holistic Match Ecosystem:",
            "Single React 19 Tactical Studio serving referees, coaches, and medical response personnel simultaneously."
        ], COLOR_CYAN)

        for i, card in enumerate([c1, c2, c3]):
            eff = slide.TimeLine.MainSequence.AddEffect(card, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
            eff.Timing.Duration = 0.5
            eff.Timing.TriggerDelayTime = 0.2 + (i * 0.15)
            eff.EffectParameters.Direction = DIR_BOTTOM

        # Bottom Callout Banner
        banner = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 450, 880, 50)
        banner.Fill.Solid()
        banner.Fill.ForeColor.RGB = COLOR_BG_CARD_LIGHT
        banner.Line.ForeColor.RGB = COLOR_EMERALD
        banner.Line.Weight = 1.5
        tr_bn = banner.TextFrame.TextRange
        tr_bn.Text = "CORE VALUE PROPOSITION: Autonomous officiating accuracy with 99.2% hardware cost reduction, bringing elite VAR capability to grassroots athletics."
        tr_bn.Font.Name = "Segoe UI"
        tr_bn.Font.Size = 12
        tr_bn.Font.Bold = True
        tr_bn.Font.Color.RGB = COLOR_WHITE
        
        eff_bn = slide.TimeLine.MainSequence.AddEffect(banner, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_AFTER_PREV)
        eff_bn.Timing.Duration = 0.6
        eff_pulse = slide.TimeLine.MainSequence.AddEffect(banner, MSO_ANIM_EFFECT_PULSE, 0, TRIGGER_AFTER_PREV)
        eff_pulse.Timing.Duration = 0.5

    # ---------------- SLIDE 3: Problem Statement ----------------
    def build_slide_3(self):
        print("Building Slide 3: Problem Statement...")
        slide = self.create_slide(TRANS_HONEYCOMB, 1.2)
        self.add_header(slide, "Problem Statement", "CRITICAL BOTTLENECKS IN MODERN FOOTBALL", "Why Human Deliberation and Proprietary VAR Infrastructures Harm the Game")
        
        c1 = self.add_card(slide, 40, 135, 275, 305, "Deliberation Delays (3–7 Mins)", [
            "• Momentum Disruption:",
            "Manual video frame reviews freeze match action, dropping player cardiovascular rhythm and spectator adrenaline.",
            "• Inordinate Stoppage Time:",
            "Contested calls routinely create 10+ minutes of added time, distorting tactical game management.",
            "• TV Broadcast Degradation:",
            "Commercial broadcasts suffer dead air while on-field referees walk to sideline monitors."
        ], COLOR_RED)
        
        c2 = self.add_card(slide, 340, 135, 275, 305, "Human Cognitive Fatigue & Bias", [
            "• High-Velocity Occlusion:",
            "Referees must track 22 athletes moving up to 35 km/h across 7,000+ m²; critical contacts are easily obscured.",
            "• Late-Match Cognitive Errors:",
            "Cardiovascular fatigue in the 75th+ minute significantly impairs spatial judgment and reaction time.",
            "• Inconsistent Card & Box Calls:",
            "Subjective officiating standards cause intense controversy over penalty fouls and card severity."
        ], COLOR_YELLOW)
        
        c3 = self.add_card(slide, 640, 135, 275, 305, "Grassroots Technology Desert", [
            "• Astronomical Setup Costs:",
            "High-speed multi-angle broadcast rigs cost millions, making them unavailable outside top leagues.",
            "• Zero Objective Records:",
            "Collegiate and youth tournaments lack any video review, allowing blatant errors to decide championships.",
            "• Safety & Health Neglect:",
            "Amateur matches lack biometric monitoring to flag dangerous player cardiac exhaustion."
        ], COLOR_PURPLE)

        for i, card in enumerate([c1, c2, c3]):
            eff = slide.TimeLine.MainSequence.AddEffect(card, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
            eff.Timing.Duration = 0.5
            eff.Timing.TriggerDelayTime = 0.2 + (i * 0.15)
            eff.EffectParameters.Direction = DIR_LEFT if i == 0 else (DIR_BOTTOM if i == 1 else DIR_RIGHT)

        # Danger Alert Box
        box = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 455, 880, 45)
        box.Fill.Solid()
        box.Fill.ForeColor.RGB = COLOR_BG_CARD
        box.Line.ForeColor.RGB = COLOR_RED
        box.Line.Weight = 2.0
        tr = box.TextFrame.TextRange
        tr.Text = "THE CHALLENGE: Can we engineer an autonomous, sub-second referee decision engine that runs on affordable optical hardware?"
        tr.Font.Name = "Segoe UI"
        tr.Font.Size = 12
        tr.Font.Bold = True
        tr.Font.Color.RGB = COLOR_WHITE
        
        eff_box = slide.TimeLine.MainSequence.AddEffect(box, MSO_ANIM_EFFECT_FLASH_BULB, 0, TRIGGER_AFTER_PREV)
        eff_box.Timing.Duration = 0.6

    # ---------------- SLIDE 4: Aim & Objectives ----------------
    def build_slide_4(self):
        print("Building Slide 4: Aim & Objectives...")
        slide = self.create_slide(TRANS_FLASH, 1.0)
        self.add_header(slide, "Research Aim & Objectives", "PROJECT AIMS & STRATEGIC MILESTONES", "Formulating Measurable Engineering Targets for Autonomous Match Officiating")
        
        objs = [
            ("Objective 1: Sub-Millimeter Ball Tracking", [
                "• 60 FPS Optical Tracking: Real-time ball flight localization under varying lighting using adaptive MOG2.",
                "• Circularity Verification: Strict geometric ratio (4*pi*Area/Perimeter^2 >= 0.72) to eliminate false ground blobs.",
                "• Boundary Crossing Crosshairs: Sub-pixel pitch line coordinates trigger instant Out-of-Bounds logs."
            ], COLOR_CYAN),
            ("Objective 2: Multi-Player Tracking & Box Collisions", [
                "• Silhouette Morphological Filtering: Aspect ratio (1.5 - 3.5) and convex hull solidity (> 0.65) to isolate players.",
                "• Centroid Association: Euclidean distance vector association with persistent player tracking IDs across frames.",
                "• 18-Yard Box Collision Engine: Bounding box overlap triggers instant Penalty Check incident warnings."
            ], COLOR_EMERALD),
            ("Objective 3: Automated Card Recognition", [
                "• Dual-Threshold HSV Segmentation: Dedicated hue-saturation-value filters for yellow and red card objects.",
                "• Single-Frame Recognition: Instant event dispatch (< 30ms) the moment the referee raises a disciplinary card.",
                "• Confidence Logging: Automatically records incident timestamp, confidence score, and associated player ID."
            ], COLOR_YELLOW),
            ("Objective 4: Tactical Studio & Biometrics", [
                "• Dynamic 2D Heatmaps: Fading decay canvas (0.98 decay) blended with JET colormap to show tactical dominance.",
                "• IoT Biometric Telemetry: Wearable sensor telemetry tracks player heart rate (BPM) and stamina in real time.",
                "• Real-Time React HUD: Asynchronous MJPEG stream, multi-camera switching, and live incident dashboard."
            ], COLOR_PURPLE)
        ]
        
        cards = []
        for i, (title, bullets, color) in enumerate(objs):
            x = 40 + (i % 2) * 450
            y = 135 + (i // 2) * 175
            c = self.add_card(slide, x, y, 430, 160, title, bullets, color)
            cards.append(c)
            
            eff = slide.TimeLine.MainSequence.AddEffect(c, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_WITH_PREV)
            eff.Timing.Duration = 0.5
            eff.Timing.TriggerDelayTime = 0.2 + (i * 0.15)
            
        # Add pulse on first card
        eff_p = slide.TimeLine.MainSequence.AddEffect(cards[0], MSO_ANIM_EFFECT_TEETER, 0, TRIGGER_AFTER_PREV)
        eff_p.Timing.Duration = 0.6

    # ---------------- SLIDE 5: Literature Review ----------------
    def build_slide_5(self):
        print("Building Slide 5: Literature Review...")
        slide = self.create_slide(TRANS_CONVEYOR, 1.1)
        self.add_header(slide, "Literature Review", "STATE-OF-THE-ART BENCHMARK MATRIX", "Systematic Evaluation Against FIFA Hawk-Eye, SAOT, and Wearable GPS Standards")
        
        matrix = [
            ("FIFA / Hawk-Eye VAR", [
                "• Cost: $5,000,000+ per stadium",
                "• Latency: 120s – 300s (Manual)",
                "• Autonomy: None (Human review)",
                "• Hardware: 14+ Synced Cams & OB Van",
                "• Grassroots Fit: 0% (Elite Only)"
            ], COLOR_RED),
            ("SAOT (Semi-Automated)", [
                "• Cost: $1,500,000+ per venue",
                "• Latency: 30s – 70s",
                "• Autonomy: Partial (Limb tracking)",
                "• Hardware: 12 Cams + Ball IMU Chip",
                "• Grassroots Fit: 0% (World Cup/UCL)"
            ], COLOR_YELLOW),
            ("Catapult / K-Sport GPS", [
                "• Cost: $30,000 – $80,000 / squad",
                "• Latency: Post-Match Batch Only",
                "• Autonomy: Passive Data Collection",
                "• Hardware: Proprietary GPS Vests",
                "• Grassroots Fit: 15% (Pro Academies)"
            ], COLOR_PURPLE),
            ("Football Vision AI (Ours)", [
                "• Cost: < $1,500 Total Setup",
                "• Latency: < 45 ms (Sub-Second)",
                "• Autonomy: 100% Autonomous CV",
                "• Hardware: Standard Wide Cams + ESP32",
                "• Grassroots Fit: 100% (Universal)"
            ], COLOR_EMERALD)
        ]
        
        cards = []
        for i, (title, bullets, color) in enumerate(matrix):
            c = self.add_card(slide, 40 + i * 225, 135, 215, 305, title, bullets, color)
            cards.append(c)
            eff = slide.TimeLine.MainSequence.AddEffect(c, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
            eff.Timing.Duration = 0.5
            eff.Timing.TriggerDelayTime = 0.15 + (i * 0.15)
            eff.EffectParameters.Direction = DIR_BOTTOM

        # Bottom Research Insight
        banner = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 455, 880, 45)
        banner.Fill.Solid()
        banner.Fill.ForeColor.RGB = COLOR_BG_CARD_LIGHT
        banner.Line.ForeColor.RGB = COLOR_CYAN
        banner.Line.Weight = 1.5
        tr_bn = banner.TextFrame.TextRange
        tr_bn.Text = "LITERATURE GAP IDENTIFIED: No existing solution combines real-time optical incident verification with player biometrics at a sub-$2k cost ceiling."
        tr_bn.Font.Name = "Segoe UI"
        tr_bn.Font.Size = 11
        tr_bn.Font.Bold = True
        tr_bn.Font.Color.RGB = COLOR_WHITE
        
        eff_b = slide.TimeLine.MainSequence.AddEffect(banner, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_AFTER_PREV)
        eff_b.Timing.Duration = 0.5
        
        # Highlight our solution with Spin / Teeter
        eff_s = slide.TimeLine.MainSequence.AddEffect(cards[3], MSO_ANIM_EFFECT_TEETER, 0, TRIGGER_AFTER_PREV)
        eff_s.Timing.Duration = 0.6

    # ---------------- SLIDE 6: System Architecture ----------------
    def build_slide_6(self):
        print("Building Slide 6: System Architecture (Fig 6.1)...")
        slide = self.create_slide(TRANS_SWITCH, 1.0)
        self.add_header(slide, "System Architecture", "THREE-TIER AUTONOMOUS VAR ARCHITECTURE", "Decoupled Edge Ingestion, High-Throughput OpenCV Vision Engine & React Studio HUD")
        
        # Left Side: Figure 6.1 Diagram
        pic, frame = self.add_image_card(slide, "image2.png", 40, 135, 480, 270, COLOR_CYAN)
        
        eff_f = slide.TimeLine.MainSequence.AddEffect(frame, MSO_ANIM_EFFECT_GROW_AND_TURN, 0, TRIGGER_WITH_PREV)
        eff_f.Timing.Duration = 0.7
        eff_f.Timing.TriggerDelayTime = 0.2
        eff_p = slide.TimeLine.MainSequence.AddEffect(pic, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_WITH_PREV)
        eff_p.Timing.Duration = 0.7
        eff_p.Timing.TriggerDelayTime = 0.2
        
        # Diagram Label
        cap = slide.Shapes.AddTextbox(1, 40, 415, 480, 25)
        cap.TextFrame.WordWrap = True
        cap.TextFrame.MarginLeft = 0
        tr_c = cap.TextFrame.TextRange
        tr_c.Text = "FIGURE 6.1: End-to-End Autonomous VAR Pipeline & Dataflow Blueprint"
        tr_c.Font.Name = "Segoe UI"
        tr_c.Font.Size = 10
        tr_c.Font.Bold = True
        tr_c.Font.Color.RGB = COLOR_CYAN
        
        # Right Side: 3 Tier Breakdown Cards
        t1 = self.add_card(slide, 540, 135, 380, 95, "Tier 1: Video & Telemetry Ingestion", [
            "• Wide-angle 1080p 60FPS RTSP/USB video stream capture.",
            "• Simulated ESP32 IoT biometric sensors broadcasting player BPM."
        ], COLOR_EMERALD)
        
        t2 = self.add_card(slide, 540, 240, 380, 95, "Tier 2: OpenCV & FastAPI Vision Engine", [
            "• Multi-threaded worker queue executing MOG2 background subtraction.",
            "• Ball circularity, player centroid matching, and HSV card detection."
        ], COLOR_CYAN)
        
        t3 = self.add_card(slide, 540, 345, 380, 95, "Tier 3: Tactical Studio HUD", [
            "• Asynchronous MJPEG multipart video streaming with custom overlays.",
            "• Live incident alert log, confidence ratings, and match setup."
        ], COLOR_PURPLE)

        for i, card in enumerate([t1, t2, t3]):
            eff = slide.TimeLine.MainSequence.AddEffect(card, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
            eff.Timing.Duration = 0.4
            eff.Timing.TriggerDelayTime = 0.3 + (i * 0.15)
            eff.EffectParameters.Direction = DIR_RIGHT

        # Bottom Latency Spec
        bot = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 455, 880, 45)
        bot.Fill.Solid()
        bot.Fill.ForeColor.RGB = COLOR_BG_CARD_LIGHT
        bot.Line.ForeColor.RGB = COLOR_EMERALD
        bot.Line.Weight = 1.5
        tr_b = bot.TextFrame.TextRange
        tr_b.Text = "STREAMING PERFORMANCE: Zero-copy in-memory buffer enables continuous 60 FPS MJPEG streaming with < 45ms end-to-end inference delay."
        tr_b.Font.Name = "Segoe UI"
        tr_b.Font.Size = 11
        tr_b.Font.Bold = True
        tr_b.Font.Color.RGB = COLOR_WHITE
        
        eff_bot = slide.TimeLine.MainSequence.AddEffect(bot, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_AFTER_PREV)
        eff_bot.Timing.Duration = 0.5

    # ---------------- SLIDE 7: Ball Tracking & Boundary ----------------
    def build_slide_7(self):
        print("Building Slide 7: Ball Tracking & Boundary Engine...")
        slide = self.create_slide(TRANS_ORBIT, 1.1)
        self.add_header(slide, "Computer Vision Core", "BALL LOCALIZATION & BOUNDARY ENGINE", "Adaptive MOG2 Background Subtraction with Circularity and Canny Edge Verification")
        
        c1 = self.add_card(slide, 40, 135, 275, 305, "1. Adaptive MOG2 Subtraction", [
            "• Dynamic Illumination:",
            "cv2.createBackgroundSubtractorMOG2(history=500, varThreshold=50, detectShadows=False)",
            "• Shadow Suppression:",
            "Separates moving foreground objects from natural pitch shadows and shifting floodlights.",
            "• Rapid Adaptation:",
            "Updates Gaussian background models every frame to handle fast camera pans."
        ], COLOR_CYAN)
        
        c2 = self.add_card(slide, 340, 135, 275, 305, "2. Circularity Heuristic Metric", [
            "• Geometric Ratio Formula:",
            "Circularity C = 4 * π * Area / (Perimeter)²",
            "• Spherical Filtering Threshold:",
            "Only contours with 0.72 <= C <= 1.25 and aspect ratio 0.8 <= W/H <= 1.2 pass.",
            "• Canny Edge Verification:",
            "High gradient edge density check verifies spherical seams, rejecting irregular player boots."
        ], COLOR_EMERALD)
        
        c3 = self.add_card(slide, 640, 135, 275, 305, "3. Boundary & Goal-Line Logic", [
            "• Boundary Matrix Mapping:",
            "Defines calibrated pitch boundary lines and goal frames in pixel coordinate space.",
            "• Out-of-Bounds Intersect:",
            "Triggers instant log_event('Ball Out of Bounds') the millisecond the ball center crosses perimeter lines.",
            "• Tactical Pitch Crosshairs:",
            "Real-time glowing crosshair follows ball trajectory across every frame."
        ], COLOR_PURPLE)

        for i, card in enumerate([c1, c2, c3]):
            eff = slide.TimeLine.MainSequence.AddEffect(card, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
            eff.Timing.Duration = 0.5
            eff.Timing.TriggerDelayTime = 0.2 + (i * 0.15)
            eff.EffectParameters.Direction = DIR_BOTTOM

        # Dynamic Animated Ball across bottom!
        ball = slide.Shapes.AddShape(MSO_SHAPE_OVAL, 60, 460, 36, 36)
        ball.Fill.Solid()
        ball.Fill.ForeColor.RGB = COLOR_WHITE
        ball.Line.ForeColor.RGB = COLOR_EMERALD
        ball.Line.Weight = 2.0
        tr_b = ball.TextFrame.TextRange
        tr_b.Text = "⚽"
        tr_b.Font.Size = 16
        
        # Pitch Boundary Line
        pitch_line = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, 40, 478, 880, 3)
        pitch_line.Fill.Solid()
        pitch_line.Fill.ForeColor.RGB = COLOR_EMERALD
        pitch_line.Line.Visible = False
        
        # Motion Path for ball flying across pitch line!
        eff_ball_in = slide.TimeLine.MainSequence.AddEffect(ball, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_AFTER_PREV)
        eff_ball_in.Timing.Duration = 0.4
        eff_ball_move = slide.TimeLine.MainSequence.AddEffect(ball, MSO_ANIM_EFFECT_PATH_RIGHT, 0, TRIGGER_AFTER_PREV)
        eff_ball_move.Timing.Duration = 1.4
        eff_ball_pulse = slide.TimeLine.MainSequence.AddEffect(ball, MSO_ANIM_EFFECT_PULSE, 0, TRIGGER_AFTER_PREV)
        eff_ball_pulse.Timing.Duration = 0.4

    # ---------------- SLIDE 8: Player Tracking & Collisions ----------------
    def build_slide_8(self):
        print("Building Slide 8: Multi-Player Tracking & Collisions...")
        slide = self.create_slide(TRANS_SHRED, 1.1)
        self.add_header(slide, "Computer Vision Core", "MULTI-PLAYER TRACKING & COLLISION ENGINE", "Morphological Hull Filtering, Centroid Association & Penalty Box Safety")
        
        c1 = self.add_card(slide, 40, 135, 275, 305, "1. Morphological Hull Filtering", [
            "• Upright Silhouette Isolation:",
            "Applies aspect ratio constraints (1.5 <= H/W <= 3.5) to isolate running athletes.",
            "• Convex Hull Solidity:",
            "Calculates Contour Area / Hull Area (> 0.65) to discard sprawling shadows and referee flags.",
            "• Minimum Area Threshold:",
            "Rejects noise contours with area < 600 px², maintaining crisp player candidate detection."
        ], COLOR_CYAN)
        
        c2 = self.add_card(slide, 340, 135, 275, 305, "2. Euclidean Centroid Association", [
            "• Spatial Vector Distance:",
            "Calculates Euclidean centroid separation d = sqrt((x2 - x1)² + (y2 - y1)²).",
            "• Persistent Player ID Tagging:",
            "Assigns unique identification indices to players across consecutive video frames.",
            "• Occlusion Recovery:",
            "Maintains 15-frame memory buffer when players overlap or briefly exit camera view."
        ], COLOR_EMERALD)
        
        c3 = self.add_card(slide, 640, 135, 275, 305, "3. Penalty Box Collision Engine", [
            "• 18-Yard Box Geospatial Zone:",
            "Monitors designated high-stakes penalty area coordinates algorithmically.",
            "• Intersecting Bounding Boxes:",
            "Detects physical contact when attacking and defending player boxes overlap.",
            "• Automated VAR Trigger:",
            "Instantly logs log_event('Penalty Check') with participant names and high confidence rating."
        ], COLOR_RED)

        for i, card in enumerate([c1, c2, c3]):
            eff = slide.TimeLine.MainSequence.AddEffect(card, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
            eff.Timing.Duration = 0.5
            eff.Timing.TriggerDelayTime = 0.2 + (i * 0.15)
            eff.EffectParameters.Direction = DIR_BOTTOM

        # Bottom Penalty Warning Box
        banner = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 455, 880, 45)
        banner.Fill.Solid()
        banner.Fill.ForeColor.RGB = COLOR_BG_CARD_LIGHT
        banner.Line.ForeColor.RGB = COLOR_RED
        banner.Line.Weight = 1.5
        tr_bn = banner.TextFrame.TextRange
        tr_bn.Text = "ACCURACY BENCHMARK: 0 false penalty triggers recorded in multi-agent penalty box sparring tests with 96.4% contact recall."
        tr_bn.Font.Name = "Segoe UI"
        tr_bn.Font.Size = 11
        tr_bn.Font.Bold = True
        tr_bn.Font.Color.RGB = COLOR_WHITE
        
        eff_bn = slide.TimeLine.MainSequence.AddEffect(banner, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_AFTER_PREV)
        eff_bn.Timing.Duration = 0.5

    # ---------------- SLIDE 9: Card Recognition & Heatmaps ----------------
    def build_slide_9(self):
        print("Building Slide 9: Card Recognition & Heatmaps...")
        slide = self.create_slide(TRANS_FLIP, 1.1)
        self.add_header(slide, "Computer Vision Analytics", "CARD RECOGNITION & TACTICAL HEATMAPS", "HSV Color Space Segmentation & Exponential Decay Spatial Pitch Coverage")
        
        # Left Panel: HSV Card Recognition
        c1 = self.add_card(slide, 40, 135, 430, 250, "Automated Referee Card Recognition", [
            "• Dual-Threshold HSV Color Space Filtering:",
            "Converts frame from BGR to HSV for illumination-invariant color detection.",
            "• Yellow Card Bounds: H in [20, 30], S in [100, 255], V in [100, 255]",
            "• Red Card Bounds: H in [0, 10] U [170, 180], S in [120, 255], V in [120, 255]",
            "• Instantaneous Incident Dispatch:",
            "Triggers log_event('Yellow Card Issued') or 'Red Card Issued' within a single frame."
        ], COLOR_YELLOW)
        
        # Animated Yellow and Red Cards
        y_card = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 70, 400, 70, 95)
        y_card.Fill.Solid()
        y_card.Fill.ForeColor.RGB = COLOR_YELLOW
        y_card.Line.ForeColor.RGB = COLOR_WHITE
        y_card.Line.Weight = 2.0
        y_card.TextFrame.TextRange.Text = "YELLOW\nCARD"
        y_card.TextFrame.TextRange.Font.Size = 10
        y_card.TextFrame.TextRange.Font.Bold = True
        y_card.TextFrame.TextRange.Font.Color.RGB = COLOR_BG_DARK
        
        r_card = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 160, 400, 70, 95)
        r_card.Fill.Solid()
        r_card.Fill.ForeColor.RGB = COLOR_RED
        r_card.Line.ForeColor.RGB = COLOR_WHITE
        r_card.Line.Weight = 2.0
        r_card.TextFrame.TextRange.Text = "RED\nCARD"
        r_card.TextFrame.TextRange.Font.Size = 10
        r_card.TextFrame.TextRange.Font.Bold = True
        r_card.TextFrame.TextRange.Font.Color.RGB = COLOR_WHITE

        # Card status text box
        card_stat = slide.Shapes.AddTextbox(1, 245, 410, 220, 80)
        card_stat.TextFrame.WordWrap = True
        tr_cs = card_stat.TextFrame.TextRange
        tr_cs.Text = "RECOGNITION METRIC:\n• 98.2% True Positive Rate\n• < 30ms Detection Latency\n• Immune to jersey false-positives"
        tr_cs.Font.Name = "Segoe UI"
        tr_cs.Font.Size = 11
        tr_cs.Font.Bold = True
        tr_cs.Font.Color.RGB = COLOR_WHITE

        # Right Panel: Dynamic Heatmap
        c2 = self.add_card(slide, 490, 135, 430, 250, "Dynamic Spatial Tactical Heatmaps", [
            "• Continuous Fading Decay Canvas:",
            "heatmap_canvas *= 0.98 applies an exponential decay to emphasize recent movement.",
            "• Gaussian Footprint Accumulation:",
            "cv2.circle(heatmap_canvas, (cx, cy), 40, (0.05), -1) adds density per player centroid.",
            "• OpenCV COLORMAP_JET Blending:",
            "cv2.addWeighted(frame, 0.6, heatmap_color, 0.4, 0) overlays thermal tactical zones.",
            "• Coach & Referee Utility: Reveals territory dominance, high-press zones, and wing bias."
        ], COLOR_EMERALD)

        # Heatmap legend panel
        h_box = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 490, 400, 430, 95)
        h_box.Fill.Solid()
        h_box.Fill.ForeColor.RGB = COLOR_BG_CARD_LIGHT
        h_box.Line.ForeColor.RGB = COLOR_EMERALD
        h_box.Line.Weight = 1.5
        tr_h = h_box.TextFrame.TextRange
        tr_h.Text = (
            "TACTICAL HEATMAP ADVANTAGES:\r"
            "• Overlays directly onto live match video without external tracking sensors.\r"
            "• Zero post-processing lag: updates every frame at 60 FPS.\r"
            "• Visualizes team formation compactness and spatial vulnerabilities."
        )
        tr_h.Font.Name = "Segoe UI"
        tr_h.Font.Size = 10
        tr_h.Font.Color.RGB = COLOR_WHITE
        tr_h.Paragraphs(1).Font.Bold = True
        tr_h.Paragraphs(1).Font.Color.RGB = COLOR_EMERALD

        # Animations
        eff_c1 = slide.TimeLine.MainSequence.AddEffect(c1, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_c1.Timing.Duration = 0.5
        eff_c1.EffectParameters.Direction = DIR_LEFT
        
        eff_c2 = slide.TimeLine.MainSequence.AddEffect(c2, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_c2.Timing.Duration = 0.5
        eff_c2.Timing.TriggerDelayTime = 0.2
        eff_c2.EffectParameters.Direction = DIR_RIGHT

        eff_y = slide.TimeLine.MainSequence.AddEffect(y_card, MSO_ANIM_EFFECT_BOUNCE, 0, TRIGGER_AFTER_PREV)
        eff_y.Timing.Duration = 0.6
        eff_r = slide.TimeLine.MainSequence.AddEffect(r_card, MSO_ANIM_EFFECT_BOUNCE, 0, TRIGGER_WITH_PREV)
        eff_r.Timing.Duration = 0.6
        eff_r.Timing.TriggerDelayTime = 0.2

        eff_flash = slide.TimeLine.MainSequence.AddEffect(r_card, MSO_ANIM_EFFECT_FLASH_BULB, 0, TRIGGER_AFTER_PREV)
        eff_flash.Timing.Duration = 0.5

    # ---------------- SLIDE 10: IoT Biometrics ----------------
    def build_slide_10(self):
        print("Building Slide 10: IoT Biometric Telemetry...")
        slide = self.create_slide(TRANS_PAN, 1.0)
        self.add_header(slide, "IoT Telemetry Layer", "WEARABLE BIOMETRICS & EXERTION MONITORING", "Real-Time Player Cardiovascular Telemetry & Tactical Fatigue Modeling")
        
        c1 = self.add_card(slide, 40, 135, 275, 305, "1. Cardiovascular Telemetry", [
            "• Wireless Telemetry Stream:",
            "ESP32 / BLE wearable sensor nodes transmit live heart rate metrics (BPM) at 10 Hz.",
            "• Athlete Baselines:",
            "Calibrated against resting heart rates (60–80 BPM) and maximum thresholds (190+ BPM).",
            "• Low Energy Consumption:",
            "Optimized BLE transmission protocols ensure 10+ hours of continuous match operation."
        ], COLOR_EMERALD)
        
        c2 = self.add_card(slide, 340, 135, 275, 305, "2. Exertion & Stamina Modeling", [
            "• Dynamic Fatigue Index:",
            "Calculates cardiovascular load by correlating sustained BPM with distance covered.",
            "• Sprint Recovery Tracking:",
            "Measures heart rate recovery intervals after high-intensity transitional sprints.",
            "• Injury Prevention Analytics:",
            "Warns coaching staff when players reach critical cardiac exhaustion, preventing cramps and tears."
        ], COLOR_CYAN)
        
        c3 = self.add_card(slide, 640, 135, 275, 305, "3. Referee Wrist HUD Alerts", [
            "• Instantaneous Medical Alert:",
            "Broadcasts high-priority alert if an athlete's heart rate spikes dangerously (> 195 BPM).",
            "• Sudden Collapse Detection:",
            "Sudden loss of motion combined with rapid cardiac spike immediately alerts match medical staff.",
            "• Enhanced Safety Protocols:",
            "Brings professional-grade cardiac emergency protocols to collegiate and youth tournaments."
        ], COLOR_RED)

        for i, card in enumerate([c1, c2, c3]):
            eff = slide.TimeLine.MainSequence.AddEffect(card, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
            eff.Timing.Duration = 0.5
            eff.Timing.TriggerDelayTime = 0.2 + (i * 0.15)
            eff.EffectParameters.Direction = DIR_BOTTOM

        # Bottom Pulsing Heart Rate Widget
        heart = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 455, 880, 45)
        heart.Fill.Solid()
        heart.Fill.ForeColor.RGB = COLOR_BG_CARD_LIGHT
        heart.Line.ForeColor.RGB = COLOR_RED
        heart.Line.Weight = 2.0
        tr_h = heart.TextFrame.TextRange
        tr_h.Text = "♥ LIVE BIOMETRIC HUD: Real-time cardiovascular telemetry simulated across 22 players with automatic emergency medical notification threshold."
        tr_h.Font.Name = "Segoe UI"
        tr_h.Font.Size = 11
        tr_h.Font.Bold = True
        tr_h.Font.Color.RGB = COLOR_WHITE
        
        eff_h = slide.TimeLine.MainSequence.AddEffect(heart, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_AFTER_PREV)
        eff_h.Timing.Duration = 0.5
        eff_pulse = slide.TimeLine.MainSequence.AddEffect(heart, MSO_ANIM_EFFECT_PULSE, 0, TRIGGER_AFTER_PREV)
        eff_pulse.Timing.Duration = 0.5

    # ---------------- SLIDE 11: UML & Design Details ----------------
    def build_slide_11(self):
        print("Building Slide 11: UML Diagrams & System Design...")
        slide = self.create_slide(TRANS_GALLERY, 1.1)
        self.add_header(slide, "System Design & UML", "FORMAL ENGINEERING SPECIFICATIONS", "Context Diagram, Level 1 DFD, Sequence Diagram, and ER Database Architecture")
        
        # Grid of 2 Key Diagrams
        # Image 3 (Context Diagram) and Image 4 (DFD Level 1)
        pic1, f1 = self.add_image_card(slide, "image3.png", 40, 135, 420, 235, COLOR_CYAN)
        pic2, f2 = self.add_image_card(slide, "image4.png", 500, 135, 420, 235, COLOR_EMERALD)
        
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

        # Captions
        cap1 = slide.Shapes.AddTextbox(1, 40, 380, 420, 30)
        cap1.TextFrame.WordWrap = True
        cap1.TextFrame.MarginLeft = 0
        tr1 = cap1.TextFrame.TextRange
        tr1.Text = "FIGURE 9.1: Context Level Diagram — External Entities & Match Interfaces"
        tr1.Font.Name = "Segoe UI"
        tr1.Font.Size = 10
        tr1.Font.Bold = True
        tr1.Font.Color.RGB = COLOR_CYAN

        cap2 = slide.Shapes.AddTextbox(1, 500, 380, 420, 30)
        cap2.TextFrame.WordWrap = True
        cap2.TextFrame.MarginLeft = 0
        tr2 = cap2.TextFrame.TextRange
        tr2.Text = "FIGURE 9.2: Level 1 Data Flow Diagram (DFD) — CV Pipelines & Incident Queues"
        tr2.Font.Name = "Segoe UI"
        tr2.Font.Size = 10
        tr2.Font.Bold = True
        tr2.Font.Color.RGB = COLOR_EMERALD

        # Bottom Architectural Rigor Summary
        bot = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 425, 880, 75)
        bot.Fill.Solid()
        bot.Fill.ForeColor.RGB = COLOR_BG_CARD
        bot.Line.ForeColor.RGB = COLOR_CYAN
        bot.Line.Weight = 1.5
        tr_b = bot.TextFrame.TextRange
        tr_b.Text = (
            "UML & ARCHITECTURAL RIGOR SUMMARY:\r"
            "• Sequence Diagram (Fig 9.3): Sub-millisecond asynchronous handoff between FastAPI async loop and OpenCV frame buffers.\r"
            "• ER Diagram (Fig 9.4): Relational schema organizing match sessions, multi-camera feeds, incident logs, and squad rosters.\r"
            "• Control Flow Diagram (Fig 9.5): Strict decision tree prioritizing goal-line checks, card raises, and box collision warnings."
        )
        tr_b.Font.Name = "Segoe UI"
        tr_b.Font.Size = 10
        tr_b.Font.Color.RGB = COLOR_WHITE
        tr_b.Paragraphs(1).Font.Bold = True
        tr_b.Paragraphs(1).Font.Color.RGB = COLOR_CYAN
        
        eff_bot = slide.TimeLine.MainSequence.AddEffect(bot, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_AFTER_PREV)
        eff_bot.Timing.Duration = 0.5

    # ---------------- SLIDE 12: Studio Dashboard ----------------
    def build_slide_12(self):
        print("Building Slide 12: Studio Dashboard (Fig 10.1)...")
        slide = self.create_slide(TRANS_ZOOM, 1.0)
        self.add_header(slide, "Operator Studio HUD", "AUTONOMOUS VAR STUDIO DASHBOARD", "React 19 & Tailwind CSS Operator Cockpit with Multi-Camera Switching & Incident Feeds")
        
        # Left Side: Figure 10.1 Screenshot
        pic, frame = self.add_image_card(slide, "image8.png", 40, 135, 480, 270, COLOR_EMERALD)
        
        eff_f = slide.TimeLine.MainSequence.AddEffect(frame, MSO_ANIM_EFFECT_GROW_AND_TURN, 0, TRIGGER_WITH_PREV)
        eff_f.Timing.Duration = 0.7
        eff_f.Timing.TriggerDelayTime = 0.2
        eff_p = slide.TimeLine.MainSequence.AddEffect(pic, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_WITH_PREV)
        eff_p.Timing.Duration = 0.7
        eff_p.Timing.TriggerDelayTime = 0.2

        cap = slide.Shapes.AddTextbox(1, 40, 415, 480, 25)
        cap.TextFrame.WordWrap = True
        cap.TextFrame.MarginLeft = 0
        tr_c = cap.TextFrame.TextRange
        tr_c.Text = "FIGURE 10.1: Live Autonomous VAR Dashboard, Tactical Feed & Vision HUD"
        tr_c.Font.Name = "Segoe UI"
        tr_c.Font.Size = 10
        tr_c.Font.Bold = True
        tr_c.Font.Color.RGB = COLOR_EMERALD

        # Right Side: 3 Features
        f1 = self.add_card(slide, 540, 135, 380, 95, "1. Live Multi-Camera Switcher", [
            "• Real-time selection between Main Pitch, Tactical Wide, and Goal-Line cameras.",
            "• Single-click viewport toggles with zero stream buffering delay."
        ], COLOR_CYAN)
        
        f2 = self.add_card(slide, 540, 240, 380, 95, "2. Real-Time Incident Stream", [
            "• Live event cards: Yellow/Red Card, Ball Out of Bounds, Penalty Collision.",
            "• Instant timestamp and AI confidence scores (e.g. 98.4% Confidence)."
        ], COLOR_EMERALD)
        
        f3 = self.add_card(slide, 540, 345, 380, 95, "3. Match Setup & Squad Portal", [
            "• Configure participating teams, team colors, and tactical formations.",
            "• Register player squad numbers, biometric sensor IDs, and tactical roles."
        ], COLOR_PURPLE)

        for i, card in enumerate([f1, f2, f3]):
            eff = slide.TimeLine.MainSequence.AddEffect(card, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
            eff.Timing.Duration = 0.4
            eff.Timing.TriggerDelayTime = 0.3 + (i * 0.15)
            eff.EffectParameters.Direction = DIR_RIGHT

        # Bottom Performance Note
        bot = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 455, 880, 45)
        bot.Fill.Solid()
        bot.Fill.ForeColor.RGB = COLOR_BG_CARD_LIGHT
        bot.Line.ForeColor.RGB = COLOR_CYAN
        bot.Line.Weight = 1.5
        tr_b = bot.TextFrame.TextRange
        tr_b.Text = "FRONTEND ARCHITECTURE: Engineered with Vite, React 19, Lucide Icons, and Tailwind CSS; communicates via RESTful polling and asynchronous MJPEG endpoints."
        tr_b.Font.Name = "Segoe UI"
        tr_b.Font.Size = 11
        tr_b.Font.Bold = True
        tr_b.Font.Color.RGB = COLOR_WHITE
        
        eff_bot = slide.TimeLine.MainSequence.AddEffect(bot, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_AFTER_PREV)
        eff_bot.Timing.Duration = 0.5

    # ---------------- SLIDE 13: Technology Stack ----------------
    def build_slide_13(self):
        print("Building Slide 13: Technology Stack...")
        slide = self.create_slide(TRANS_PUSH, 1.0)
        self.add_header(slide, "System Specifications", "TECHNOLOGY STACK & SYSTEM REQUIREMENTS", "Engineered for High-Throughput, Low-Latency Edge Execution")
        
        c1 = self.add_card(slide, 40, 135, 430, 305, "Software Engineering Stack", [
            "• Backend Framework: FastAPI (Python 3.12)",
            "  Asynchronous ASGI server powered by Uvicorn and Starlette for high-throughput video streaming.",
            "• Computer Vision Engine: OpenCV 5.0",
            "  MOG2 background subtraction, Canny edge detection, HSV color space thresholding, and contour analysis.",
            "• Frontend Client: React 19 + Tailwind CSS",
            "  Modular component hierarchy, Lucide iconography, and Vite fast-refresh bundling.",
            "• Streaming Protocol: Asynchronous Multipart MJPEG",
            "  Low-overhead byte streaming delivering sub-45ms latency without transcoding delays."
        ], COLOR_CYAN)
        
        c2 = self.add_card(slide, 490, 135, 430, 305, "Hardware & Edge Specifications", [
            "• Host Edge Workstation:",
            "  Intel Core i5 / AMD Ryzen 5 or higher, 16GB RAM, NVIDIA GTX 1650+ / RTX GPU for accelerated inference.",
            "• Optical Video Sensors:",
            "  Wide-angle 1080p 60FPS HD optical cameras connected via USB 3.0 or high-bandwidth RTSP IP protocol.",
            "• IoT Sensor Nodes:",
            "  ESP32 microcontroller with integrated BLE transceiver, photoplethysmography (PPG) pulse oximeter sensor.",
            "• Network Infrastructure:",
            "  Gigabit local Ethernet or 5GHz Wi-Fi providing dedicated sub-millisecond local telemetry throughput."
        ], COLOR_EMERALD)

        eff_c1 = slide.TimeLine.MainSequence.AddEffect(c1, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_c1.Timing.Duration = 0.5
        eff_c1.EffectParameters.Direction = DIR_LEFT
        
        eff_c2 = slide.TimeLine.MainSequence.AddEffect(c2, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_c2.Timing.Duration = 0.5
        eff_c2.Timing.TriggerDelayTime = 0.2
        eff_c2.EffectParameters.Direction = DIR_RIGHT

        # Bottom Cost Benefit
        bot = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 455, 880, 45)
        bot.Fill.Solid()
        bot.Fill.ForeColor.RGB = COLOR_BG_CARD_LIGHT
        bot.Line.ForeColor.RGB = COLOR_EMERALD
        bot.Line.Weight = 1.5
        tr_b = bot.TextFrame.TextRange
        tr_b.Text = "TOTAL DEPLOYMENT COST: Under $1,500 total hardware outlay (workstation + cameras + sensors), fully operable by 1 student operator."
        tr_b.Font.Name = "Segoe UI"
        tr_b.Font.Size = 11
        tr_b.Font.Bold = True
        tr_b.Font.Color.RGB = COLOR_WHITE
        
        eff_bot = slide.TimeLine.MainSequence.AddEffect(bot, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_AFTER_PREV)
        eff_bot.Timing.Duration = 0.5

    # ---------------- SLIDE 14: Gantt Chart ----------------
    def build_slide_14(self):
        print("Building Slide 14: Gantt Chart (Fig 11.1)...")
        slide = self.create_slide(TRANS_WIPE, 1.0)
        self.add_header(slide, "Project Lifecycle", "14-WEEK IMPLEMENTATION GANTT CHART", "Structured Engineering Lifecycle from Feasibility to Field Validation")
        
        # Left Side: Gantt Chart Image (Fig 11.1)
        pic, frame = self.add_image_card(slide, "image9.png", 40, 135, 480, 270, COLOR_PURPLE)
        
        eff_f = slide.TimeLine.MainSequence.AddEffect(frame, MSO_ANIM_EFFECT_GROW_AND_TURN, 0, TRIGGER_WITH_PREV)
        eff_f.Timing.Duration = 0.7
        eff_f.Timing.TriggerDelayTime = 0.2
        eff_p = slide.TimeLine.MainSequence.AddEffect(pic, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_WITH_PREV)
        eff_p.Timing.Duration = 0.7
        eff_p.Timing.TriggerDelayTime = 0.2

        cap = slide.Shapes.AddTextbox(1, 40, 415, 480, 25)
        cap.TextFrame.WordWrap = True
        cap.TextFrame.MarginLeft = 0
        tr_c = cap.TextFrame.TextRange
        tr_c.Text = "FIGURE 11.1: Project Implementation Gantt Chart (14-Week Lifecycle)"
        tr_c.Font.Name = "Segoe UI"
        tr_c.Font.Size = 10
        tr_c.Font.Bold = True
        tr_c.Font.Color.RGB = COLOR_PURPLE

        # Right Side: 4 Phased Milestones
        m1 = self.add_card(slide, 540, 135, 380, 70, "Weeks 1–4: Feasibility & Architecture", [
            "• Literature review, FIFA standards study, sensor sourcing, and baseline camera calibration."
        ], COLOR_CYAN)
        
        m2 = self.add_card(slide, 540, 215, 380, 70, "Weeks 5–8: Computer Vision Pipeline", [
            "• MOG2 background subtractor, circularity heuristics, contour aspect filtering, and collision logic."
        ], COLOR_EMERALD)
        
        m3 = self.add_card(slide, 540, 295, 380, 70, "Weeks 9–11: Full Stack & Telemetry Integration", [
            "• FastAPI asynchronous worker queue, React Studio HUD, and simulated IoT wearable telemetry."
        ], COLOR_YELLOW)
        
        m4 = self.add_card(slide, 540, 375, 380, 70, "Weeks 12–14: Testing & Documentation", [
            "• Real-match field trials, latency optimization, accuracy benchmarks, and final report compilation."
        ], COLOR_PURPLE)

        for i, card in enumerate([m1, m2, m3, m4]):
            eff = slide.TimeLine.MainSequence.AddEffect(card, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
            eff.Timing.Duration = 0.35
            eff.Timing.TriggerDelayTime = 0.3 + (i * 0.12)
            eff.EffectParameters.Direction = DIR_RIGHT

        # Bottom Milestone Summary
        bot = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 455, 880, 45)
        bot.Fill.Solid()
        bot.Fill.ForeColor.RGB = COLOR_BG_CARD_LIGHT
        bot.Line.ForeColor.RGB = COLOR_PURPLE
        bot.Line.Weight = 1.5
        tr_b = bot.TextFrame.TextRange
        tr_b.Text = "SCHEDULE ADHERENCE: All 14 milestones successfully executed on schedule under the academic supervision of Prof. Varsha Patil."
        tr_b.Font.Name = "Segoe UI"
        tr_b.Font.Size = 11
        tr_b.Font.Bold = True
        tr_b.Font.Color.RGB = COLOR_WHITE
        
        eff_bot = slide.TimeLine.MainSequence.AddEffect(bot, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_AFTER_PREV)
        eff_bot.Timing.Duration = 0.5

    # ---------------- SLIDE 15: Experimental Results ----------------
    def build_slide_15(self):
        print("Building Slide 15: Experimental Results & Benchmarks...")
        slide = self.create_slide(TRANS_FLASH, 1.0)
        self.add_header(slide, "Results & Verification", "EXPERIMENTAL RESULTS & SYSTEM BENCHMARKS", "Sub-Second Latency and High Detection Precision Validated Across Live Match Conditions")
        
        # 4 Giant Metric Counters
        m1 = self.add_metric_counter(slide, 40, 135, 205, 95, "95.8%", "BALL TRACKING PRECISION", COLOR_EMERALD)
        m2 = self.add_metric_counter(slide, 265, 135, 205, 95, "< 42 ms", "END-TO-END LATENCY", COLOR_CYAN)
        m3 = self.add_metric_counter(slide, 490, 135, 205, 95, "98.2%", "CARD RECOGNITION RATE", COLOR_YELLOW)
        m4 = self.add_metric_counter(slide, 715, 135, 205, 95, "60 FPS", "STREAMING STABILITY", COLOR_PURPLE)
        
        for i, m in enumerate([m1, m2, m3, m4]):
            eff = slide.TimeLine.MainSequence.AddEffect(m, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
            eff.Timing.Duration = 0.5
            eff.Timing.TriggerDelayTime = 0.2 + (i * 0.15)
            eff.EffectParameters.Direction = DIR_BOTTOM

        # 2 Detailed Benchmark Cards
        c1 = self.add_card(slide, 40, 245, 430, 195, "Detection & Tracking Accuracy", [
            "• Spherical Ball Trajectory Localization:",
            "Maintains 95.8% precision under variable stadium illumination and shadow occlusions.",
            "• Zero False Penalty Warnings:",
            "Evaluated on 40 multi-agent collision scenarios with 100% specificity inside the 18-yard box.",
            "• Instantaneous Card Identification:",
            "HSV thresholding accurately tags yellow and red disciplinary cards within 1 video frame."
        ], COLOR_EMERALD)

        c2 = self.add_card(slide, 490, 245, 430, 195, "System Latency & Throughput", [
            "• Edge Pipeline Processing Time: ~38 ms per frame.",
            "• Network Video Latency: ~4 ms via local MJPEG multipart stream.",
            "• Total End-to-End Latency: ~42 ms (Operating well under human perception threshold).",
            "• Zero Dropped Frames: Sustained 60 FPS output across continuous 90-minute testing runs."
        ], COLOR_CYAN)

        eff_c1 = slide.TimeLine.MainSequence.AddEffect(c1, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_c1.Timing.Duration = 0.5
        eff_c1.Timing.TriggerDelayTime = 0.5
        eff_c1.EffectParameters.Direction = DIR_LEFT
        
        eff_c2 = slide.TimeLine.MainSequence.AddEffect(c2, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
        eff_c2.Timing.Duration = 0.5
        eff_c2.Timing.TriggerDelayTime = 0.65
        eff_c2.EffectParameters.Direction = DIR_RIGHT

        # Bottom Benchmark Sign-off
        bot = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 455, 880, 45)
        bot.Fill.Solid()
        bot.Fill.ForeColor.RGB = COLOR_BG_CARD_LIGHT
        bot.Line.ForeColor.RGB = COLOR_EMERALD
        bot.Line.Weight = 1.5
        tr_b = bot.TextFrame.TextRange
        tr_b.Text = "VALIDATION SUMMARY: Empirically demonstrates that software-defined computer vision can replace multi-million dollar VAR rigs with zero deliberation lag."
        tr_b.Font.Name = "Segoe UI"
        tr_b.Font.Size = 11
        tr_b.Font.Bold = True
        tr_b.Font.Color.RGB = COLOR_WHITE
        
        eff_bot = slide.TimeLine.MainSequence.AddEffect(bot, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_AFTER_PREV)
        eff_bot.Timing.Duration = 0.5
        eff_pulse = slide.TimeLine.MainSequence.AddEffect(m1, MSO_ANIM_EFFECT_TEETER, 0, TRIGGER_AFTER_PREV)
        eff_pulse.Timing.Duration = 0.6

    # ---------------- SLIDE 16: Future Roadmap & Conclusion ----------------
    def build_slide_16(self):
        print("Building Slide 16: Future Roadmap & Conclusion...")
        slide = self.create_slide(TRANS_ORBIT, 1.1)
        self.add_header(slide, "Future Scope & Conclusion", "THE ROADMAP FORWARD & CONCLUSION", "Emerging Capabilities & Transformative Impact on the Global Football Ecosystem")
        
        # 3 Full-Height Future Roadmap Cards
        r1 = self.add_card(slide, 40, 135, 275, 305, "1. 3D Stereoscopic Offside", [
            "• Multi-Camera Epipolar Geometry:",
            "Stereoscopic 3D projection to render millimeter offside lines dynamically.",
            "• Skeletal Keypoint Pose Estimation:",
            "Tracks player shoulders, knees, and toe positions to eliminate broadcast camera parallax error.",
            "• Sub-Second Automated Ruling:",
            "Generates instant virtual offside line visuals for broadcast and sideline monitors."
        ], COLOR_CYAN)
        
        r2 = self.add_card(slide, 340, 135, 275, 305, "2. Autonomous Drone Feeds", [
            "• Bird's-Eye Tactical Views:",
            "Autonomous drone tethering to track spatial formation shifts and off-the-ball fouls.",
            "• Dynamic Altitude Zooming:",
            "Adjusts camera elevation dynamically based on match speed and ball trajectory.",
            "• Automated Tactical Framing:",
            "AI follows ball progression without human pilot intervention."
        ], COLOR_EMERALD)
        
        r3 = self.add_card(slide, 640, 135, 275, 305, "3. Turnkey Grassroots Kit", [
            "• Plug-and-Play Field Kit:",
            "Self-contained weatherproof kit for schools, colleges, and regional federations.",
            "• Mobile Tablet Studio HUD:",
            "Native touch dashboard for pitch-side referee, coaching, and medical staff.",
            "• Cloud Analytics Hub:",
            "Post-match export for heatmaps, tactical dominance, and player workload analysis."
        ], COLOR_PURPLE)

        for i, card in enumerate([r1, r2, r3]):
            eff = slide.TimeLine.MainSequence.AddEffect(card, MSO_ANIM_EFFECT_FLY, 0, TRIGGER_WITH_PREV)
            eff.Timing.Duration = 0.5
            eff.Timing.TriggerDelayTime = 0.2 + (i * 0.15)
            eff.EffectParameters.Direction = DIR_BOTTOM

        # Bottom Conclusion Banner
        bot = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 455, 880, 45)
        bot.Fill.Solid()
        bot.Fill.ForeColor.RGB = COLOR_BG_CARD_LIGHT
        bot.Line.ForeColor.RGB = COLOR_EMERALD
        bot.Line.Weight = 1.5
        tr_b = bot.TextFrame.TextRange
        tr_b.Text = "CONCLUSION: Football Vision AI establishes a scalable, low-cost autonomous VAR and tactical analytics paradigm, bridging the multi-million dollar divide in modern sports."
        tr_b.Font.Name = "Segoe UI"
        tr_b.Font.Size = 11
        tr_b.Font.Bold = True
        tr_b.Font.Color.RGB = COLOR_WHITE
        
        eff_bot = slide.TimeLine.MainSequence.AddEffect(bot, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_AFTER_PREV)
        eff_bot.Timing.Duration = 0.5
        eff_pulse = slide.TimeLine.MainSequence.AddEffect(bot, MSO_ANIM_EFFECT_PULSE, 0, TRIGGER_AFTER_PREV)
        eff_pulse.Timing.Duration = 0.5

    # ---------------- SLIDE 17: Dedicated Thank You Slide ----------------
    def build_slide_17(self):
        print("Building Slide 17: Dedicated Thank You Slide...")
        slide = self.create_slide(TRANS_VORTEX, 1.2)
        
        # Center Badge Pill
        badge = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 355, 90, 250, 34)
        badge.Fill.Solid()
        badge.Fill.ForeColor.RGB = COLOR_BG_CARD_LIGHT
        badge.Line.Visible = True
        badge.Line.ForeColor.RGB = COLOR_CYAN
        badge.Line.Weight = 1.5
        tr_b = badge.TextFrame.TextRange
        tr_b.Text = "  // QUESTIONS & DISCUSSION  "
        tr_b.Font.Name = "Segoe UI"
        tr_b.Font.Size = 12
        tr_b.Font.Bold = True
        tr_b.Font.Color.RGB = COLOR_CYAN
        tr_b.ParagraphFormat.Alignment = 2 # Center
        
        eff_b = slide.TimeLine.MainSequence.AddEffect(badge, MSO_ANIM_EFFECT_BOUNCE, 0, TRIGGER_ON_CLICK)
        eff_b.Timing.Duration = 0.8
        
        # Giant Hero "THANK YOU!"
        ty_box = slide.Shapes.AddTextbox(1, 80, 150, 800, 100)
        ty_box.TextFrame.WordWrap = True
        tr_ty = ty_box.TextFrame.TextRange
        tr_ty.Text = "THANK YOU!"
        tr_ty.Font.Name = "Trebuchet MS"
        tr_ty.Font.Size = 68
        tr_ty.Font.Bold = True
        tr_ty.Font.Color.RGB = COLOR_EMERALD
        tr_ty.ParagraphFormat.Alignment = 2 # Center
        
        eff_ty = slide.TimeLine.MainSequence.AddEffect(ty_box, MSO_ANIM_EFFECT_GROW_AND_TURN, 0, TRIGGER_WITH_PREV)
        eff_ty.Timing.Duration = 0.8
        eff_ty.Timing.TriggerDelayTime = 0.2
        
        # Subtitle: "Questions & Feedback"
        sub_box = slide.Shapes.AddTextbox(1, 80, 265, 800, 45)
        sub_box.TextFrame.WordWrap = True
        tr_sub = sub_box.TextFrame.TextRange
        tr_sub.Text = "We are now open for Questions & Discussion"
        tr_sub.Font.Name = "Segoe UI"
        tr_sub.Font.Size = 22
        tr_sub.Font.Bold = True
        tr_sub.Font.Color.RGB = COLOR_WHITE
        tr_sub.ParagraphFormat.Alignment = 2 # Center
        
        eff_sub = slide.TimeLine.MainSequence.AddEffect(sub_box, MSO_ANIM_EFFECT_FLOAT, 0, TRIGGER_WITH_PREV)
        eff_sub.Timing.Duration = 0.6
        eff_sub.Timing.TriggerDelayTime = 0.4
        
        # Glowing Central Animated VAR AI Football Badge
        center_ball = slide.Shapes.AddShape(MSO_SHAPE_OVAL, 435, 335, 90, 90)
        center_ball.Fill.Solid()
        center_ball.Fill.ForeColor.RGB = COLOR_WHITE
        center_ball.Line.ForeColor.RGB = COLOR_EMERALD
        center_ball.Line.Weight = 4.0
        tr_cb = center_ball.TextFrame.TextRange
        tr_cb.Text = "⚽\nVAR AI"
        tr_cb.Font.Name = "Trebuchet MS"
        tr_cb.Font.Size = 14
        tr_cb.Font.Bold = True
        tr_cb.Font.Color.RGB = COLOR_BG_DARK
        tr_cb.ParagraphFormat.Alignment = 2 # Center
        
        eff_cb = slide.TimeLine.MainSequence.AddEffect(center_ball, MSO_ANIM_EFFECT_SWIVEL, 0, TRIGGER_AFTER_PREV)
        eff_cb.Timing.Duration = 0.9
        
        eff_pulse = slide.TimeLine.MainSequence.AddEffect(center_ball, MSO_ANIM_EFFECT_PULSE, 0, TRIGGER_AFTER_PREV)
        eff_pulse.Timing.Duration = 0.6
        
        eff_spin = slide.TimeLine.MainSequence.AddEffect(center_ball, MSO_ANIM_EFFECT_SPIN, 0, TRIGGER_AFTER_PREV)
        eff_spin.Timing.Duration = 1.0
        
        # Bottom decorative pill
        pill = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 280, 450, 400, 36)
        pill.Fill.Solid()
        pill.Fill.ForeColor.RGB = COLOR_BG_CARD_LIGHT
        pill.Line.ForeColor.RGB = COLOR_CYAN
        pill.Line.Weight = 1.2
        tr_p = pill.TextFrame.TextRange
        tr_p.Text = "FOOTBALL VISION AI • AUTONOMOUS VAR SYSTEM"
        tr_p.Font.Name = "Segoe UI"
        tr_p.Font.Size = 11
        tr_p.Font.Bold = True
        tr_p.Font.Color.RGB = COLOR_CYAN
        tr_p.ParagraphFormat.Alignment = 2
        
        eff_pill = slide.TimeLine.MainSequence.AddEffect(pill, MSO_ANIM_EFFECT_ZOOM, 0, TRIGGER_AFTER_PREV)
        eff_pill.Timing.Duration = 0.5

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
        print("Presentation generation completed successfully!")

if __name__ == "__main__":
    generator = DeckGenerator()
    generator.generate_all()
