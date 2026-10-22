"""
Football Vision AI - Presentation Generator with High-Impact Animations
Generates a 16:9 widescreen presentation using PowerPoint COM automation.
"""

import os
import sys
import time
import win32com.client

# PowerPoint Constants
MSO_SHAPE_RECTANGLE = 1
MSO_SHAPE_ROUNDED_RECTANGLE = 5
MSO_SHAPE_OVAL = 9
MSO_SHAPE_UP_ARROW = 33
MSO_SHAPE_RIGHT_ARROW = 34
MSO_SHAPE_BALLOON = 137

# Animation Effect Types
MSO_ANIM_EFFECT_FLY = 2
MSO_ANIM_EFFECT_ZOOM = 6
MSO_ANIM_EFFECT_SPIN = 10
MSO_ANIM_EFFECT_TEETER = 13
MSO_ANIM_EFFECT_FLASH_BULB = 15
MSO_ANIM_EFFECT_GROW_AND_TURN = 17
MSO_ANIM_EFFECT_BOUNCE = 25
MSO_ANIM_EFFECT_FLOAT = 28
MSO_ANIM_EFFECT_PINWHEEL = 30
MSO_ANIM_EFFECT_SPIRAL = 32
MSO_ANIM_EFFECT_SWIVEL = 45
MSO_ANIM_EFFECT_BOOMERANG = 48
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
DIR_TOP_LEFT = 5
DIR_TOP_RIGHT = 6
DIR_BOTTOM_LEFT = 7
DIR_BOTTOM_RIGHT = 8

# Triggers
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
    """COM uses BGR color ordering: R + (G << 8) + (B << 16)"""
    return r + (g << 8) + (b << 16)

# Premium Dark Stadium / Cyberpunk Color Palette
COLOR_BG_DARK = rgb(7, 11, 20)           # #070B14 Deep pitch dark
COLOR_BG_CARD = rgb(15, 23, 42)          # #0F172A Slate navy
COLOR_BG_CARD_LIGHT = rgb(23, 37, 68)    # #172544 Lighter panel
COLOR_EMERALD = rgb(0, 255, 136)         # #00FF88 Neon pitch emerald
COLOR_CYAN = rgb(0, 229, 255)            # #00E5FF Electric VAR cyan
COLOR_PURPLE = rgb(168, 85, 247)         # #A855F7 AI Violet
COLOR_RED = rgb(239, 68, 68)             # #EF4444 Card Red / Warning
COLOR_YELLOW = rgb(245, 158, 11)         # #F59E0B Card Yellow
COLOR_WHITE = rgb(255, 255, 255)
COLOR_TEXT_MUTED = rgb(148, 163, 184)    # Cool Silver
COLOR_BORDER_SUBTLE = rgb(30, 41, 59)

MEDIA_DIR = os.path.abspath("extracted_media")

class PresentationBuilder:
    def __init__(self):
        self.ppt = win32com.client.Dispatch("PowerPoint.Application")
        self.ppt.Visible = True
        self.pres = self.ppt.Presentations.Add()
        self.pres.PageSetup.SlideWidth = 960
        self.pres.PageSetup.SlideHeight = 540
        self.current_slide_num = 0

    def add_slide(self, transition_type=TRANS_ZOOM, transition_duration=1.0):
        self.current_slide_num += 1
        slide = self.pres.Slides.Add(self.current_slide_num, 12) # ppLayoutBlank
        slide.FollowMasterBackground = False
        slide.Background.Fill.Solid()
        slide.Background.Fill.ForeColor.RGB = COLOR_BG_DARK
        
        # Set Transition
        try:
            slide.SlideShowTransition.EntryEffect = transition_type
            slide.SlideShowTransition.Duration = transition_duration
        except:
            pass
            
        # Draw subtle stadium pitch markings / header border
        self._add_pitch_accents(slide)
        return slide

    def _add_pitch_accents(self, slide):
        # Subtle glowing top accent line
        line = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, 40, 18, 880, 2)
        line.Fill.Solid()
        line.Fill.ForeColor.RGB = COLOR_CYAN
        line.Line.Visible = False
        
        # Subtle footer text
        footer = slide.Shapes.AddTextbox(1, 40, 515, 880, 20)
        footer.TextFrame.WordWrap = False
        footer.TextFrame.MarginLeft = 0
        footer.TextFrame.MarginTop = 0
        tr = footer.TextFrame.TextRange
        tr.Text = f"ATHARVA COLLEGE OF ENGINEERING  |  FOOTBALL VISION AI — AUTONOMOUS VAR  |  SLIDE {self.current_slide_num:02d} / 16"
        tr.Font.Name = "Segoe UI"
        tr.Font.Size = 9
        tr.Font.Color.RGB = COLOR_TEXT_MUTED

    def add_header(self, slide, category, title, subtitle):
        """Adds a top header with category pill badge, main title, and subtitle."""
        # 1. Category Pill Badge
        badge = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, 40, 28, 260, 24)
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
        title_box = slide.Shapes.AddTextbox(1, 40, 52, 880, 42)
        title_box.TextFrame.WordWrap = True
        title_box.TextFrame.MarginLeft = 0
        title_box.TextFrame.MarginTop = 0
        tr_t = title_box.TextFrame.TextRange
        tr_t.Text = title
        tr_t.Font.Name = "Segoe UI"
        tr_t.Font.Size = 24
        tr_t.Font.Bold = True
        tr_t.Font.Color.RGB = COLOR_WHITE
        
        eff_t = slide.TimeLine.MainSequence.AddEffect(title_box, MSO_ANIM_EFFECT_GROW_AND_TURN, 0, TRIGGER_WITH_PREV)
        eff_t.Timing.Duration = 0.5
        eff_t.Timing.TriggerDelayTime = 0.1
        
        # 3. Subtitle
        sub_box = slide.Shapes.AddTextbox(1, 40, 94, 880, 26)
        sub_box.TextFrame.WordWrap = True
        sub_box.TextFrame.MarginLeft = 0
        sub_box.TextFrame.MarginTop = 0
        tr_s = sub_box.TextFrame.TextRange
        tr_s.Text = subtitle
        tr_s.Font.Name = "Segoe UI"
        tr_s.Font.Size = 13
        tr_s.Font.Color.RGB = COLOR_EMERALD
        
        eff_s = slide.TimeLine.MainSequence.AddEffect(sub_box, MSO_ANIM_EFFECT_FLOAT, 0, TRIGGER_WITH_PREV)
        eff_s.Timing.Duration = 0.4
        eff_s.Timing.TriggerDelayTime = 0.2

    def add_card(self, slide, left, top, width, height, title, body_paragraphs, border_color=COLOR_CYAN, bg_color=COLOR_BG_CARD):
        """Creates a styled glassmorphic card with header and bulleted content."""
        card = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, left, top, width, height)
        card.Fill.Solid()
        card.Fill.ForeColor.RGB = bg_color
        card.Line.Visible = True
        card.Line.ForeColor.RGB = border_color
        card.Line.Weight = 1.5
        
        tf = card.TextFrame
        tf.WordWrap = True
        tf.MarginLeft = 14
        tf.MarginRight = 14
        tf.MarginTop = 14
        tf.MarginBottom = 14
        
        full_text = title + "\n\n" + "\n".join(body_paragraphs)
        tr = tf.TextRange
        tr.Text = full_text
        
        # Format Title (first paragraph)
        p1 = tr.Paragraphs(1)
        p1.Font.Name = "Segoe UI"
        p1.Font.Size = 16
        p1.Font.Bold = True
        p1.Font.Color.RGB = border_color
        
        # Format Body
        for i in range(2, tr.Paragraphs.Count + 1):
            p = tr.Paragraphs(i)
            p.Font.Name = "Segoe UI"
            p.Font.Size = 12
            p.Font.Color.RGB = COLOR_WHITE if i % 2 == 0 else COLOR_TEXT_MUTED
            
        return card

    def add_metric_badge(self, slide, left, top, width, height, number_str, label_str, accent_color=COLOR_EMERALD):
        """Creates a high-impact metric stat counter box with glowing border."""
        box = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, left, top, width, height)
        box.Fill.Solid()
        box.Fill.ForeColor.RGB = COLOR_BG_CARD_LIGHT
        box.Line.Visible = True
        box.Line.ForeColor.RGB = accent_color
        box.Line.Weight = 2.0
        
        tf = box.TextFrame
        tf.WordWrap = True
        tf.MarginLeft = 8
        tf.MarginRight = 8
        tf.MarginTop = 10
        tf.MarginBottom = 8
        
        tr = tf.TextRange
        tr.Text = f"{number_str}\n{label_str}"
        
        p1 = tr.Paragraphs(1)
        p1.Font.Name = "Trebuchet MS"
        p1.Font.Size = 24
        p1.Font.Bold = True
        p1.Font.Color.RGB = accent_color
        
        p2 = tr.Paragraphs(2)
        p2.Font.Name = "Segoe UI"
        p2.Font.Size = 11
        p2.Font.Bold = True
        p2.Font.Color.RGB = COLOR_WHITE
        
        return box

    def add_image_panel(self, slide, img_name, left, top, width, height, border_color=COLOR_CYAN, title=None):
        """Embeds an image from extracted_media with a styled tech monitor frame."""
        img_path = os.path.join(MEDIA_DIR, img_name)
        if not os.path.exists(img_path):
            print(f"Warning: Image {img_name} not found!")
            return None
            
        # Frame container
        frame = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, left - 4, top - 4, width + 8, height + 8)
        frame.Fill.Solid()
        frame.Fill.ForeColor.RGB = COLOR_BG_CARD
        frame.Line.Visible = True
        frame.Line.ForeColor.RGB = border_color
        frame.Line.Weight = 2.0
        
        # Image
        pic = slide.Shapes.AddPicture(img_path, False, True, left, top, width, height)
        return pic, frame

    def animate_fly(self, shape, direction, delay=0.0, duration=0.5, trigger=TRIGGER_WITH_PREV):
        eff = shape.Parent.TimeLine.MainSequence.AddEffect(shape, MSO_ANIM_EFFECT_FLY, 0, trigger)
        eff.Timing.Duration = duration
        eff.Timing.TriggerDelayTime = delay
        eff.EffectParameters.Direction = direction
        return eff

    def animate_zoom(self, shape, delay=0.0, duration=0.5, trigger=TRIGGER_WITH_PREV):
        eff = shape.Parent.TimeLine.MainSequence.AddEffect(shape, MSO_ANIM_EFFECT_ZOOM, 0, trigger)
        eff.Timing.Duration = duration
        eff.Timing.TriggerDelayTime = delay
        return eff

    def animate_pulse(self, shape, delay=0.0, duration=0.5, trigger=TRIGGER_AFTER_PREV):
        eff = shape.Parent.TimeLine.MainSequence.AddEffect(shape, MSO_ANIM_EFFECT_PULSE, 0, trigger)
        eff.Timing.Duration = duration
        eff.Timing.TriggerDelayTime = delay
        return eff

    def animate_teeter(self, shape, delay=0.0, duration=0.5, trigger=TRIGGER_AFTER_PREV):
        eff = shape.Parent.TimeLine.MainSequence.AddEffect(shape, MSO_ANIM_EFFECT_TEETER, 0, trigger)
        eff.Timing.Duration = duration
        eff.Timing.TriggerDelayTime = delay
        return eff

    def animate_spin(self, shape, delay=0.0, duration=0.8, trigger=TRIGGER_AFTER_PREV):
        eff = shape.Parent.TimeLine.MainSequence.AddEffect(shape, MSO_ANIM_EFFECT_SPIN, 0, trigger)
        eff.Timing.Duration = duration
        eff.Timing.TriggerDelayTime = delay
        return eff

    def animate_motion_path(self, shape, path_type, delay=0.0, duration=1.2, trigger=TRIGGER_AFTER_PREV):
        eff = shape.Parent.TimeLine.MainSequence.AddEffect(shape, path_type, 0, trigger)
        eff.Timing.Duration = duration
        eff.Timing.TriggerDelayTime = delay
        return eff

    def save_and_close(self, filename="Football_Vision_AI_Presentation.pptx"):
        out_path = os.path.abspath(filename)
        # Save as PPTX format (24 = ppSaveAsOpenXMLPresentation)
        try:
            self.pres.SaveAs(out_path, 24)
            print(f"SUCCESS: Presentation saved to {out_path}")
        except Exception as e:
            print(f"Error saving presentation: {e}")
            self.pres.SaveAs(out_path)
        self.pres.Close()
        self.ppt.Quit()

print("PresentationBuilder class initialized successfully.")
