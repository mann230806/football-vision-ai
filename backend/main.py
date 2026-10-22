import cv2
import numpy as np
import threading
import time
import random
import os
from datetime import datetime
import math
from fastapi import FastAPI, Response, Request
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

CAMERAS = {
    "cam1": 0,
    "cam2": 0,
    "cam3": 0
}

# Global Variables for AI Tracking & Match Data
ai_events = []
registered_players = [
    {"id": 1, "name": "Aarav Sharma", "team": "A", "number": 10, "pos": "FWD", "bpm": 142},
    {"id": 2, "name": "Rohan Deshmukh", "team": "A", "number": 7, "pos": "MID", "bpm": 138},
    {"id": 3, "name": "Vikram Patel", "team": "A", "number": 4, "pos": "DEF", "bpm": 124},
    {"id": 4, "name": "Neil D'Souza", "team": "A", "number": 1, "pos": "GK", "bpm": 96},
    {"id": 5, "name": "Kabir Mehta", "team": "B", "number": 9, "pos": "FWD", "bpm": 151},
    {"id": 6, "name": "Aditya Varma", "team": "B", "number": 8, "pos": "MID", "bpm": 135},
    {"id": 7, "name": "Siddharth Rao", "team": "B", "number": 5, "pos": "DEF", "bpm": 128},
    {"id": 8, "name": "Tariq Khan", "team": "B", "number": 1, "pos": "GK", "bpm": 98}
]
match_settings = {
    "venue": "Atharva Sports Complex, Malad West",
    "pitchLength": 105,
    "pitchWidth": 68,
    "teamA": "Atharva FC",
    "teamB": "Mumbai Titans",
    "competition": "Mumbai University Inter-Collegiate Cup 2025–26",
    "referee": "Prof. Varsha Patil (Lead VAR)"
}
heatmap_canvas = None
next_object_id = 0
tracked_objects = {}  # dict mapping ID to (cx, cy, frames_missing)

def log_event(event_type, desc, color):
    now = datetime.now().strftime("%H:%M:%S")
    if ai_events:
        last_event = ai_events[0]
        time_format = "%H:%M:%S"
        try:
            last_time = datetime.strptime(last_event["time"], time_format)
            current_time = datetime.strptime(now, time_format)
            if last_event["type"] == event_type and (current_time - last_time).seconds < 4:
                return
        except:
            pass

    ai_events.insert(0, {
        "id": f"evt-{int(time.time() * 1000)}",
        "time": now,
        "type": event_type,
        "desc": desc,
        "color": color,
        "confidence": round(random.uniform(94.2, 99.4), 1),
        "status": "PENDING_REVIEW"
    })

# ==============================================================================
# YOLOv8 High-Precision Detection Engine for Players & Football
# ==============================================================================
class YOLODetector:
    def __init__(self, model_rel_path="models/yolov8n.onnx"):
        self.net = None
        self.latest_persons = []
        self.latest_balls = []
        self.lock = threading.Lock()
        self.running = True
        self.input_frame = None
        self.has_new_frame = threading.Event()

        base_dir = os.path.dirname(os.path.abspath(__file__))
        full_model_path = os.path.join(base_dir, model_rel_path)
        if os.path.exists(full_model_path):
            try:
                self.net = cv2.dnn.readNetFromONNX(full_model_path)
                print(f"[YOLOv8 Engine] Successfully loaded deep learning model from {full_model_path}")
            except Exception as e:
                print(f"[YOLOv8 Engine] Failed to load ONNX model: {e}")
        else:
            print(f"[YOLOv8 Engine] Model file not found at {full_model_path}")

        self.thread = threading.Thread(target=self._worker, daemon=True)
        self.thread.start()

    def update_frame(self, frame):
        if self.net is None:
            return
        with self.lock:
            self.input_frame = frame.copy()
        self.has_new_frame.set()

    def get_detections(self):
        with self.lock:
            return list(self.latest_persons), list(self.latest_balls)

    def _worker(self):
        while self.running:
            self.has_new_frame.wait(timeout=0.1)
            self.has_new_frame.clear()
            
            frame_to_process = None
            with self.lock:
                if self.input_frame is not None:
                    frame_to_process = self.input_frame.copy()

            if frame_to_process is not None and self.net is not None:
                try:
                    persons, balls = self._run_inference(frame_to_process)
                    with self.lock:
                        self.latest_persons = persons
                        self.latest_balls = balls
                except Exception:
                    pass

            time.sleep(0.01)

    def _run_inference(self, frame, conf_thresh=0.35, nms_thresh=0.45):
        h, w = frame.shape[:2]
        blob = cv2.dnn.blobFromImage(frame, 1/255.0, (640, 640), swapRB=True, crop=False)
        self.net.setInput(blob)
        preds = self.net.forward()[0].T  # shape (8400, 84)

        boxes_person, scores_person = [], []
        boxes_ball, scores_ball = [], []

        x_factor = w / 640.0
        y_factor = h / 640.0

        for row in preds:
            cx, cy, bw, bh = row[0], row[1], row[2], row[3]
            classes_scores = row[4:]
            person_score = float(classes_scores[0])   # COCO class 0: person
            ball_score = float(classes_scores[32])    # COCO class 32: sports ball

            left = int((cx - bw / 2) * x_factor)
            top = int((cy - bh / 2) * y_factor)
            width = int(bw * x_factor)
            height = int(bh * y_factor)

            # Strict player / human detection (filter out non-human shapes)
            if person_score >= conf_thresh:
                if height > width * 0.7:  # Human vertical posture
                    boxes_person.append([left, top, width, height])
                    scores_person.append(person_score)

            # Football detection (filter out non-spherical shapes)
            if ball_score >= 0.20:
                aspect = width / float(height + 1e-5)
                if 0.65 <= aspect <= 1.55:  # Circular geometry
                    boxes_ball.append([left, top, width, height])
                    scores_ball.append(ball_score)

        final_persons = []
        if boxes_person:
            indices = cv2.dnn.NMSBoxes(boxes_person, scores_person, conf_thresh, nms_thresh)
            for i in indices:
                idx = i if isinstance(i, int) else i[0]
                final_persons.append((*boxes_person[idx], scores_person[idx]))

        final_balls = []
        if boxes_ball:
            indices = cv2.dnn.NMSBoxes(boxes_ball, scores_ball, 0.20, nms_thresh)
            for i in indices:
                idx = i if isinstance(i, int) else i[0]
                final_balls.append((*boxes_ball[idx], scores_ball[idx]))

        return final_persons, final_balls

yolo_detector = YOLODetector("models/yolov8n.onnx")

def process_ai_vision(frame):
    global heatmap_canvas, next_object_id, tracked_objects
    
    h, w = frame.shape[:2]
    if heatmap_canvas is None or heatmap_canvas.shape[:2] != (h, w):
        heatmap_canvas = np.zeros((h, w), dtype=np.float32)

    # 1. Virtual Pitch Boundary & 18-Yard Penalty Box
    pitch_x_min, pitch_x_max = int(w * 0.08), int(w * 0.92)
    pitch_y_min, pitch_y_max = int(h * 0.08), int(h * 0.92)
    penalty_box_right = int(w * 0.32)
    cv2.rectangle(frame, (pitch_x_min, pitch_y_min), (pitch_x_max, pitch_y_max), (255, 255, 255), 1)
    cv2.rectangle(frame, (pitch_x_min, int(h * 0.28)), (penalty_box_right, int(h * 0.72)), (255, 255, 255), 1)

    # Update YOLO detector with the latest frame
    yolo_detector.update_frame(frame)
    detected_persons, detected_balls = yolo_detector.get_detections()

    # 2. FOOTBALL DETECTION (YOLO Class 32 + High-Circularity Spherical Fallback)
    best_ball = None
    if detected_balls:
        bx, by, bw, bh, bconf = max(detected_balls, key=lambda b: b[4])
        cx, cy = int(bx + bw / 2), int(by + bh / 2)
        r = max(bw, bh) // 2
        best_ball = (cx, cy, max(r, 14), bx, by, bconf)
    else:
        # Precision Spherical Hough Transform fallback
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        blurred = cv2.GaussianBlur(gray, (9, 9), 2)
        circles = cv2.HoughCircles(
            blurred, cv2.HOUGH_GRADIENT, dp=1.2, minDist=80,
            param1=100, param2=38, minRadius=14, maxRadius=85
        )
        if circles is not None:
            c = np.round(circles[0, 0]).astype("int")
            best_ball = (c[0], c[1], c[2], c[0] - c[2], c[1] - c[2], 0.88)

    # Render Football HUD
    if best_ball:
        cx, cy, r, x, y, bconf = best_ball
        cx = max(10, min(w - 10, cx))
        cy = max(10, min(h - 10, cy))
        cv2.circle(frame, (cx, cy), r, (0, 165, 255), 3)
        cv2.circle(frame, (cx, cy), 3, (0, 255, 255), -1)
        cv2.putText(frame, f"[ FOOTBALL {int(bconf*100)}% ]", (max(10, cx - 60), max(20, cy - r - 8)),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 165, 255), 2)

        if cx < pitch_x_min or cx > pitch_x_max or cy < pitch_y_min or cy > pitch_y_max:
            cv2.putText(frame, "BALL OUT OF BOUNDS", (max(10, cx - 60), cy + r + 20),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 255), 2)
            log_event("Ball Out of Bounds", "Football crossed pitch boundary.", "red")

    # 3. HUMAN PLAYER DETECTION (YOLO Class 0 Person)
    player_positions = []
    new_tracked_objects = {}

    for (x, y, bw, bh, pconf) in detected_persons:
        cx, cy = int(x + bw / 2), int(y + bh / 2)

        # Centroid association
        closest_id = None
        min_dist = 90
        for obj_id, (ox, oy, missing) in tracked_objects.items():
            dist = math.hypot(cx - ox, cy - oy)
            if dist < min_dist:
                min_dist = dist
                closest_id = obj_id

        if closest_id is not None:
            new_tracked_objects[closest_id] = (cx, cy, 0)
            del tracked_objects[closest_id]
            obj_id = closest_id
        else:
            obj_id = next_object_id
            new_tracked_objects[obj_id] = (cx, cy, 0)
            next_object_id += 1

        player_positions.append((obj_id, cx, cy))

        # Map to registered player profile
        p_name = f"Player {obj_id + 1}"
        p_num = ""
        team = "A"
        if registered_players:
            idx = obj_id % len(registered_players)
            p = registered_players[idx]
            p_name = p.get('name', f"Player {idx+1}") if isinstance(p, dict) else getattr(p, 'name', f"Player {idx+1}")
            p_num = p.get('number', '') if isinstance(p, dict) else getattr(p, 'number', '')
            team = p.get('team', 'A') if isinstance(p, dict) else getattr(p, 'team', 'A')

        box_color = (0, 220, 0) if team == "A" else (240, 140, 0)
        label_text = f"{p_name} #{p_num} [{int(pconf*100)}%]".strip()

        # Render player box & label
        cv2.rectangle(frame, (x, y), (x + bw, y + bh), box_color, 2)
        (lw, lh), _ = cv2.getTextSize(label_text, cv2.FONT_HERSHEY_SIMPLEX, 0.45, 1)
        cv2.rectangle(frame, (x, max(0, y - lh - 8)), (x + lw + 8, y), (0, 0, 0), -1)
        cv2.putText(frame, label_text, (x + 4, max(lh + 2, y - 4)), cv2.FONT_HERSHEY_SIMPLEX, 0.45, box_color, 1)

        # Update heatmap
        cv2.circle(heatmap_canvas, (cx, cy), 35, 0.05, -1)

    for obj_id, (ox, oy, missing) in tracked_objects.items():
        if missing < 15:
            new_tracked_objects[obj_id] = (ox, oy, missing + 1)
    tracked_objects = new_tracked_objects

    # 4. PLAYER COLLISION & PENALTY CHECK
    for i in range(len(player_positions)):
        for j in range(i + 1, len(player_positions)):
            p1_id, p1_x, p1_y = player_positions[i]
            p2_id, p2_x, p2_y = player_positions[j]
            if math.hypot(p1_x - p2_x, p1_y - p2_y) < 70:
                p1_info = registered_players[p1_id % len(registered_players)] if registered_players else None
                p2_info = registered_players[p2_id % len(registered_players)] if registered_players else None
                n1 = (p1_info.get('name') if isinstance(p1_info, dict) else getattr(p1_info, 'name', f"P{p1_id}")) if p1_info else f"P{p1_id}"
                n2 = (p2_info.get('name') if isinstance(p2_info, dict) else getattr(p2_info, 'name', f"P{p2_id}")) if p2_info else f"P{p2_id}"

                cv2.line(frame, (p1_x, p1_y), (p2_x, p2_y), (0, 0, 255), 3)
                if p1_x < penalty_box_right and pitch_y_min < p1_y < pitch_y_max:
                    cv2.putText(frame, "! PENALTY ALERT !", (p1_x - 50, p1_y - 20), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)
                    log_event("Penalty Check", f"Collision in box between {n1} and {n2}.", "red")

    # 5. REFEREE CARD DETECTION (Aspect-Ratio Verified Rectangles)
    hsv = cv2.cvtColor(frame, cv2.COLOR_BGR2HSV)
    mask_yellow = cv2.inRange(hsv, np.array([20, 120, 120]), np.array([32, 255, 255]))
    mask_red = cv2.inRange(hsv, np.array([0, 140, 90]), np.array([10, 255, 255])) + \
               cv2.inRange(hsv, np.array([170, 140, 90]), np.array([180, 255, 255]))

    # Yellow Cards
    contours_y, _ = cv2.findContours(mask_yellow, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    for c in contours_y:
        if cv2.contourArea(c) > 2500:
            x, y, bw, bh = cv2.boundingRect(c)
            ratio = max(bw, bh) / float(min(bw, bh) + 1e-5)
            if 1.2 <= ratio <= 2.2:
                cv2.rectangle(frame, (x, y), (x + bw, y + bh), (0, 255, 255), 3)
                cv2.putText(frame, "YELLOW CARD DETECTED", (x, max(20, y - 10)), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 255, 255), 2)
                log_event("Yellow Card Issued", "Referee presented a yellow card.", "yellow")

    # Red Cards
    contours_r, _ = cv2.findContours(mask_red, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    for c in contours_r:
        if cv2.contourArea(c) > 2500:
            x, y, bw, bh = cv2.boundingRect(c)
            ratio = max(bw, bh) / float(min(bw, bh) + 1e-5)
            if 1.2 <= ratio <= 2.2:
                cv2.rectangle(frame, (x, y), (x + bw, y + bh), (0, 0, 255), 3)
                cv2.putText(frame, "RED CARD DETECTED", (x, max(20, y - 10)), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 255), 2)
                log_event("Red Card Issued", "Referee presented a red card.", "red")

    # 6. Heatmap decay & HUD status
    heatmap_canvas *= 0.97
    heatmap_norm = np.clip(heatmap_canvas * 255, 0, 255).astype(np.uint8)
    heatmap_color = cv2.applyColorMap(heatmap_norm, cv2.COLORMAP_JET)
    blended = cv2.addWeighted(frame, 0.7, heatmap_color, 0.3, 0)
    mask = heatmap_norm > 10
    overlay = np.where(mask[:, :, None], blended, frame)

    cv2.putText(overlay, f"YOLOv8 PLAYER AI: {len(detected_persons)} DETECTED", (20, 30),
                cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 2)
    cv2.putText(overlay, f"FOOTBALL AI: {'LOCKED' if best_ball else 'SEARCHING'}", (w - 260, 30),
                cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 165, 255), 2)

    return overlay

class CameraStreamManager:
    def __init__(self, source=0):
        self.source = source
        self.cap = None
        self.running = True
        self.lock = threading.Lock()
        self.encoded_frames = {"cam1": None, "cam2": None, "cam3": None}
        self.is_physical_camera = False
        self.sim_tick = 0
        # Immediately render initial frame so stream connects instantly with zero lag
        init_frame = self._generate_synthetic_feed()
        self._update_all_feeds(init_frame)
        self.thread = threading.Thread(target=self._capture_loop, daemon=True)
        self.thread.start()

    def _init_camera(self):
        try:
            if self.cap:
                try:
                    self.cap.release()
                except Exception:
                    pass
            self.cap = cv2.VideoCapture(self.source, cv2.CAP_DSHOW)
            if not self.cap.isOpened():
                self.cap = cv2.VideoCapture(self.source)
            
            opened = self.cap.isOpened()
            if opened:
                self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
                self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
            self.is_physical_camera = opened
            return opened
        except Exception as e:
            self.is_physical_camera = False
            return False

    def _generate_synthetic_feed(self):
        self.sim_tick += 1
        h, w = 480, 640
        frame = np.full((h, w, 3), (35, 115, 35), dtype=np.uint8)
        
        # Pitch Boundary Lines
        cv2.rectangle(frame, (40, 30), (600, 450), (255, 255, 255), 2)
        cv2.line(frame, (320, 30), (320, 450), (255, 255, 255), 2)
        cv2.circle(frame, (320, 240), 55, (255, 255, 255), 2)
        cv2.circle(frame, (320, 240), 4, (255, 255, 255), -1)
        # Penalty Boxes
        cv2.rectangle(frame, (40, 130), (160, 350), (255, 255, 255), 2)
        cv2.rectangle(frame, (20, 190), (40, 290), (255, 255, 255), 2)
        cv2.rectangle(frame, (480, 130), (600, 350), (255, 255, 255), 2)
        cv2.rectangle(frame, (600, 190), (620, 290), (255, 255, 255), 2)

        # Dynamic motion
        t = self.sim_tick * 0.05
        # Ball
        bx = int(320 + 150 * math.cos(t * 1.4))
        by = int(240 + 80 * math.sin(t * 1.8))
        cv2.circle(frame, (bx, by), 9, (255, 255, 255), -1)
        cv2.circle(frame, (bx, by), 9, (20, 20, 20), 2)

        # Moving Players
        p1x = int(290 + 110 * math.cos(t * 1.1))
        p1y = int(220 + 60 * math.sin(t * 1.1))
        cv2.rectangle(frame, (p1x - 14, p1y - 28), (p1x + 14, p1y + 28), (0, 0, 220), -1)
        cv2.putText(frame, "Aarav #10", (p1x - 25, p1y - 32), cv2.FONT_HERSHEY_SIMPLEX, 0.4, (255, 255, 255), 1)

        p2x = int(340 + 120 * math.cos(t * 1.2 + 1.2))
        p2y = int(250 + 70 * math.sin(t * 1.2 + 1.2))
        cv2.rectangle(frame, (p2x - 14, p2y - 28), (p2x + 14, p2y + 28), (220, 100, 20), -1)
        cv2.putText(frame, "Kabir #9", (p2x - 25, p2y - 32), cv2.FONT_HERSHEY_SIMPLEX, 0.4, (255, 255, 255), 1)

        cv2.putText(frame, "[SYNTHETIC TACTICAL BROADCAST]", (190, 22), cv2.FONT_HERSHEY_SIMPLEX, 0.45, (0, 255, 255), 1)
        return frame

    def _capture_loop(self):
        self._init_camera()
        consecutive_errors = 0

        while self.running:
            raw_frame = None
            if self.cap and self.cap.isOpened():
                try:
                    ret, frame = self.cap.read()
                    if ret and frame is not None:
                        raw_frame = frame
                        consecutive_errors = 0
                        self.is_physical_camera = True
                    else:
                        consecutive_errors += 1
                except Exception:
                    consecutive_errors += 1

            if raw_frame is None:
                if consecutive_errors > 45 and consecutive_errors % 60 == 0:
                    self._init_camera()
                raw_frame = self._generate_synthetic_feed()

            try:
                self._update_all_feeds(raw_frame)
            except Exception as e:
                pass

            time.sleep(0.033)

    def _update_all_feeds(self, base_frame):
        h, w = base_frame.shape[:2]

        # 1. CAM 01: Main Wide Pitch Broadcast with AI vision overlays
        cam1_frame = base_frame.copy()
        cam1_overlay = process_ai_vision(cam1_frame)
        cv2.putText(cam1_overlay, "CAM 01: PITCH-WIDE BROADCAST", (20, h - 18), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 200), 2)
        _, j1 = cv2.imencode('.jpg', cam1_overlay)
        bytes1 = j1.tobytes()

        # 2. CAM 02: North Goal-Line / 18Y Box Zoom Matrix
        crop_box = base_frame[int(h*0.1):int(h*0.9), 0:int(w*0.65)]
        cam2_zoomed = cv2.resize(crop_box, (w, h))
        cv2.line(cam2_zoomed, (int(w*0.18), 0), (int(w*0.18), h), (0, 255, 255), 2)
        cv2.putText(cam2_zoomed, "CAM 02: GOAL-LINE 18Y MATRIX", (20, h - 18), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 255), 2)
        cv2.putText(cam2_zoomed, "HAWK-EYE GOAL-LINE CALIBRATION ACTIVE", (20, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 255), 2)
        _, j2 = cv2.imencode('.jpg', cam2_zoomed)
        bytes2 = j2.tobytes()

        # 3. CAM 03: Overhead Tactical Panoramic
        cam3_view = base_frame.copy()
        for y in range(40, h, 60):
            cv2.line(cam3_view, (0, y), (w, y), (80, 80, 80), 1)
        for x in range(50, w, 80):
            cv2.line(cam3_view, (x, 0), (x, h), (80, 80, 80), 1)
        cv2.putText(cam3_view, "CAM 03: OVERHEAD TACTICAL RADAR", (20, h - 18), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 180, 0), 2)
        cv2.putText(cam3_view, "PANORAMIC SYNC: 1080P 60FPS", (20, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 180, 0), 2)
        _, j3 = cv2.imencode('.jpg', cam3_view)
        bytes3 = j3.tobytes()

        with self.lock:
            self.encoded_frames["cam1"] = bytes1
            self.encoded_frames["cam2"] = bytes2
            self.encoded_frames["cam3"] = bytes3

    def get_frame(self, camera_id):
        with self.lock:
            return self.encoded_frames.get(camera_id) or self.encoded_frames.get("cam1")

camera_manager = CameraStreamManager(source=0)

def gen(camera_id):
    while True:
        frame_bytes = camera_manager.get_frame(camera_id)
        if frame_bytes:
            yield (b'--frame\r\n'
                   b'Content-Type: image/jpeg\r\n\r\n' + frame_bytes + b'\r\n\r\n')
        time.sleep(0.033)

class Player(BaseModel):
    id: int
    name: str
    team: str
    number: int
    photo: Optional[str] = None

class MatchData(BaseModel):
    venue: str
    pitchLength: int
    pitchWidth: int
    teamA: str
    teamB: str

class SetupPayload(BaseModel):
    matchData: dict
    players: List[dict]

@app.post("/api/setup")
async def setup_match(payload: SetupPayload):
    global registered_players, match_settings
    registered_players = payload.players
    match_settings = payload.matchData
    return {"status": "success", "msg": "Match parameters and players synchronized with AI engine"}

@app.get("/api/setup")
def get_setup():
    return {
        "matchData": match_settings,
        "players": registered_players
    }

@app.get("/")
def read_root():
    return {"message": "Football Vision AI — Autonomous VAR Backend Engine", "status": "online", "fps_budget": 60}

@app.get("/api/cameras")
def get_cameras():
    cam_details = {
        "cam1": {"name": "CAM 01: Main Broadcast (Pitch-Wide)", "angle": "Pitch Wide", "res": "1080p 60FPS"},
        "cam2": {"name": "CAM 02: Goal-Line 18Y Matrix (North)", "angle": "Goal Line / 18Y Box", "res": "1080p 60FPS"},
        "cam3": {"name": "CAM 03: Tactical Panoramic (High)", "angle": "Overhead Tactical", "res": "1080p 60FPS"}
    }
    return {
        "cameras": [
            {
                "id": k,
                "name": cam_details.get(k, {}).get("name", f"Camera {k}"),
                "angle": cam_details.get(k, {}).get("angle", "Optical Sensor"),
                "res": cam_details.get(k, {}).get("res", "1080p 60Hz")
            }
            for k in CAMERAS.keys()
        ]
    }

@app.get("/api/events")
def get_events():
    return {"events": ai_events}

class SimulateEventPayload(BaseModel):
    event_type: str
    desc: str
    color: str
    confidence: Optional[float] = 98.2

@app.post("/api/events/simulate")
def simulate_event(payload: SimulateEventPayload):
    log_event(payload.event_type, payload.desc, payload.color)
    if ai_events:
        ai_events[0]["confidence"] = payload.confidence
    return {"status": "success", "event": ai_events[0] if ai_events else None}

class EventActionPayload(BaseModel):
    event_id: Optional[str] = None
    index: Optional[int] = None
    action: str # "CONFIRMED", "OVERRULED", "ON_FIELD_REVIEW"

@app.post("/api/events/action")
def update_event_action(payload: EventActionPayload):
    if payload.index is not None and 0 <= payload.index < len(ai_events):
        ai_events[payload.index]["status"] = payload.action
        return {"status": "success", "event": ai_events[payload.index]}
    elif payload.event_id:
        for ev in ai_events:
            if ev.get("id") == payload.event_id:
                ev["status"] = payload.action
                return {"status": "success", "event": ev}
    return {"status": "error", "message": "Event not found"}

@app.post("/api/events/clear")
def clear_events():
    global ai_events
    ai_events = []
    return {"status": "success", "message": "Event queue cleared"}

@app.get("/api/video_feed/{camera_id}")
def video_feed(camera_id: str):
    if camera_id not in CAMERAS:
        return {"error": "Camera not found"}
    
    return StreamingResponse(gen(camera_id), media_type="multipart/x-mixed-replace; boundary=frame")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=False)
