# Football Vision AI — Autonomous VAR & Match Analytics Engine ⚽🤖

An end-to-end intelligent Football Video Assistant Referee (VAR) and tactical computer vision platform designed for real-time player tracking, ball trajectory monitoring, collision and penalty analysis, semi-automated offside (SAOT) line rendering, referee card detection, and IoT biometric health integration.

---

## 🌟 Key Features

1. **YOLOv8 Deep Learning Detection & Tracking**:
   - Neural-network powered detection targeting actual **players** (`person`, Class 0) and **football** (`sports ball`, Class 32).
   - Eliminates false detections from arbitrary background motion, shadows, or desk movement.
   - Real-time centroid tracking mapping on-pitch individuals to registered squad profiles.

2. **Autonomous VAR Decision Assistance**:
   - **Ball Out of Bounds**: Tracks the football's coordinates relative to calibrated pitch boundaries.
   - **Penalty Box Collision Analyzer**: Evaluates proximity between opposing players inside the 18-yard box to flag potential fouls and penalties.
   - **Referee Card Recognition**: Real-time HSV color and aspect-ratio detection for instant Yellow and Red card logging.
   - **Semi-Automated Offside (SAOT) Simulation**: Real-time tactical offside line rendering.

3. **Multi-Camera Matrix Hub**:
   - **CAM 01**: Pitch-Wide Broadcast Perspective with live AI overlays and tactical heatmap.
   - **CAM 02**: Goal-Line / 18-Yard Box Zoom Matrix with Hawk-Eye calibration lines.
   - **CAM 03**: Tactical Overhead Panoramic Radar with pitch grid overlays.
   - Single-worker DirectShow architecture preventing webcam device lock collisions on Windows.

4. **IoT Biometrics & Wearable Telemetry**:
   - Live telemetry monitoring for player heart rates (BPM), sprint counts, pitch distance covered, and stamina metrics.

---

## 🛠️ System Architecture

- **Backend**: Python 3.12, FastAPI, OpenCV (cv2.dnn), YOLOv8 (ONNX), Uvicorn.
- **Frontend**: React 19, Vite, Tailwind CSS, Lucide Icons, React Router.
- **Hardware Integration**: OpenCV DirectShow camera capture + simulated synthetic match fallback.

---

## 🚀 Getting Started

### 1. Prerequisites
- **Python 3.10+** (Tested on Python 3.12)
- **Node.js 18+** & npm

### 2. Backend Setup
```bash
# Navigate to backend directory
cd backend

# Create and activate virtual environment
python -m venv venv
# On Windows:
.\venv\Scripts\activate
# On Linux/macOS:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Start FastAPI server (runs on port 8000)
uvicorn main:app --host 127.0.0.1 --port 8000
```

### 3. Frontend Setup
```bash
# Navigate to frontend directory
cd frontend

# Install packages
npm install

# Start Vite dev server (runs on port 5173)
npm run dev
```

### 4. Access the Dashboard
- **Match Setup & Squad Registration**: [http://127.0.0.1:5173/setup](http://127.0.0.1:5173/setup)
- **VAR Command Center & Live Feed**: [http://127.0.0.1:5173/match](http://127.0.0.1:5173/match)
- **FastAPI Interactive API Docs**: [http://127.0.0.1:8000/docs](http://127.0.0.1:8000/docs)

---

## 📁 Project Structure

```
football-vision-ai/
├── backend/
│   ├── main.py                  # FastAPI server, CameraStreamManager & YOLOv8 pipeline
│   ├── requirements.txt         # Python dependencies
│   └── models/
│       └── yolov8n.onnx         # YOLOv8 nano neural network weights
├── frontend/
│   ├── src/
│   │   ├── Match.jsx            # Main VAR control monitor, camera matrix & replay HUD
│   │   ├── Setup.jsx            # Squad registration and pitch setup
│   │   ├── App.jsx              # Routing
│   │   └── index.css            # Styles & Tailwind setup
│   ├── package.json             # Frontend dependencies
│   └── vite.config.js           # Vite configuration
├── .gitignore
└── README.md
```

---

## 👥 Authors
Developed as part of the **BE Project — Football Vision AI** (Mumbai University Inter-Collegiate Cup VAR Initiative).
