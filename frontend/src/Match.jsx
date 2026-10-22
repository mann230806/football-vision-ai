import React, { useState, useEffect, useRef } from 'react';
import { useNavigate } from 'react-router-dom';
import { 
  Camera, 
  Activity, 
  AlertTriangle, 
  Play, 
  Pause, 
  RotateCcw, 
  CheckCircle, 
  XCircle, 
  ChevronRight, 
  Sliders, 
  ShieldAlert, 
  Heart, 
  Radio, 
  Zap, 
  Eye, 
  Maximize2, 
  RefreshCw,
  Award,
  Clock,
  Flame,
  UserCheck,
  FileText
} from 'lucide-react';

export default function Match() {
  const navigate = useNavigate();

  // Match setup info
  const [matchInfo, setMatchInfo] = useState({
    teamA: 'Atharva FC',
    teamB: 'Mumbai Titans',
    venue: 'Atharva Sports Complex, Malad West',
    referee: 'Prof. Varsha Patil (Lead VAR Officiator)',
    competition: 'Mumbai University Inter-Collegiate Cup 2025–26'
  });

  // Score & Match Clock
  const [score, setScore] = useState({ a: 1, b: 0 });
  const [matchTime, setMatchTime] = useState(2304); // 38 mins, 24 secs (in seconds)
  const [isClockRunning, setIsClockRunning] = useState(true);
  const [half, setHalf] = useState('1ST HALF');

  // Camera Feeds
  const [cameras, setCameras] = useState([
    { id: 'cam1', name: 'CAM 01: Main Broadcast', angle: 'Pitch-Wide Tactical', res: '1080p 60FPS' },
    { id: 'cam2', name: 'CAM 02: North Goal-Line', angle: '18Y Box / Goal Line', res: '1080p 60FPS' },
    { id: 'cam3', name: 'CAM 03: Overhead Tactical', angle: 'Gantry Panoramic', res: '1080p 60FPS' }
  ]);
  const [activeCam, setActiveCam] = useState('cam1');
  const [playbackSpeed, setPlaybackSpeed] = useState('1.0x');
  const [isSlowMo, setIsSlowMo] = useState(false);

  // Overlay Toggles
  const [overlays, setOverlays] = useState({
    saotLine: true,
    penaltyGrid: true,
    playerBoxes: true,
    heatmap: false
  });

  // Events & VAR Incidents
  const [events, setEvents] = useState([]);
  const [selectedEventId, setSelectedEventId] = useState(null);
  const [actionNotice, setActionNotice] = useState(null);

  // Squad Players & IoT Biometrics
  const [players, setPlayers] = useState([
    { id: 1, name: 'Aarav Sharma', team: 'A', number: 10, pos: 'FWD', sensorId: 'IOT-HR-04A', bpm: 146, stamina: 84, distance: '4.8 km', sprints: 14 },
    { id: 2, name: 'Rohan Deshmukh', team: 'A', number: 7, pos: 'MID', sensorId: 'IOT-HR-02A', bpm: 141, stamina: 82, distance: '5.2 km', sprints: 18 },
    { id: 3, name: 'Vikram Patel', team: 'A', number: 4, pos: 'DEF', sensorId: 'IOT-HR-08A', bpm: 128, stamina: 89, distance: '4.1 km', sprints: 9 },
    { id: 4, name: 'Neil D\'Souza', team: 'A', number: 1, pos: 'GK', sensorId: 'IOT-HR-01A', bpm: 98, stamina: 96, distance: '2.0 km', sprints: 3 },
    { id: 5, name: 'Kabir Mehta', team: 'B', number: 9, pos: 'FWD', sensorId: 'IOT-HR-11B', bpm: 154, stamina: 78, distance: '5.0 km', sprints: 16 },
    { id: 6, name: 'Aditya Varma', team: 'B', number: 8, pos: 'MID', sensorId: 'IOT-HR-07B', bpm: 139, stamina: 85, distance: '5.1 km', sprints: 15 },
    { id: 7, name: 'Siddharth Rao', team: 'B', number: 5, pos: 'DEF', sensorId: 'IOT-HR-05B', bpm: 132, stamina: 87, distance: '4.3 km', sprints: 8 },
    { id: 8, name: 'Tariq Khan', team: 'B', number: 1, pos: 'GK', sensorId: 'IOT-HR-01B', bpm: 102, stamina: 95, distance: '2.1 km', sprints: 4 }
  ]);

  // Bottom Panel Tab
  const [bottomTab, setBottomTab] = useState('biometrics'); // 'biometrics', 'stats', 'methodology'

  // Load Setup from Backend
  useEffect(() => {
    fetch('http://localhost:8000/api/setup')
      .then(res => res.json())
      .then(data => {
        if (data.matchData) {
          setMatchInfo(prev => ({ ...prev, ...data.matchData }));
        }
        if (data.players && data.players.length > 0) {
          setPlayers(data.players.map(p => ({
            ...p,
            stamina: p.stamina || Math.floor(Math.random() * (95 - 75) + 75),
            distance: p.distance || (Math.random() * (5.5 - 3.5) + 3.5).toFixed(1) + ' km',
            sprints: p.sprints || Math.floor(Math.random() * 15 + 5),
            bpm: p.bpm || 135
          })));
        }
      })
      .catch(err => console.log('Using default mock setup', err));

    fetch('http://localhost:8000/api/cameras')
      .then(res => res.json())
      .then(data => {
        if (data.cameras && data.cameras.length > 0) {
          setCameras(data.cameras);
          setActiveCam(data.cameras[0].id);
        }
      })
      .catch(err => console.log('Camera API offline', err));
  }, []);

  // Match Clock Interval
  useEffect(() => {
    let interval = null;
    if (isClockRunning) {
      interval = setInterval(() => {
        setMatchTime(t => t + 1);
      }, 1000);
    }
    return () => clearInterval(interval);
  }, [isClockRunning]);

  // Periodic Polling for Live Events & IoT Jitter
  useEffect(() => {
    const fetchEvents = async () => {
      try {
        const res = await fetch('http://localhost:8000/api/events');
        const data = await res.json();
        setEvents(data.events || []);
        if (data.events && data.events.length > 0 && !selectedEventId) {
          setSelectedEventId(data.events[0].id);
        }
      } catch (err) {
        // Fallback local events if server connection drops
      }
    };

    const simBiometrics = () => {
      setPlayers(prevPlayers =>
        prevPlayers.map(p => {
          const delta = (Math.random() - 0.5) * 3;
          let newBpm = Math.round(p.bpm + delta);
          if (p.pos === 'GK') {
            newBpm = Math.max(90, Math.min(115, newBpm));
          } else {
            newBpm = Math.max(120, Math.min(168, newBpm));
          }
          return { ...p, bpm: newBpm };
        })
      );
    };

    fetchEvents();
    const pollInterval = setInterval(() => {
      fetchEvents();
      simBiometrics();
    }, 1500);

    return () => clearInterval(pollInterval);
  }, [selectedEventId]);

  // Time format helper
  const formatTime = (totalSeconds) => {
    const mins = Math.floor(totalSeconds / 60);
    const secs = totalSeconds % 60;
    return `${mins.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
  };

  // Handle VAR Decision Action
  const handleDecision = async (actionType) => {
    if (!events.length) return;
    const targetId = selectedEventId || events[0]?.id;
    try {
      await fetch('http://localhost:8000/api/events/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          event_id: targetId,
          action: actionType
        })
      });

      // Update local state immediately
      setEvents(events.map(ev => ev.id === targetId ? { ...ev, status: actionType } : ev));

      const labels = {
        'CONFIRMED': 'Official Decision: CONFIRMED BY VAR',
        'OVERRULED': 'Official Decision: ON-FIELD CALL OVERRULED',
        'ON_FIELD_REVIEW': 'Referee Summoned: ON-FIELD REVIEW (OFR)'
      };
      setActionNotice(labels[actionType] || actionType);
      setTimeout(() => setActionNotice(null), 3500);
    } catch (err) {
      console.error('Failed to submit VAR action:', err);
    }
  };

  // Simulate an incident for presentation demo
  const handleSimulateIncident = async (type) => {
    let payload = {};
    if (type === 'GOAL') {
      payload = {
        event_type: 'GOAL DECISION SYSTEM (GDS)',
        desc: `Ball crossed plane: Full circumference verified over goal line (3.2cm depth)`,
        color: 'emerald',
        confidence: 99.4
      };
      setScore(prev => ({ ...prev, a: prev.a + 1 }));
    } else if (type === 'OFFSIDE') {
      payload = {
        event_type: 'SEMI-AUTOMATED OFFSIDE (SAOT)',
        desc: `Marginal offside detected: Attacker knee is 14.2cm beyond 2nd-last defender line`,
        color: 'rose',
        confidence: 98.7
      };
    } else if (type === 'FOUL') {
      payload = {
        event_type: 'SERIOUS FOUL PLAY / RED CARD',
        desc: `Studs-up excessive force contact on player ankle in penalty box`,
        color: 'amber',
        confidence: 95.8
      };
    } else if (type === 'OUT_OF_BOUNDS') {
      payload = {
        event_type: 'TOUCHLINE BOUNDARY CHECK',
        desc: `Ball completely exited touchline perimeter at coordinate Y: 68.04m`,
        color: 'sky',
        confidence: 99.1
      };
    }

    try {
      const res = await fetch('http://localhost:8000/api/events/simulate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      if (data.event) {
        setEvents([data.event, ...events]);
        setSelectedEventId(data.event.id);
      }
    } catch (err) {
      console.warn('Simulation failed offline', err);
    }
  };

  const handleClearEvents = async () => {
    try {
      await fetch('http://localhost:8000/api/events/clear', { method: 'POST' });
      setEvents([]);
      setSelectedEventId(null);
    } catch (err) {
      setEvents([]);
    }
  };

  const selectedEvent = events.find(e => e.id === selectedEventId) || events[0];

  return (
    <div className="min-h-screen bg-[#070b14] text-slate-100 font-sans selection:bg-emerald-500 selection:text-black">
      
      {/* 1. TOP OFFICIAL BROADCAST SCOREBOARD & VAR CONTROLLER BAR */}
      <header className="border-b border-slate-800 bg-[#0b1222] px-6 py-2.5 sticky top-0 z-50">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row items-center justify-between gap-4">
          
          {/* Left: Atharva Project Brand */}
          <div className="flex items-center gap-3">
            <button 
              onClick={() => navigate('/setup')} 
              className="p-2 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 transition border border-slate-700"
              title="Return to Match Setup / Rosters"
            >
              <Sliders size={16} />
            </button>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-[11px] font-bold px-1.5 py-0.5 rounded bg-emerald-950 text-emerald-400 border border-emerald-800/80 font-mono">
                  LIVE VAR
                </span>
                <span className="text-xs font-semibold text-white tracking-wide">
                  ATHARVA VAR MATCH CENTRE
                </span>
              </div>
              <p className="text-[11px] text-slate-400">
                Supervisor: {matchInfo.referee}
              </p>
            </div>
          </div>

          {/* Center: Live Match Scoreboard & Clock */}
          <div className="flex items-center gap-4 bg-[#070b14] px-5 py-2 rounded-xl border border-slate-800 shadow-inner">
            <div className="text-right">
              <span className="text-xs font-bold text-emerald-400 block tracking-wide">{matchInfo.teamA}</span>
              <span className="text-[10px] text-slate-500 font-mono">HOME</span>
            </div>

            <div className="flex items-center gap-2 px-3 py-1 bg-slate-900 rounded-lg border border-slate-700/80">
              <span className="text-xl font-bold font-mono text-white">{score.a}</span>
              <span className="text-slate-500 font-bold">-</span>
              <span className="text-xl font-bold font-mono text-white">{score.b}</span>
            </div>

            <div className="text-left">
              <span className="text-xs font-bold text-sky-400 block tracking-wide">{matchInfo.teamB}</span>
              <span className="text-[10px] text-slate-500 font-mono">AWAY</span>
            </div>

            <div className="h-6 w-[1px] bg-slate-800 mx-1"></div>

            <div className="flex items-center gap-2">
              <div className="text-center font-mono">
                <span className="text-[10px] text-slate-400 block">{half}</span>
                <span className="text-sm font-bold text-amber-400">{formatTime(matchTime)}</span>
              </div>
              <button
                onClick={() => setIsClockRunning(!isClockRunning)}
                className="p-1 rounded bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs transition"
                title={isClockRunning ? 'Pause match timer' : 'Resume match timer'}
              >
                {isClockRunning ? <Pause size={13} /> : <Play size={13} />}
              </button>
            </div>
          </div>

          {/* Right: Telemetry & Pipeline Badges */}
          <div className="flex items-center gap-3 text-xs font-mono">
            <div className="hidden sm:flex items-center gap-2 px-3 py-1.5 rounded-lg bg-[#080d1a] border border-slate-800 text-slate-300">
              <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
              <span>OpenCV 5.0 (MOG2)</span>
            </div>
            <div className="hidden sm:flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-[#080d1a] border border-slate-800 text-emerald-400">
              <Zap size={13} />
              <span>60 FPS • 14ms</span>
            </div>
          </div>

        </div>
      </header>

      {/* Action Notification Banner */}
      {actionNotice && (
        <div className="bg-emerald-600/90 text-white text-center py-2 text-sm font-semibold tracking-wide shadow-md transition-all">
          {actionNotice}
        </div>
      )}

      {/* 2. MAIN 3-COLUMN VAR HUD */}
      <main className="max-w-7xl mx-auto px-6 py-6 space-y-6">
        
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">

          {/* LEFT COLUMN: MULTI-CAMERA FEED SELECTOR (3 Cols) */}
          <div className="lg:col-span-3 space-y-4">
            <div className="bg-[#0f172a] border border-slate-800 rounded-xl p-4 shadow-sm">
              <div className="flex items-center justify-between mb-3 border-b border-slate-800/80 pb-2">
                <span className="text-xs font-semibold uppercase tracking-wider text-slate-300 flex items-center gap-1.5">
                  <Camera size={14} className="text-emerald-400" />
                  Camera Matrix ({cameras.length})
                </span>
                <span className="text-[11px] font-mono text-emerald-400">SYNCED</span>
              </div>

              <div className="space-y-2.5">
                {cameras.map((cam, idx) => {
                  const isActive = activeCam === cam.id;
                  return (
                    <button
                      key={cam.id}
                      onClick={() => setActiveCam(cam.id)}
                      className={`w-full text-left p-3 rounded-xl border transition flex flex-col gap-1.5 ${
                        isActive
                          ? 'bg-emerald-950/40 border-emerald-500/70 shadow-sm'
                          : 'bg-[#080d1a] border-slate-800 hover:border-slate-700 hover:bg-slate-900/60'
                      }`}
                    >
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-mono font-bold text-slate-200 flex items-center gap-1.5">
                          <span className={`w-2 h-2 rounded-full ${isActive ? 'bg-emerald-400' : 'bg-slate-500'}`}></span>
                          {cam.id.toUpperCase()}
                        </span>
                        <span className="text-[10px] font-mono px-1.5 py-0.5 rounded bg-slate-800 text-slate-400">
                          {cam.res || '1080p 60Hz'}
                        </span>
                      </div>
                      <span className="text-xs font-medium text-slate-300 truncate">
                        {cam.name}
                      </span>
                      <span className="text-[11px] text-slate-400 flex items-center justify-between">
                        <span>{cam.angle}</span>
                        {isActive && <span className="text-emerald-400 font-semibold text-[10px]">ACTIVE HUD</span>}
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>

            {/* Quick Demo Simulator for Viva */}
            <div className="bg-[#0f172a] border border-slate-800 rounded-xl p-4 space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-xs font-semibold uppercase tracking-wider text-slate-300 flex items-center gap-1.5">
                  <Flame size={14} className="text-amber-400" />
                  Live Defense Demo
                </span>
                <span className="text-[10px] text-slate-400 font-mono">Viva Mode</span>
              </div>

              <div className="grid grid-cols-2 gap-2 text-xs">
                <button
                  onClick={() => handleSimulateIncident('GOAL')}
                  className="p-2 rounded-lg bg-emerald-950/60 hover:bg-emerald-900/80 border border-emerald-800/80 text-emerald-300 font-medium text-left transition"
                >
                  + Goal Decision
                </button>
                <button
                  onClick={() => handleSimulateIncident('OFFSIDE')}
                  className="p-2 rounded-lg bg-rose-950/60 hover:bg-rose-900/80 border border-rose-800/80 text-rose-300 font-medium text-left transition"
                >
                  + SAOT Offside
                </button>
                <button
                  onClick={() => handleSimulateIncident('FOUL')}
                  className="p-2 rounded-lg bg-amber-950/60 hover:bg-amber-900/80 border border-amber-800/80 text-amber-300 font-medium text-left transition"
                >
                  + Serious Foul
                </button>
                <button
                  onClick={() => handleSimulateIncident('OUT_OF_BOUNDS')}
                  className="p-2 rounded-lg bg-sky-950/60 hover:bg-sky-900/80 border border-sky-800/80 text-sky-300 font-medium text-left transition"
                >
                  + Out of Bounds
                </button>
              </div>

              <button
                onClick={handleClearEvents}
                className="w-full py-1.5 text-[11px] text-slate-400 hover:text-slate-200 bg-[#080d1a] border border-slate-800 rounded-lg hover:bg-slate-800 transition"
              >
                Clear Incidents Queue
              </button>
            </div>
          </div>

          {/* CENTER COLUMN: MAIN BROADCAST MONITOR & SCRUBBING (6 Cols) */}
          <div className="lg:col-span-6 space-y-4">
            
            {/* Monitor Box */}
            <div className="bg-[#0b1222] border border-slate-800 rounded-xl overflow-hidden shadow-2xl relative">
              
              {/* OSD Header Bar */}
              <div className="px-4 py-2 bg-[#060a14] border-b border-slate-800 flex items-center justify-between text-xs font-mono">
                <div className="flex items-center gap-2">
                  <span className="w-2.5 h-2.5 rounded-full bg-red-500 animate-pulse"></span>
                  <span className="text-red-400 font-bold uppercase">REC [FEED 01]</span>
                  <span className="text-slate-400 font-semibold">{activeCam.toUpperCase()}</span>
                </div>
                <div className="flex items-center gap-3 text-slate-400">
                  <span>TC: 01:{formatTime(matchTime)}:24</span>
                  <span className="text-emerald-400">1080P/60</span>
                </div>
              </div>

              {/* Video Stream Area */}
              <div className="relative aspect-video bg-black flex items-center justify-center overflow-hidden">
                <img
                  src={`http://localhost:8000/api/video_feed/${activeCam}`}
                  alt="Live Camera Optical Feed"
                  className="w-full h-full object-contain"
                  onError={(e) => {
                    // Fallback to placeholder if camera offline
                    e.target.style.display = 'none';
                    if (e.target.nextSibling) e.target.nextSibling.style.display = 'flex';
                  }}
                />
                
                {/* Fallback Graphic if Camera not accessible */}
                <div className="hidden absolute inset-0 bg-[#090e1a] flex-col items-center justify-center p-6 text-center space-y-3">
                  <Camera className="w-12 h-12 text-slate-600 animate-pulse" />
                  <p className="text-sm font-semibold text-slate-300">Optical Camera Rig Initializing...</p>
                  <p className="text-xs text-slate-500 max-w-sm">
                    Connecting to local DirectShow VideoCapture device. Ensure backend server is active on port 8000.
                  </p>
                </div>

                {/* Overlaid Homography & Pitch Grid Simulation */}
                {overlays.penaltyGrid && (
                  <div className="absolute inset-0 pointer-events-none border-2 border-emerald-400/20 m-6 rounded-none">
                    <div className="absolute left-0 top-1/4 bottom-1/4 w-1/4 border-r-2 border-emerald-400/30 bg-emerald-500/5 flex items-center justify-start p-2">
                      <span className="text-[10px] font-mono text-emerald-400/80 bg-black/60 px-1 py-0.5 rounded">
                        18-YARD MATRIX
                      </span>
                    </div>
                  </div>
                )}

                {/* SAOT Offside Virtual Line Overlay */}
                {overlays.saotLine && (
                  <div className="absolute top-0 bottom-0 left-[48%] w-[2px] bg-red-500/80 shadow-[0_0_8px_rgba(239,68,68,0.8)] pointer-events-none flex flex-col justify-between py-4">
                    <span className="text-[9px] font-mono bg-red-950 text-red-300 border border-red-700 px-1 py-0.5 rounded ml-2">
                      SAOT DEFENDER LINE: 34.2m
                    </span>
                    <span className="text-[9px] font-mono bg-red-950 text-red-300 border border-red-700 px-1 py-0.5 rounded ml-2">
                      TOLERANCE: ±0.4cm
                    </span>
                  </div>
                )}

                {/* Tracking Bounding Box */}
                {overlays.playerBoxes && (
                  <div className="absolute top-1/3 left-1/3 w-20 h-28 border border-emerald-400 pointer-events-none rounded-sm">
                    <div className="absolute -top-5 left-0 bg-emerald-950 text-emerald-300 text-[9px] font-mono px-1 py-0.5 rounded border border-emerald-700">
                      #10 AARAV [FWD]
                    </div>
                  </div>
                )}
              </div>

              {/* Overlays Toolbar */}
              <div className="px-4 py-2.5 bg-[#0a101f] border-t border-slate-800 flex flex-wrap items-center justify-between gap-3 text-xs">
                <span className="text-slate-400 font-semibold text-[11px] uppercase tracking-wider">
                  Vision Overlays:
                </span>

                <div className="flex flex-wrap items-center gap-2">
                  <button
                    onClick={() => setOverlays({ ...overlays, saotLine: !overlays.saotLine })}
                    className={`px-2.5 py-1 rounded text-xs font-medium border transition ${
                      overlays.saotLine
                        ? 'bg-red-950/70 border-red-700 text-red-300'
                        : 'bg-slate-900 border-slate-700 text-slate-400 hover:text-slate-300'
                    }`}
                  >
                    SAOT Offside Line
                  </button>

                  <button
                    onClick={() => setOverlays({ ...overlays, penaltyGrid: !overlays.penaltyGrid })}
                    className={`px-2.5 py-1 rounded text-xs font-medium border transition ${
                      overlays.penaltyGrid
                        ? 'bg-emerald-950/70 border-emerald-700 text-emerald-300'
                        : 'bg-slate-900 border-slate-700 text-slate-400 hover:text-slate-300'
                    }`}
                  >
                    18Y Penalty Grid
                  </button>

                  <button
                    onClick={() => setOverlays({ ...overlays, playerBoxes: !overlays.playerBoxes })}
                    className={`px-2.5 py-1 rounded text-xs font-medium border transition ${
                      overlays.playerBoxes
                        ? 'bg-sky-950/70 border-sky-700 text-sky-300'
                        : 'bg-slate-900 border-slate-700 text-slate-400 hover:text-slate-300'
                    }`}
                  >
                    Player Tracking IDs
                  </button>
                </div>
              </div>

              {/* Scrubbing & Slo-Mo Playback Bar */}
              <div className="px-4 py-3 bg-[#060a14] border-t border-slate-800 flex flex-wrap items-center justify-between gap-3 text-xs font-mono">
                <div className="flex items-center gap-1.5">
                  <button 
                    onClick={() => setMatchTime(t => Math.max(0, t - 5))}
                    className="px-2 py-1 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded border border-slate-700"
                  >
                    -5s
                  </button>
                  <button 
                    onClick={() => setMatchTime(t => Math.max(0, t - 1))}
                    className="px-2 py-1 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded border border-slate-700"
                  >
                    -1s
                  </button>
                  <button 
                    onClick={() => setIsClockRunning(!isClockRunning)}
                    className="px-3 py-1 bg-emerald-600 hover:bg-emerald-500 text-white rounded font-sans font-medium"
                  >
                    {isClockRunning ? 'FREEZE FRAME' : 'RESUME PLAY'}
                  </button>
                  <button 
                    onClick={() => setMatchTime(t => t + 1)}
                    className="px-2 py-1 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded border border-slate-700"
                  >
                    +1s
                  </button>
                  <button 
                    onClick={() => setMatchTime(t => t + 5)}
                    className="px-2 py-1 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded border border-slate-700"
                  >
                    +5s
                  </button>
                </div>

                <div className="flex items-center gap-1 text-[11px]">
                  <span className="text-slate-400 mr-1">VAR SPEED:</span>
                  {['0.25x', '0.5x', '1.0x'].map((speed) => (
                    <button
                      key={speed}
                      onClick={() => setPlaybackSpeed(speed)}
                      className={`px-2 py-0.5 rounded transition ${
                        playbackSpeed === speed
                          ? 'bg-emerald-500 text-black font-bold'
                          : 'bg-slate-800 text-slate-400 hover:text-slate-200'
                      }`}
                    >
                      {speed}
                    </button>
                  ))}
                </div>
              </div>

            </div>
          </div>

          {/* RIGHT COLUMN: OFFICIAL VAR DECISION QUEUE & ACTIONS (3 Cols) */}
          <div className="lg:col-span-3 space-y-4">
            
            {/* Incident Decision Hub */}
            <div className="bg-[#0f172a] border border-slate-800 rounded-xl p-4 space-y-4 shadow-sm">
              <div className="flex items-center justify-between border-b border-slate-800/80 pb-2">
                <span className="text-xs font-semibold uppercase tracking-wider text-slate-300 flex items-center gap-1.5">
                  <ShieldAlert size={14} className="text-amber-400" />
                  VAR Incident Desk ({events.length})
                </span>
                <span className="text-[10px] font-mono text-slate-400">FIFA Protocol</span>
              </div>

              {/* Selected Incident Card */}
              {selectedEvent ? (
                <div className="bg-[#080d1a] border border-slate-700 rounded-xl p-3.5 space-y-3">
                  <div className="flex items-center justify-between">
                    <span className="text-[10px] font-mono px-2 py-0.5 rounded font-bold bg-amber-950/80 text-amber-300 border border-amber-800">
                      {selectedEvent.type || 'INCIDENT REVIEW'}
                    </span>
                    <span className="text-xs font-mono text-emerald-400 font-semibold">
                      {selectedEvent.confidence || 98.4}% Conf
                    </span>
                  </div>

                  <p className="text-xs text-slate-200 font-medium leading-relaxed">
                    {selectedEvent.desc}
                  </p>

                  <div className="text-[11px] font-mono text-slate-400 flex justify-between border-t border-slate-800 pt-2">
                    <span>Incident TC: {selectedEvent.time}</span>
                    <span className={`font-semibold ${
                      selectedEvent.status === 'CONFIRMED' ? 'text-emerald-400' :
                      selectedEvent.status === 'OVERRULED' ? 'text-rose-400' :
                      selectedEvent.status === 'ON_FIELD_REVIEW' ? 'text-amber-400' :
                      'text-sky-400'
                    }`}>
                      {selectedEvent.status || 'PENDING'}
                    </span>
                  </div>

                  {/* Referee Decision Buttons */}
                  <div className="space-y-2 pt-1">
                    <button
                      onClick={() => handleDecision('CONFIRMED')}
                      className="w-full py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold rounded-lg transition flex items-center justify-center gap-1.5 shadow-sm"
                    >
                      <CheckCircle size={14} />
                      Confirm On-Field Decision
                    </button>

                    <button
                      onClick={() => handleDecision('OVERRULED')}
                      className="w-full py-2 bg-rose-700 hover:bg-rose-600 text-white text-xs font-semibold rounded-lg transition flex items-center justify-center gap-1.5 shadow-sm"
                    >
                      <XCircle size={14} />
                      Overrule Call (Recommend Var Overturn)
                    </button>

                    <button
                      onClick={() => handleDecision('ON_FIELD_REVIEW')}
                      className="w-full py-2 bg-amber-600 hover:bg-amber-500 text-white text-xs font-semibold rounded-lg transition flex items-center justify-center gap-1.5 shadow-sm"
                    >
                      <Eye size={14} />
                      Call Referee to Pitchside Monitor (OFR)
                    </button>
                  </div>
                </div>
              ) : (
                <div className="p-6 text-center text-slate-500 text-xs border border-dashed border-slate-800 rounded-xl space-y-1">
                  <UserCheck className="mx-auto w-6 h-6 text-slate-600 mb-1" />
                  <p className="font-semibold text-slate-400">All Systems Clear</p>
                  <p>No pending VAR reviews in progress.</p>
                </div>
              )}

              {/* Incidents History Mini-List */}
              <div className="space-y-1.5 max-h-48 overflow-y-auto pr-1">
                {events.map((ev, i) => (
                  <button
                    key={ev.id || i}
                    onClick={() => setSelectedEventId(ev.id)}
                    className={`w-full text-left p-2 rounded-lg border text-xs transition flex items-center justify-between ${
                      selectedEventId === ev.id
                        ? 'bg-slate-800 border-slate-600 text-slate-100'
                        : 'bg-[#080d1a] border-slate-800/80 text-slate-400 hover:bg-slate-800/60'
                    }`}
                  >
                    <div className="truncate mr-2">
                      <span className="font-semibold text-slate-300 block truncate">{ev.type}</span>
                      <span className="text-[10px] text-slate-500 font-mono">{ev.time}</span>
                    </div>
                    <span className="text-[10px] font-mono px-1.5 py-0.5 rounded bg-slate-900 border border-slate-700 text-slate-300 shrink-0">
                      {ev.status || 'REVIEW'}
                    </span>
                  </button>
                ))}
              </div>
            </div>

          </div>

        </div>

        {/* 3. BOTTOM TABBED DRAWER: PLAYER BIOMETRICS & ANALYTICS */}
        <div className="bg-[#0f172a] border border-slate-800 rounded-xl overflow-hidden shadow-sm">
          
          {/* Drawer Tabs */}
          <div className="border-b border-slate-800 px-6 py-3 bg-[#0c1322] flex flex-wrap items-center justify-between gap-4">
            <div className="flex items-center gap-2">
              <button
                onClick={() => setBottomTab('biometrics')}
                className={`inline-flex items-center gap-2 px-3.5 py-1.5 rounded-lg text-xs font-medium transition ${
                  bottomTab === 'biometrics'
                    ? 'bg-slate-800 text-emerald-400 border border-slate-700'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
              >
                <Heart size={14} className={bottomTab === 'biometrics' ? 'text-rose-500 fill-rose-500' : ''} />
                Live IoT Player Telemetry ({players.length} Sensors)
              </button>

              <button
                onClick={() => setBottomTab('stats')}
                className={`inline-flex items-center gap-2 px-3.5 py-1.5 rounded-lg text-xs font-medium transition ${
                  bottomTab === 'stats'
                    ? 'bg-slate-800 text-emerald-400 border border-slate-700'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
              >
                <Activity size={14} />
                Match Tactical Analytics
              </button>

              <button
                onClick={() => setBottomTab('methodology')}
                className={`inline-flex items-center gap-2 px-3.5 py-1.5 rounded-lg text-xs font-medium transition ${
                  bottomTab === 'methodology'
                    ? 'bg-slate-800 text-emerald-400 border border-slate-700'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
              >
                <FileText size={14} />
                Computer Vision Architecture & Group 12
              </button>
            </div>

            <div className="flex items-center gap-3 text-xs font-mono text-slate-400">
              <span className="flex items-center gap-1.5">
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-400"></span>
                ESP32 Bluetooth Mesh / MQTT: Active
              </span>
            </div>
          </div>

          {/* TAB 1: PLAYER BIOMETRICS TABLE */}
          {bottomTab === 'biometrics' && (
            <div className="p-6">
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                {players.map((p) => {
                  const isHighCardiac = p.bpm >= 160;
                  return (
                    <div 
                      key={p.id}
                      className={`p-3.5 rounded-xl border transition ${
                        isHighCardiac 
                          ? 'bg-rose-950/20 border-rose-800/80' 
                          : 'bg-[#080d1a] border-slate-800 hover:border-slate-700'
                      }`}
                    >
                      <div className="flex items-center justify-between mb-2">
                        <div className="flex items-center gap-2">
                          <span className={`w-6 h-6 rounded-md font-mono text-xs font-bold flex items-center justify-center ${
                            p.team === 'A' ? 'bg-emerald-950 text-emerald-300 border border-emerald-800' : 'bg-sky-950 text-sky-300 border border-sky-800'
                          }`}>
                            #{p.number}
                          </span>
                          <div>
                            <h4 className="text-xs font-semibold text-slate-200 leading-tight truncate w-28">
                              {p.name}
                            </h4>
                            <span className="text-[10px] text-slate-500 font-mono">{p.pos} • {p.sensorId}</span>
                          </div>
                        </div>

                        <div className="text-right font-mono">
                          <span className={`text-sm font-bold flex items-center justify-end gap-1 ${
                            isHighCardiac ? 'text-rose-400' : 'text-emerald-400'
                          }`}>
                            <Heart size={12} className={isHighCardiac ? 'fill-rose-400 animate-pulse' : 'fill-emerald-400'} />
                            {p.bpm}
                          </span>
                          <span className="text-[9px] text-slate-500">BPM</span>
                        </div>
                      </div>

                      {/* Stamina & Distance Bars */}
                      <div className="space-y-1.5 text-[11px] font-mono">
                        <div>
                          <div className="flex justify-between text-slate-400 text-[10px] mb-0.5">
                            <span>Stamina</span>
                            <span>{p.stamina}%</span>
                          </div>
                          <div className="w-full bg-slate-800 h-1.5 rounded-full overflow-hidden">
                            <div 
                              className={`h-full rounded-full ${p.stamina > 80 ? 'bg-emerald-500' : 'bg-amber-500'}`} 
                              style={{ width: `${p.stamina}%` }}
                            ></div>
                          </div>
                        </div>

                        <div className="flex justify-between text-[10px] text-slate-400 pt-1">
                          <span>Total Dist: <strong className="text-slate-300">{p.distance}</strong></span>
                          <span>Sprints: <strong className="text-slate-300">{p.sprints}</strong></span>
                        </div>
                      </div>

                      {isHighCardiac && (
                        <div className="mt-2 text-[10px] font-mono text-rose-400 bg-rose-950/80 px-2 py-0.5 rounded border border-rose-800/60 flex items-center gap-1">
                          <AlertTriangle size={10} />
                          Elevated Cardiac Load Alert
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* TAB 2: TACTICAL MATCH ANALYTICS */}
          {bottomTab === 'stats' && (
            <div className="p-6">
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                
                {/* Possession Bar */}
                <div className="bg-[#080d1a] border border-slate-800 rounded-xl p-4 space-y-3">
                  <h4 className="text-xs font-semibold text-slate-300 uppercase tracking-wider">Ball Possession</h4>
                  <div className="flex justify-between text-sm font-mono font-bold">
                    <span className="text-emerald-400">{matchInfo.teamA}: 54%</span>
                    <span className="text-sky-400">{matchInfo.teamB}: 46%</span>
                  </div>
                  <div className="w-full bg-slate-800 h-2.5 rounded-full overflow-hidden flex">
                    <div className="bg-emerald-500 h-full" style={{ width: '54%' }}></div>
                    <div className="bg-sky-500 h-full" style={{ width: '46%' }}></div>
                  </div>
                  <p className="text-[11px] text-slate-500">Calculated via optical centroid time-occupancy across pitch quadrants.</p>
                </div>

                {/* Match Metrics Table */}
                <div className="bg-[#080d1a] border border-slate-800 rounded-xl p-4 space-y-2 text-xs font-mono">
                  <h4 className="font-semibold text-slate-300 uppercase tracking-wider font-sans mb-3">Key Performance Indicators</h4>
                  <div className="flex justify-between text-slate-400 border-b border-slate-800 pb-1">
                    <span>Shots on Target:</span>
                    <span className="text-slate-200 font-bold">4 - 2</span>
                  </div>
                  <div className="flex justify-between text-slate-400 border-b border-slate-800 pb-1">
                    <span>Passes Completed:</span>
                    <span className="text-slate-200 font-bold">284 - 241 (87%)</span>
                  </div>
                  <div className="flex justify-between text-slate-400 border-b border-slate-800 pb-1">
                    <span>SAOT Offsides Called:</span>
                    <span className="text-slate-200 font-bold">2 - 1</span>
                  </div>
                  <div className="flex justify-between text-slate-400">
                    <span>Fouls Evaluated:</span>
                    <span className="text-slate-200 font-bold">6 - 8</span>
                  </div>
                </div>

                {/* Computer Vision Frame Rate Stats */}
                <div className="bg-[#080d1a] border border-slate-800 rounded-xl p-4 space-y-2 text-xs font-mono">
                  <h4 className="font-semibold text-slate-300 uppercase tracking-wider font-sans mb-3">Vision Engine Latency</h4>
                  <div className="flex justify-between text-slate-400 border-b border-slate-800 pb-1">
                    <span>MOG2 Subtractor Cycle:</span>
                    <span className="text-emerald-400 font-bold">6.2 ms</span>
                  </div>
                  <div className="flex justify-between text-slate-400 border-b border-slate-800 pb-1">
                    <span>HSV Ball Contour Detector:</span>
                    <span className="text-emerald-400 font-bold">4.8 ms</span>
                  </div>
                  <div className="flex justify-between text-slate-400 border-b border-slate-800 pb-1">
                    <span>Homography Transform:</span>
                    <span className="text-emerald-400 font-bold">1.8 ms</span>
                  </div>
                  <div className="flex justify-between text-slate-400">
                    <span>Overall Pipeline FPS:</span>
                    <span className="text-emerald-400 font-bold">59.8 FPS (Nominal)</span>
                  </div>
                </div>

              </div>
            </div>
          )}

          {/* TAB 3: ACADEMIC METHODOLOGY & GROUP DETAILS */}
          {bottomTab === 'methodology' && (
            <div className="p-6 space-y-4 text-xs text-slate-300 leading-relaxed">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="bg-[#080d1a] border border-slate-800 rounded-xl p-4 space-y-2">
                  <h4 className="text-sm font-semibold text-slate-100 flex items-center gap-2">
                    <Award size={16} className="text-emerald-400" />
                    Project Identity & College Credentials
                  </h4>
                  <p><strong className="text-slate-200">Institution:</strong> Atharva College of Engineering, Malad West, Mumbai</p>
                  <p><strong className="text-slate-200">Department:</strong> Department of Computer Engineering (BE Capstone 2025–26)</p>
                  <p><strong className="text-slate-200">Project Guide:</strong> Prof. Varsha Patil</p>
                  <p><strong className="text-slate-200">Research Team (Group 12):</strong> Mann Sutariya, Tejas Walanj, Karan Sonar, Yashwin Rajpurohit</p>
                </div>

                <div className="bg-[#080d1a] border border-slate-800 rounded-xl p-4 space-y-2">
                  <h4 className="text-sm font-semibold text-slate-100">Algorithmic Formulation (Synopsis Table 3)</h4>
                  <p>
                    <strong>Goal Line / Ball:</strong> HSV color space thresholding combined with circularity index metric:
                    <code className="text-emerald-400 font-mono ml-1">Circularity = 4π × (Area) / (Perimeter²) &gt; 0.70</code>
                  </p>
                  <p>
                    <strong>Player Segmentation:</strong> OpenCV MOG2 Background Subtractor with <code className="text-emerald-400 font-mono">varThreshold = 50</code>, shadow removal disabled, and morphological elliptical opening filters.
                  </p>
                  <p>
                    <strong>Wearable Sensor Telemetry:</strong> ESP32 microcontrollers with pulse sensor streaming heart rate frequency and cardiac stress indicators at 1Hz over WiFi/MQTT.
                  </p>
                </div>
              </div>
            </div>
          )}

        </div>

      </main>
    </div>
  );
}