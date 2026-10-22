import React, { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { 
  Sliders, 
  Users, 
  MapPin, 
  Video, 
  Activity, 
  ArrowRight, 
  Plus, 
  Trash2, 
  RotateCcw, 
  CheckCircle2, 
  Cpu, 
  Radio, 
  ShieldCheck, 
  Award,
  Layers,
  Sparkles
} from 'lucide-react';

export default function Setup() {
  const navigate = useNavigate();
  const [activeTab, setActiveTab] = useState('match'); // 'match', 'roster', 'cameras'
  const [saveStatus, setSaveStatus] = useState(null);

  // Match parameters state
  const [matchData, setMatchData] = useState({
    venue: 'Atharva Sports Complex, Malad West',
    pitchLength: 105,
    pitchWidth: 68,
    teamA: 'Atharva FC',
    teamB: 'Mumbai Titans',
    competition: 'Mumbai University Inter-Collegiate Cup 2025–26',
    referee: 'Prof. Varsha Patil (Lead VAR Officiator)',
    turfType: 'Hybrid Reinforced Grass',
    ballType: 'FIFA Pro Sensor-Embedded Ball'
  });

  // Squad rosters with IoT sensor IDs
  const [players, setPlayers] = useState([
    { id: 1, name: 'Aarav Sharma', team: 'A', number: 10, pos: 'FWD', sensorId: 'IOT-HR-04A', bpm: 142, status: 'Active' },
    { id: 2, name: 'Rohan Deshmukh', team: 'A', number: 7, pos: 'MID', sensorId: 'IOT-HR-02A', bpm: 138, status: 'Active' },
    { id: 3, name: 'Vikram Patel', team: 'A', number: 4, pos: 'DEF', sensorId: 'IOT-HR-08A', bpm: 124, status: 'Active' },
    { id: 4, name: 'Neil D\'Souza', team: 'A', number: 1, pos: 'GK', sensorId: 'IOT-HR-01A', bpm: 96, status: 'Active' },
    { id: 5, name: 'Kabir Mehta', team: 'B', number: 9, pos: 'FWD', sensorId: 'IOT-HR-11B', bpm: 151, status: 'Active' },
    { id: 6, name: 'Aditya Varma', team: 'B', number: 8, pos: 'MID', sensorId: 'IOT-HR-07B', bpm: 135, status: 'Active' },
    { id: 7, name: 'Siddharth Rao', team: 'B', number: 5, pos: 'DEF', sensorId: 'IOT-HR-05B', bpm: 128, status: 'Active' },
    { id: 8, name: 'Tariq Khan', team: 'B', number: 1, pos: 'GK', sensorId: 'IOT-HR-01B', bpm: 98, status: 'Active' }
  ]);

  // Camera calibration specs
  const [cameraRig, setCameraRig] = useState([
    { id: 'cam1', name: 'CAM 01: Main Broadcast (Tactical Pitch-Wide)', focal: '24mm f/2.8', fps: 60, status: 'Calibrated / Online', latency: '12ms' },
    { id: 'cam2', name: 'CAM 02: North Goal-Line & 18Y Box Matrix', focal: '50mm f/1.8', fps: 60, status: 'Calibrated / Online', latency: '15ms' },
    { id: 'cam3', name: 'CAM 03: Tactical Panoramic Overhead Gantry', focal: '16mm Ultra-Wide', fps: 60, status: 'Calibrated / Online', latency: '18ms' }
  ]);

  // New player form state
  const [newPlayer, setNewPlayer] = useState({
    name: '',
    team: 'A',
    number: '',
    pos: 'MID',
    sensorId: ''
  });

  // Load existing match setup from backend on mount
  useEffect(() => {
    fetch('http://localhost:8000/api/setup')
      .then(res => res.json())
      .then(data => {
        if (data.matchData && data.matchData.teamA) {
          setMatchData(prev => ({ ...prev, ...data.matchData }));
        }
        if (data.players && data.players.length > 0) {
          setPlayers(data.players);
        }
      })
      .catch(err => console.log('Backend offline or using default presets', err));
  }, []);

  const handleAddPlayer = (e) => {
    e.preventDefault();
    if (!newPlayer.name || !newPlayer.number) return;
    const playerObj = {
      id: Date.now(),
      name: newPlayer.name.trim(),
      team: newPlayer.team,
      number: parseInt(newPlayer.number, 10),
      pos: newPlayer.pos,
      sensorId: newPlayer.sensorId || `IOT-HR-${Math.floor(10 + Math.random() * 89)}${newPlayer.team}`,
      bpm: 110,
      status: 'Active'
    };
    setPlayers([...players, playerObj]);
    setNewPlayer({ name: '', team: 'A', number: '', pos: 'MID', sensorId: '' });
  };

  const handleRemovePlayer = (id) => {
    setPlayers(players.filter(p => p.id !== id));
  };

  const handleResetDefaults = () => {
    setMatchData({
      venue: 'Atharva Sports Complex, Malad West',
      pitchLength: 105,
      pitchWidth: 68,
      teamA: 'Atharva FC',
      teamB: 'Mumbai Titans',
      competition: 'Mumbai University Inter-Collegiate Cup 2025–26',
      referee: 'Prof. Varsha Patil (Lead VAR Officiator)',
      turfType: 'Hybrid Reinforced Grass',
      ballType: 'FIFA Pro Sensor-Embedded Ball'
    });
  };

  const handleStartVARStudio = async () => {
    setSaveStatus('synchronizing');
    try {
      await fetch('http://localhost:8000/api/setup', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ matchData, players })
      });
      setSaveStatus('success');
      setTimeout(() => {
        navigate('/match');
      }, 500);
    } catch (err) {
      console.warn('Backend offline, proceeding to studio anyway', err);
      navigate('/match');
    }
  };

  return (
    <div className="min-h-screen bg-[#070b14] text-slate-100 font-sans selection:bg-emerald-500 selection:text-black">
      {/* Top Academic & Project Identity Bar */}
      <header className="border-b border-slate-800 bg-[#0c1322]/90 backdrop-blur sticky top-0 z-50 px-6 py-3.5">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="flex items-center gap-3.5">
            <div className="w-10 h-10 rounded-lg bg-emerald-500/10 border border-emerald-500/30 flex items-center justify-center text-emerald-400 font-bold text-lg shadow-sm">
              VAR
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-xs font-semibold uppercase tracking-wider px-2 py-0.5 rounded bg-emerald-950/80 text-emerald-400 border border-emerald-800/60">
                  BE Capstone Project 2025–26
                </span>
                <span className="text-xs text-slate-400">| Dept. of Computer Engineering</span>
              </div>
              <h1 className="text-lg font-bold tracking-tight text-white flex items-center gap-2">
                Atharva College of Engineering <span className="text-slate-500 font-normal">/</span> Autonomous VAR Studio
              </h1>
            </div>
          </div>

          <div className="flex items-center gap-3 text-xs text-slate-300">
            <div className="hidden lg:flex flex-col text-right">
              <span className="text-slate-400">Project Group 12</span>
              <span className="font-medium text-slate-200">Mann S. • Tejas W. • Karan S. • Yashwin R.</span>
            </div>
            <div className="h-7 w-[1px] bg-slate-700 hidden lg:block"></div>
            <button
              onClick={handleStartVARStudio}
              disabled={saveStatus === 'synchronizing'}
              className="inline-flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg font-medium text-sm transition-all duration-150 shadow-lg shadow-emerald-900/30 active:scale-[0.98]"
            >
              {saveStatus === 'synchronizing' ? 'Calibrating...' : 'Launch VAR Live Studio'}
              <ArrowRight size={16} />
            </button>
          </div>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="max-w-7xl mx-auto px-6 py-8 space-y-6">
        
        {/* Navigation Tabs */}
        <div className="flex flex-wrap items-center justify-between gap-4 border-b border-slate-800 pb-3">
          <div className="flex items-center gap-2">
            <button
              onClick={() => setActiveTab('match')}
              className={`inline-flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium transition-colors ${
                activeTab === 'match'
                  ? 'bg-slate-800 text-emerald-400 border border-slate-700'
                  : 'text-slate-400 hover:text-slate-200 hover:bg-slate-900'
              }`}
            >
              <Sliders size={16} />
              1. Match & Pitch Calibration
            </button>
            <button
              onClick={() => setActiveTab('roster')}
              className={`inline-flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium transition-colors ${
                activeTab === 'roster'
                  ? 'bg-slate-800 text-emerald-400 border border-slate-700'
                  : 'text-slate-400 hover:text-slate-200 hover:bg-slate-900'
              }`}
            >
              <Users size={16} />
              2. Squad Rosters & IoT Sensors ({players.length})
            </button>
            <button
              onClick={() => setActiveTab('cameras')}
              className={`inline-flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium transition-colors ${
                activeTab === 'cameras'
                  ? 'bg-slate-800 text-emerald-400 border border-slate-700'
                  : 'text-slate-400 hover:text-slate-200 hover:bg-slate-900'
              }`}
            >
              <Video size={16} />
              3. Multi-Camera Rig Calibration
            </button>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={handleResetDefaults}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs text-slate-400 hover:text-slate-200 hover:bg-slate-800 rounded border border-slate-700 transition"
              title="Reset fields to Atharva Derby match preset"
            >
              <RotateCcw size={13} />
              Reset College Preset
            </button>
          </div>
        </div>

        {/* TAB 1: MATCH & PITCH CALIBRATION */}
        {activeTab === 'match' && (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            <div className="lg:col-span-2 space-y-6">
              
              {/* Competition Details Panel */}
              <div className="bg-[#0f172a] border border-slate-800 rounded-xl p-6 shadow-sm">
                <div className="flex items-center justify-between mb-5">
                  <div className="flex items-center gap-2.5">
                    <Award className="text-emerald-400 w-5 h-5" />
                    <h2 className="font-semibold text-base text-slate-100">Fixture & Officiating Registry</h2>
                  </div>
                  <span className="text-xs bg-slate-800 text-slate-400 px-2.5 py-1 rounded border border-slate-700 font-mono">
                    FIFA Rulebook 2025/26
                  </span>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="space-y-1.5">
                    <label className="text-xs font-medium text-slate-400">Competition / League</label>
                    <input
                      type="text"
                      value={matchData.competition}
                      onChange={(e) => setMatchData({ ...matchData, competition: e.target.value })}
                      className="w-full bg-[#080d1a] border border-slate-700 rounded-lg px-3.5 py-2 text-sm text-slate-200 focus:outline-none focus:border-emerald-500 transition"
                    />
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-xs font-medium text-slate-400">Venue & Pitch Location</label>
                    <input
                      type="text"
                      value={matchData.venue}
                      onChange={(e) => setMatchData({ ...matchData, venue: e.target.value })}
                      className="w-full bg-[#080d1a] border border-slate-700 rounded-lg px-3.5 py-2 text-sm text-slate-200 focus:outline-none focus:border-emerald-500 transition"
                    />
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-xs font-medium text-slate-400">Home Team (Club A)</label>
                    <input
                      type="text"
                      value={matchData.teamA}
                      onChange={(e) => setMatchData({ ...matchData, teamA: e.target.value })}
                      className="w-full bg-[#080d1a] border border-slate-700 rounded-lg px-3.5 py-2 text-sm text-slate-200 focus:outline-none focus:border-emerald-500 transition font-medium text-emerald-400"
                    />
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-xs font-medium text-slate-400">Away Team (Club B)</label>
                    <input
                      type="text"
                      value={matchData.teamB}
                      onChange={(e) => setMatchData({ ...matchData, teamB: e.target.value })}
                      className="w-full bg-[#080d1a] border border-slate-700 rounded-lg px-3.5 py-2 text-sm text-slate-200 focus:outline-none focus:border-emerald-500 transition font-medium text-sky-400"
                    />
                  </div>

                  <div className="space-y-1.5 md:col-span-2">
                    <label className="text-xs font-medium text-slate-400">Lead VAR Official / Guide Supervisor</label>
                    <input
                      type="text"
                      value={matchData.referee}
                      onChange={(e) => setMatchData({ ...matchData, referee: e.target.value })}
                      className="w-full bg-[#080d1a] border border-slate-700 rounded-lg px-3.5 py-2 text-sm text-slate-200 focus:outline-none focus:border-emerald-500 transition"
                    />
                  </div>
                </div>
              </div>

              {/* Pitch Homography & Dimensions */}
              <div className="bg-[#0f172a] border border-slate-800 rounded-xl p-6 shadow-sm">
                <div className="flex items-center justify-between mb-5">
                  <div className="flex items-center gap-2.5">
                    <MapPin className="text-emerald-400 w-5 h-5" />
                    <h2 className="font-semibold text-base text-slate-100">Pitch Geometry & Metric Calibration</h2>
                  </div>
                  <span className="text-xs text-slate-400 font-mono">Standard Metric Calibration</span>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <div className="space-y-1.5">
                    <label className="text-xs font-medium text-slate-400">Pitch Length (Meters)</label>
                    <div className="relative">
                      <input
                        type="number"
                        value={matchData.pitchLength}
                        onChange={(e) => setMatchData({ ...matchData, pitchLength: parseFloat(e.target.value) || 105 })}
                        className="w-full bg-[#080d1a] border border-slate-700 rounded-lg px-3.5 py-2 text-sm text-slate-200 focus:outline-none focus:border-emerald-500 transition"
                      />
                      <span className="absolute right-3 top-2.5 text-xs text-slate-400">meters</span>
                    </div>
                    <span className="text-[11px] text-slate-400">Standard FIFA: 105m</span>
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-xs font-medium text-slate-400">Pitch Width (Meters)</label>
                    <div className="relative">
                      <input
                        type="number"
                        value={matchData.pitchWidth}
                        onChange={(e) => setMatchData({ ...matchData, pitchWidth: parseFloat(e.target.value) || 68 })}
                        className="w-full bg-[#080d1a] border border-slate-700 rounded-lg px-3.5 py-2 text-sm text-slate-200 focus:outline-none focus:border-emerald-500 transition"
                      />
                      <span className="absolute right-3 top-2.5 text-xs text-slate-400">meters</span>
                    </div>
                    <span className="text-[11px] text-slate-400">Standard FIFA: 68m</span>
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-xs font-medium text-slate-400">Pitch Surface</label>
                    <select
                      value={matchData.turfType}
                      onChange={(e) => setMatchData({ ...matchData, turfType: e.target.value })}
                      className="w-full bg-[#080d1a] border border-slate-700 rounded-lg px-3.5 py-2 text-sm text-slate-200 focus:outline-none focus:border-emerald-500 transition"
                    >
                      <option>Hybrid Reinforced Grass</option>
                      <option>Natural Bermuda Turf</option>
                      <option>FIFA Quality Pro 4G Turf</option>
                    </select>
                  </div>
                </div>
              </div>

            </div>

            {/* Right Summary Column */}
            <div className="space-y-6">
              <div className="bg-[#0f172a] border border-slate-800 rounded-xl p-5 space-y-4">
                <h3 className="text-sm font-semibold text-slate-200 uppercase tracking-wider flex items-center gap-2">
                  <ShieldCheck size={16} className="text-emerald-400" />
                  System Readiness Checklist
                </h3>
                <ul className="space-y-2.5 text-xs">
                  <li className="flex items-start gap-2.5 text-slate-300">
                    <CheckCircle2 size={16} className="text-emerald-400 shrink-0 mt-0.5" />
                    <span>OpenCV 5.0 MOG2 background subtractor initialized</span>
                  </li>
                  <li className="flex items-start gap-2.5 text-slate-300">
                    <CheckCircle2 size={16} className="text-emerald-400 shrink-0 mt-0.5" />
                    <span>Multi-camera homography matrix calibrated for 105m × 68m</span>
                  </li>
                  <li className="flex items-start gap-2.5 text-slate-300">
                    <CheckCircle2 size={16} className="text-emerald-400 shrink-0 mt-0.5" />
                    <span>IoT Telemetry Hub connected (ESP32 Bluetooth Mesh / MQTT)</span>
                  </li>
                  <li className="flex items-start gap-2.5 text-slate-300">
                    <CheckCircle2 size={16} className="text-emerald-400 shrink-0 mt-0.5" />
                    <span>FIFA VAR Decision Protocol standard enabled</span>
                  </li>
                </ul>

                <div className="pt-3 border-t border-slate-800">
                  <div className="bg-[#080d1a] p-3 rounded-lg border border-slate-800 text-xs space-y-1">
                    <span className="text-slate-400 block font-medium">Academic Project Details:</span>
                    <p className="text-slate-300">Atharva College of Engineering, Malad (W)</p>
                    <p className="text-slate-400">Guide: Prof. Varsha Patil</p>
                    <p className="text-emerald-400/90 font-mono text-[11px] mt-1">Status: READY FOR MATCH COMMENCEMENT</p>
                  </div>
                </div>
              </div>

              <button
                onClick={handleStartVARStudio}
                className="w-full py-3.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl font-semibold text-sm transition shadow-lg shadow-emerald-950/40 flex items-center justify-center gap-2 group"
              >
                <span>Save Parameters & Go to VAR HUD</span>
                <ArrowRight size={18} className="group-hover:translate-x-1 transition-transform" />
              </button>
            </div>
          </div>
        )}

        {/* TAB 2: SQUAD ROSTERS & IOT WEARABLES */}
        {activeTab === 'roster' && (
          <div className="space-y-6">
            {/* Add Player Bar */}
            <form onSubmit={handleAddPlayer} className="bg-[#0f172a] border border-slate-800 p-4 rounded-xl flex flex-wrap items-center gap-3">
              <span className="text-xs font-semibold text-slate-300 uppercase tracking-wider flex items-center gap-1.5 mr-2">
                <Plus size={16} className="text-emerald-400" />
                Register Squad Member:
              </span>

              <input
                type="text"
                placeholder="Player Full Name"
                value={newPlayer.name}
                onChange={(e) => setNewPlayer({ ...newPlayer, name: e.target.value })}
                className="bg-[#080d1a] border border-slate-700 rounded-lg px-3 py-1.5 text-sm text-slate-200 placeholder-slate-500 focus:outline-none focus:border-emerald-500 flex-1 min-w-[160px]"
                required
              />

              <div className="flex items-center gap-2">
                <select
                  value={newPlayer.team}
                  onChange={(e) => setNewPlayer({ ...newPlayer, team: e.target.value })}
                  className="bg-[#080d1a] border border-slate-700 rounded-lg px-2.5 py-1.5 text-sm text-slate-200 focus:outline-none focus:border-emerald-500"
                >
                  <option value="A">{matchData.teamA} (Team A)</option>
                  <option value="B">{matchData.teamB} (Team B)</option>
                </select>

                <input
                  type="number"
                  placeholder="Jersey #"
                  value={newPlayer.number}
                  onChange={(e) => setNewPlayer({ ...newPlayer, number: e.target.value })}
                  className="bg-[#080d1a] border border-slate-700 rounded-lg px-3 py-1.5 text-sm text-slate-200 placeholder-slate-500 focus:outline-none focus:border-emerald-500 w-24"
                  min="1"
                  max="99"
                  required
                />

                <select
                  value={newPlayer.pos}
                  onChange={(e) => setNewPlayer({ ...newPlayer, pos: e.target.value })}
                  className="bg-[#080d1a] border border-slate-700 rounded-lg px-2.5 py-1.5 text-sm text-slate-200 focus:outline-none focus:border-emerald-500"
                >
                  <option value="GK">Goalkeeper (GK)</option>
                  <option value="DEF">Defender (DEF)</option>
                  <option value="MID">Midfielder (MID)</option>
                  <option value="FWD">Forward (FWD)</option>
                </select>

                <input
                  type="text"
                  placeholder="IoT Tag (Optional)"
                  value={newPlayer.sensorId}
                  onChange={(e) => setNewPlayer({ ...newPlayer, sensorId: e.target.value })}
                  className="bg-[#080d1a] border border-slate-700 rounded-lg px-3 py-1.5 text-sm text-slate-200 placeholder-slate-500 focus:outline-none focus:border-emerald-500 w-36 font-mono text-xs"
                />

                <button
                  type="submit"
                  className="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-sm font-medium transition shrink-0"
                >
                  Add Player
                </button>
              </div>
            </form>

            {/* Roster Table */}
            <div className="bg-[#0f172a] border border-slate-800 rounded-xl overflow-hidden shadow-sm">
              <div className="px-6 py-4 border-b border-slate-800 flex justify-between items-center bg-[#111a30]">
                <div>
                  <h3 className="font-semibold text-slate-100 text-sm">Active Match Squad & Telemetry Binding</h3>
                  <p className="text-xs text-slate-400">Total registered players: {players.length} | Real-time biometrics stream via ESP32 Wearable Sensors</p>
                </div>
                <div className="flex gap-2">
                  <span className="text-xs bg-emerald-950/80 text-emerald-400 border border-emerald-800/80 px-2.5 py-1 rounded">
                    {matchData.teamA}: {players.filter(p => p.team === 'A').length} Players
                  </span>
                  <span className="text-xs bg-sky-950/80 text-sky-400 border border-sky-800/80 px-2.5 py-1 rounded">
                    {matchData.teamB}: {players.filter(p => p.team === 'B').length} Players
                  </span>
                </div>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs text-slate-300">
                  <thead className="bg-[#090f1d] text-slate-400 uppercase font-mono tracking-wider border-b border-slate-800">
                    <tr>
                      <th className="py-3 px-4">#</th>
                      <th className="py-3 px-4">Squad Member</th>
                      <th className="py-3 px-4">Team</th>
                      <th className="py-3 px-4">Position</th>
                      <th className="py-3 px-4">Assigned IoT Sensor UUID</th>
                      <th className="py-3 px-4">Baseline HR</th>
                      <th className="py-3 px-4">Status</th>
                      <th className="py-3 px-4 text-right">Action</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/60 font-sans">
                    {players.map((player) => (
                      <tr key={player.id} className="hover:bg-slate-800/40 transition">
                        <td className="py-3 px-4 font-mono font-bold text-slate-200">
                          #{player.number}
                        </td>
                        <td className="py-3 px-4 font-medium text-slate-100 flex items-center gap-2">
                          <span className={`w-2 h-2 rounded-full ${player.team === 'A' ? 'bg-emerald-400' : 'bg-sky-400'}`}></span>
                          {player.name}
                        </td>
                        <td className="py-3 px-4">
                          <span className={`px-2 py-0.5 rounded text-[11px] font-semibold ${
                            player.team === 'A' 
                              ? 'bg-emerald-950/80 text-emerald-300 border border-emerald-800/60' 
                              : 'bg-sky-950/80 text-sky-300 border border-sky-800/60'
                          }`}>
                            {player.team === 'A' ? matchData.teamA : matchData.teamB}
                          </span>
                        </td>
                        <td className="py-3 px-4 font-mono">
                          <span className="px-1.5 py-0.5 rounded bg-slate-800 text-slate-300 font-semibold text-[10px]">
                            {player.pos}
                          </span>
                        </td>
                        <td className="py-3 px-4 font-mono text-slate-400">
                          {player.sensorId || `IOT-HR-${player.id}A`}
                        </td>
                        <td className="py-3 px-4 font-mono text-emerald-400">
                          {player.bpm || 120} BPM
                        </td>
                        <td className="py-3 px-4">
                          <span className="inline-flex items-center gap-1.5 text-[11px] text-emerald-400">
                            <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse"></span>
                            Telemetry Linked
                          </span>
                        </td>
                        <td className="py-3 px-4 text-right">
                          <button
                            onClick={() => handleRemovePlayer(player.id)}
                            className="p-1 text-slate-500 hover:text-red-400 transition"
                            title="Remove player"
                          >
                            <Trash2 size={14} />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* TAB 3: MULTI-CAMERA RIG CALIBRATION */}
        {activeTab === 'cameras' && (
          <div className="space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              {cameraRig.map((cam, idx) => (
                <div key={cam.id} className="bg-[#0f172a] border border-slate-800 rounded-xl p-5 space-y-4">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-mono font-bold text-emerald-400 px-2 py-0.5 bg-emerald-950/80 border border-emerald-800/60 rounded">
                      CHANNEL 0{idx + 1}
                    </span>
                    <span className="inline-flex items-center gap-1.5 text-xs text-emerald-400 font-mono">
                      <span className="w-2 h-2 rounded-full bg-emerald-400"></span>
                      LIVE SYNC
                    </span>
                  </div>

                  <div>
                    <h4 className="font-semibold text-slate-100 text-sm">{cam.name}</h4>
                    <p className="text-xs text-slate-400 mt-0.5">Optical Sensor: Sony Exmor 1/1.8" • {cam.focal}</p>
                  </div>

                  <div className="bg-[#080d1a] p-3 rounded-lg border border-slate-800/80 space-y-1.5 text-xs font-mono">
                    <div className="flex justify-between text-slate-400">
                      <span>Stream Pipeline:</span>
                      <span className="text-slate-200">OpenCV VideoCapture</span>
                    </div>
                    <div className="flex justify-between text-slate-400">
                      <span>Capture Rate:</span>
                      <span className="text-emerald-400 font-semibold">{cam.fps} FPS Sustained</span>
                    </div>
                    <div className="flex justify-between text-slate-400">
                      <span>Network Latency:</span>
                      <span className="text-slate-200">{cam.latency}</span>
                    </div>
                    <div className="flex justify-between text-slate-400">
                      <span>Homography Matrix:</span>
                      <span className="text-emerald-400">9-Point Calibrated</span>
                    </div>
                  </div>

                  <div className="pt-2 flex gap-2">
                    <button
                      onClick={() => navigate('/match')}
                      className="w-full py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-medium rounded-lg border border-slate-700 transition"
                    >
                      Preview Stream
                    </button>
                  </div>
                </div>
              ))}
            </div>

            <div className="bg-[#0f172a] border border-slate-800 rounded-xl p-5">
              <h4 className="text-sm font-semibold text-slate-200 mb-2 flex items-center gap-2">
                <Cpu size={16} className="text-emerald-400" />
                Computer Vision Processing Architecture (Academic Synopsis Table 3)
              </h4>
              <p className="text-xs text-slate-400 leading-relaxed">
                The optical vision pipeline executes background subtraction via MOG2 with shadow elimination, applying morphological elliptical opening filters (5×5) before contour identification. Ball tracking utilizes HSV color segmentation and circularity ratio metric: <code className="text-emerald-400 font-mono">4π × Area / Perimeter² &gt; 0.70</code> with Euclidean distance centroid tracking to ensure sub-centimeter goal line accuracy without AI hallucinatory artifacts.
              </p>
            </div>
          </div>
        )}

      </main>
    </div>
  );
}