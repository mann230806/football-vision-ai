import React from 'react';
import { BrowserRouter as Router, Routes, Route, Navigate } from 'react-router-dom';
import Setup from './Setup';
import Match from './Match';

export default function App() {
  return (
    <Router>
      <Routes>
        <Route path="/" element={<Navigate to="/setup" />} />
        <Route path="/setup" element={<Setup />} />
        <Route path="/match" element={<Match />} />
      </Routes>
    </Router>
  );
}
