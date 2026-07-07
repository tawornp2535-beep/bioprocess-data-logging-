import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend
} from 'recharts';
import {
  Activity, Droplets, Wind, Thermometer, RotateCw, PlusCircle,
  Download, LayoutDashboard, LineChart as ChartIcon, FolderPlus, Trash2, FolderOpen,
  Table as TableIcon, Users, Activity as ActivityIcon, Edit3,
  ChevronDown, ChevronUp, ChevronRight, Settings, LogOut, Cpu, Database, Folder,
  Menu, X, Check, Play, Pause, RotateCcw, Star, MessageSquare
} from 'lucide-react';
import './index.css';
import './form.css';
import './tabs.css';
import './table.css';

const compressImage = (file, maxWidth = 400, maxHeight = 400, quality = 0.75) => {
  return new Promise((resolve) => {
    if (!file) {
      resolve(null);
      return;
    }
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = (event) => {
      const img = new Image();
      img.src = event.target.result;
      img.onload = () => {
        const canvas = document.createElement('canvas');
        let width = img.width;
        let height = img.height;

        if (width > height) {
          if (width > maxWidth) {
            height = Math.round((height * maxWidth) / width);
            width = maxWidth;
          }
        } else {
          if (height > maxHeight) {
            width = Math.round((width * maxHeight) / height);
            height = maxHeight;
          }
        }

        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0, width, height);

        const type = file.type === 'image/png' ? 'image/png' : 'image/jpeg';
        resolve(canvas.toDataURL(type, quality));
      };
      img.onerror = () => {
        resolve(event.target.result); // Fallback to original base64 on error
      };
    };
    reader.onerror = () => {
      resolve(null);
    };
  });
};

const getElapsedHours = (job, dataPointTimestamp) => {
  if (!job || !dataPointTimestamp) return 0;

  let startTimeMs = null;

  // Find the earliest timestamp among all data points in this job as the start time (Time Zero)
  if (job.data && job.data.length > 0) {
    let minTimeMs = Infinity;
    job.data.forEach(row => {
      if (row.timestamp) {
        const t = new Date(row.timestamp).getTime();
        if (!isNaN(t) && t < minTimeMs) {
          minTimeMs = t;
        }
      }
    });
    if (minTimeMs !== Infinity) {
      startTimeMs = minTimeMs;
    }
  }

  // Fallback to Job ID or Job creation time if no data points have valid timestamps
  if (!startTimeMs) {
    if (job.id && typeof job.id === 'string' && job.id.startsWith('job-')) {
      const idNum = parseInt(job.id.replace('job-', ''), 10);
      if (!isNaN(idNum)) {
        startTimeMs = idNum;
      }
    }
    if (!startTimeMs && job.createdAt) {
      const d = new Date(job.createdAt);
      if (!isNaN(d.getTime())) {
        startTimeMs = d.getTime();
      }
    }
  }

  if (!startTimeMs) return 0;

  const recordTimeMs = new Date(dataPointTimestamp).getTime();
  if (isNaN(recordTimeMs)) return 0;

  const diffMs = recordTimeMs - startTimeMs;
  const hours = diffMs / (1000 * 60 * 60);
  return Math.max(0, parseFloat(hours.toFixed(1)));
};

const CustomTooltip = ({ active, payload, label }) => {
  if (active && payload && payload.length) {
    const rowData = payload[0].payload;
    let titleText = label;
    if (rowData) {
      const timeVal = rowData.time || (rowData.timestamp ? new Date(rowData.timestamp).toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' }) : '');
      const hourVal = rowData.cultureHour !== undefined ? rowData.cultureHour : '';
      if (hourVal !== '' && timeVal !== '') {
        titleText = `ชั่วโมงที่ ${hourVal} (เวลา ${timeVal})`;
      } else if (hourVal !== '') {
        titleText = `ชั่วโมงที่ ${hourVal}`;
      } else if (timeVal !== '') {
        titleText = `เวลา ${timeVal}`;
      }
    }
    return (
      <div className="custom-tooltip">
        <p className="custom-tooltip-label">{titleText}</p>
        {payload.map((entry, index) => (
          <p key={index} style={{ color: entry.color, margin: 0, fontSize: '14px', fontWeight: 600 }}>
            {entry.name}: {typeof entry.value === 'number' ? entry.value.toFixed(2) : entry.value || '-'}
          </p>
        ))}
      </div>
    );
  }
  return null;
};

// BSTR Diagram Component
const BSTRDiagram = ({ dataPoint, chartData, isReplaying, isReplayingPlaying, jobStatus = 'running', onToggleStatus, userRole, isViewingHistory, theme, aboutSystem, machineName }) => {
  if (!dataPoint) {
    return (
      <div className="glass-panel empty-state" style={{ height: '500px', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center' }}>
        <h2>No Active Session or Data</h2>
        <p>เลือกเซสชันที่มีข้อมูล หรือกดเริ่มเพิ่มข้อมูลใหม่ก่อน</p>
      </div>
    );
  }

  const maxVolumeLiters = aboutSystem?.maxVolumeLiters !== undefined ? Number(aboutSystem.maxVolumeLiters) : 5.0;
  let physicalMaxVolume = maxVolumeLiters;
  const mName = (machineName || '').toLowerCase();
  if (mName.includes('750') || mName.includes('500')) {
    physicalMaxVolume = 750.0;
  } else if (mName.includes('70')) {
    physicalMaxVolume = 70.0;
  } else if (mName.includes('10')) {
    physicalMaxVolume = 10.0;
  }

  // Extract PV & SV
  const temp_set = typeof dataPoint.temp_set === 'number' ? dataPoint.temp_set : 37.0;
  const temp_read = typeof dataPoint.temp_read === 'number' ? dataPoint.temp_read : 37.0;
  const ph_set = typeof dataPoint.ph_set === 'number' ? dataPoint.ph_set : 7.00;
  const ph_read = typeof dataPoint.ph_read === 'number' ? dataPoint.ph_read : 7.00;
  const do_set = typeof dataPoint.do_set === 'number' ? dataPoint.do_set : 60.0;
  const do_read = typeof dataPoint.do_read === 'number' ? dataPoint.do_read : 60.0;
  const agit_set = typeof dataPoint.agit_set === 'number' ? dataPoint.agit_set : 120.0;
  const agit_read = typeof dataPoint.agit_read === 'number' ? dataPoint.agit_read : 120.0;
  const air_set = typeof dataPoint.air_set === 'number' ? dataPoint.air_set : 2.0;
  const air_read = typeof dataPoint.air_read === 'number' ? dataPoint.air_read : 2.0;

  const level_set = typeof dataPoint.level_set === 'number' ? dataPoint.level_set : (0.65 * maxVolumeLiters);
  const level_read = typeof dataPoint.level_read === 'number' ? dataPoint.level_read : (0.65 * maxVolumeLiters);
  const air_out_set = typeof dataPoint.air_out_set === 'number' ? dataPoint.air_out_set : parseFloat((air_set * 0.96).toFixed(2));
  const air_out_read = typeof dataPoint.air_out_read === 'number' ? dataPoint.air_out_read : parseFloat((air_read * 0.96).toFixed(2));
  const heat_set = typeof dataPoint.heat_set === 'number' ? dataPoint.heat_set : 0.0;
  const heat_read = typeof dataPoint.heat_read === 'number' ? dataPoint.heat_read : 0.0;

  // Warning conditions: RPM is 250 AND Air is > 500
  const isRpm250 = (Math.round(agit_read) === 250 || Math.round(agit_set) === 250);
  const isAirHigh = (air_read > 500 || air_set > 500);
  const showAirOutWarning = isRpm250 && isAirHigh;

  // VVM calculation: (Air Flow in L/min) / (Working Liquid/Fermentation Volume in Liters)
  const vvmCalcType = aboutSystem?.vvmCalcType || 'dynamic';
  const airUnit = aboutSystem?.airUnit || 'mlmin';

  // Convert air_read to L/min if it is in mL/min
  const airLitersPerMinute = airUnit === 'mlmin' ? (air_read / 1000) : air_read;

  let workingVolumeLiters = 5.0;
  if (vvmCalcType === 'constant') {
    workingVolumeLiters = aboutSystem?.constantVolumeLiters !== undefined ? Number(aboutSystem.constantVolumeLiters) : 3.5;
  } else {
    workingVolumeLiters = level_read > 0 ? level_read : maxVolumeLiters;
  }
  if (workingVolumeLiters <= 0) workingVolumeLiters = 5.0; // protection against zero/negative division
  const rawVvm = airLitersPerMinute / workingVolumeLiters;
  const calculatedVvm = rawVvm < 0.1 ? rawVvm.toFixed(4) : rawVvm.toFixed(2);
  // Display air flow always in L/min for clarity
  const airDisplayLpm = airUnit === 'mlmin' ? (air_read / 1000) : air_read;
  const airSetDisplayLpm = airUnit === 'mlmin' ? (air_set / 1000) : air_set;

  // Deriving timestamp formatting
  const pad2 = (n) => String(n).padStart(2, '0');
  const formatTimeHHMM = (ts) => {
    if (!ts) return '';
    const d = new Date(ts);
    if (isNaN(d.getTime())) return ts;
    return `${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
  };

  const timestampStr = dataPoint.timestamp
    ? (() => {
      const d = new Date(dataPoint.timestamp);
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}:${String(d.getSeconds()).padStart(2, '0')}`;
    })()
    : `${dataPoint.date || ''} ${dataPoint.time || ''}`;

  // Helper to draw inline SVG sparklines
  const renderSparkline = (data, key, strokeColor, minVal, maxVal) => {
    const width = 120;
    const height = 40;
    if (!data || data.length === 0) {
      return (
        <div style={{ height, color: 'var(--text-secondary)', fontSize: '0.8rem', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
          No Data
        </div>
      );
    }

    const points = data.length === 1
      ? (() => {
          const val = data[0][key] !== undefined ? data[0][key] : 0;
          const range = maxVal - minVal;
          const normalized = range > 0 ? (val - minVal) / range : 0.5;
          const y = height - Math.min(height, Math.max(0, normalized * height));
          return `0,${y} ${width},${y}`;
        })()
      : data.map((d, i) => {
          const val = d[key] !== undefined ? d[key] : 0;
          const x = (i / Math.max(1, data.length - 1)) * width;
          const range = maxVal - minVal;
          const normalized = range > 0 ? (val - minVal) / range : 0.5;
          const y = height - Math.min(height, Math.max(0, normalized * height));
          return `${x},${y}`;
        }).join(' ');

    return (
      <svg width="100%" height={height} viewBox={`0 0 ${width} ${height}`} preserveAspectRatio="none" style={{ overflow: 'visible' }}>
        <polyline
          fill="none"
          stroke={strokeColor}
          strokeWidth="2"
          strokeLinecap="round"
          strokeLinejoin="round"
          points={points}
        />
      </svg>
    );
  };

  // Extract start, mid, end culture hours for sparkline x-axis label
  let startLabel = '0.0 ชม.';
  let midLabel = '0.0 ชม.';
  let endLabel = '0.0 ชม.';
  if (chartData && chartData.length > 0) {
    startLabel = `ชั่วโมงที่ ${chartData[0].cultureHour !== undefined ? chartData[0].cultureHour : 0}`;
    const midIdx = Math.floor(chartData.length / 2);
    midLabel = `ชั่วโมงที่ ${chartData[midIdx].cultureHour !== undefined ? chartData[midIdx].cultureHour : 0}`;
    endLabel = `ชั่วโมงที่ ${chartData[chartData.length - 1].cultureHour !== undefined ? chartData[chartData.length - 1].cultureHour : 0}`;
  }

  const isMachineStopped = jobStatus === 'stopped' || jobStatus === 'finished';
  const isMachineStoppedVisual = isMachineStopped && !isReplaying && !isViewingHistory;

  // Impeller spinner speed calculation (CSS custom property)
  const agitSpeedSec = !isMachineStoppedVisual && agit_read > 0 ? Math.max(0.1, 60 / agit_read) : 0;
  const agitSpinStyle = agitSpeedSec > 0 ? {
    animation: `spin-blade ${agitSpeedSec}s linear infinite`,
    transformOrigin: '150px 307.5px'
  } : {};

  // Calculate liquid surface Y coordinate and travel distance for bubbles
  const liquidTopY = 360 - (level_read / physicalMaxVolume) * 230;
  const travelDist = 340 - liquidTopY;

  // Bubbling animation styling based on Air Flow (VVM and L/min)
  const airBubbleCount = !isMachineStoppedVisual && airLitersPerMinute > 0 ? Math.min(150, Math.floor(rawVvm * 80) + 20) : 0;

  // Heating power color intensity mapping for the jacket glow
  const heatReadForVisual = isMachineStoppedVisual ? 0 : heat_read;
  const jacketOpacity = Math.min(0.8, Math.max(0.1, heatReadForVisual / 100));

  return (
    <div className="bstr-diagram-dashboard">
      {/* 1. DIAGRAM HEADER BAR */}
      <div className="diagram-header-bar">
        <div className="diagram-title-section">
          <h2>BATCH STIRRED TANK REACTOR (BSTR)</h2>
          <p className="diagram-subtitle">
            {machineName ? machineName : 'PROCESS MONITORING DIAGRAM'}
          </p>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
          {userRole === 'admin' && !isReplaying && jobStatus !== 'finished' && (
            <button
              onClick={onToggleStatus}
              className="status-control-btn btn-stop-machine"
              style={{ background: 'linear-gradient(135deg, #dc2626, #b91c1c)', borderColor: '#dc2626' }}
            >
              ✅ เสร็จสิ้นงาน
            </button>
          )}
          <div className="diagram-datetime-box">
            📅 {timestampStr}
          </div>
        </div>
      </div>

      {/* 2. MAIN HMI DISPLAY AREA */}
      <div className="diagram-main-grid">

        {/* Left Side: Overlaid Sensor Callout Boxes */}
        <div className="diagram-sensors-col">

          {/* pH Sensor Box */}
          <div className="sensor-callout-card ph-sensor-callout">
            <div className="sensor-callout-header">
              <span className="sensor-icon">pH</span>
              <span className="sensor-name">pH SENSOR</span>
            </div>
            <div className="sensor-lcd-display">
              <span className="lcd-value">{ph_read.toFixed(2)}</span>
              <span className="lcd-unit">pH</span>
            </div>
            <div className="sensor-sv-label">SV (Set): {ph_set.toFixed(2)}</div>
          </div>

          {/* DO Sensor Box */}
          <div className="sensor-callout-card do-sensor-callout">
            <div className="sensor-callout-header">
              <span className="sensor-icon">DO</span>
              <span className="sensor-name">DO SENSOR</span>
            </div>
            <div className="sensor-lcd-display green-lcd">
              <span className="lcd-value">{Math.round(do_read)}</span>
              <span className="lcd-unit">%</span>
            </div>
            <div className="sensor-sv-label">SV (Set): {Math.round(do_set)}%</div>
          </div>

          {/* Temperature Sensor Box */}
          <div className="sensor-callout-card temp-sensor-callout">
            <div className="sensor-callout-header">
              <span className="sensor-icon">🌡️</span>
              <span className="sensor-name">TEMP SENSOR</span>
            </div>
            <div className="sensor-lcd-display red-lcd">
              <span className="lcd-value">{temp_read.toFixed(1)}</span>
              <span className="lcd-unit">°C</span>
            </div>
            <div className="sensor-sv-label">SV (Set): {temp_set.toFixed(1)}°C</div>
          </div>

          {/* Level Sensor Box */}
          <div className="sensor-callout-card level-sensor-callout">
            <div className="sensor-callout-header">
              <span className="sensor-icon">🎚️</span>
              <span className="sensor-name">LEVEL SENSOR</span>
            </div>
            <div className="sensor-lcd-display blue-lcd">
              <span className="lcd-value">{level_read.toFixed(1)}</span>
              <span className="lcd-unit">L</span>
            </div>
            <div className="sensor-sv-label">SV (Set): {level_set.toFixed(1)} L</div>
          </div>

        </div>

        {/* Center: Scalable Reactor SVG Vessel */}
        <div className="diagram-reactor-vessel-col">

          <svg className="reactor-vessel-svg" viewBox="0 0 300 480" width="100%" height="100%">
            <defs>
              {/* Metallic Gradients */}
              <linearGradient id="metal-grad" x1="0%" y1="0%" x2="100%" y2="0%">
                <stop offset="0%" stopColor="#4b5563" />
                <stop offset="30%" stopColor="#9ca3af" />
                <stop offset="50%" stopColor="#e5e7eb" />
                <stop offset="70%" stopColor="#9ca3af" />
                <stop offset="100%" stopColor="#4b5563" />
              </linearGradient>
              <linearGradient id="media-grad" x1="0%" y1="0%" x2="0%" y2="100%">
                <stop offset="0%" stopColor="#f59e0b" stopOpacity="0.7" />
                <stop offset="100%" stopColor="#d97706" stopOpacity="0.85" />
              </linearGradient>
              <linearGradient id="jacket-grad" x1="0%" y1="0%" x2="100%" y2="0%">
                <stop offset="0%" stopColor="#ef4444" stopOpacity="0.8" />
                <stop offset="50%" stopColor="#f87171" stopOpacity="0.4" />
                <stop offset="100%" stopColor="#ef4444" stopOpacity="0.8" />
              </linearGradient>
              <clipPath id="liquid-clip">
                <path d="M 74 350 L 74 125 L 226 125 L 226 350 A 76 46 0 0 1 74 350 Z" />
              </clipPath>
              <filter id="liquid-glow" x="-10%" y="-10%" width="120%" height="120%">
                <feGaussianBlur in="SourceGraphic" stdDeviation="3" result="blur" />
                <feMerge>
                  <feMergeNode in="blur" />
                  <feMergeNode in="SourceGraphic" />
                </feMerge>
              </filter>
            </defs>

            {/* Heating Jacket Glow (Wrapper behind reactor) */}
            {!isMachineStoppedVisual && heat_read > 0 && (
              <path
                d="M 68 180 L 68 340 A 82 82 0 0 0 232 340 L 232 180 Z"
                fill="url(#jacket-grad)"
                filter="url(#glow)"
                opacity={jacketOpacity}
                stroke="#f87171"
                strokeWidth="4"
              />
            )}
            <filter id="glow" x="-20%" y="-20%" width="140%" height="140%">
              <feGaussianBlur stdDeviation="6" result="blur" />
              <feComposite in="SourceGraphic" in2="blur" operator="over" />
            </filter>

            {/* Main Stainless Tank outline */}
            {/* Top curved head dome */}
            <path d="M 70 120 A 80 50 0 0 1 230 120" fill="url(#metal-grad)" stroke="#1e293b" strokeWidth="2" />
            {/* Reactor Body Cylindrical Wall */}
            <rect x="70" y="120" width="160" height="230" fill="url(#metal-grad)" stroke="#1e293b" strokeWidth="2" />
            {/* Bottom curved bottom dish */}
            <path d="M 70 350 A 80 50 0 0 0 230 350" fill="url(#metal-grad)" stroke="#1e293b" strokeWidth="2" />

            {/* Inside Liquid Media (Golden broth) */}
            {level_read > 0 && (() => {
              const amp = isMachineStoppedVisual ? 0.5 : Math.min(4, 1 + (agit_read / 80));
              const wavePath = `M -100 ${liquidTopY}
                Q -75 ${liquidTopY - amp} -50 ${liquidTopY}
                T 0 ${liquidTopY}
                T 50 ${liquidTopY}
                T 100 ${liquidTopY}
                T 150 ${liquidTopY}
                T 200 ${liquidTopY}
                T 250 ${liquidTopY}
                T 300 ${liquidTopY}
                T 350 ${liquidTopY}
                L 350 400 L -100 400 Z`;
              const waveSpeedSec = !isMachineStoppedVisual && agit_read > 0
                ? Math.max(0.6, 5 - (agit_read / 100))
                : 8;
              return (
                <g clipPath="url(#liquid-clip)">
                  <path
                    d={wavePath}
                    fill="url(#media-grad)"
                    stroke="#d97706"
                    strokeWidth="1"
                    className="liquid-wave-animated"
                    filter={theme === 'hmi' ? 'url(#liquid-glow)' : 'none'}
                    style={{
                      animationDuration: `${waveSpeedSec}s`
                    }}
                  />
                </g>
              );
            })()}

            {/* Foam Layer on Liquid Surface */}
            {airLitersPerMinute > 0 && level_read > 0 && (
              <g clipPath="url(#liquid-clip)">
                {(() => {
                  const foamCircles = [];
                  const layers = Math.min(3, Math.max(1, Math.floor(rawVvm * 1.5) + 1));
                  const bubblesPerLayer = Math.min(30, Math.floor(rawVvm * 10) + 12);
                  
                  for (let layer = 0; layer < layers; layer++) {
                    const yOffset = liquidTopY - (layer * 3.5);
                    const opacity = 0.5 - (layer * 0.1);
                    
                    for (let j = 0; j < bubblesPerLayer; j++) {
                      const ratio = j / (bubblesPerLayer - 1 || 1);
                      const rx = 76 + ratio * 148 + (Math.sin(j + layer) * 2.5);
                      const rRadius = 2.5 + ((j + layer) % 3) * 1.2;
                      const floatAnimDelay = `${(j * 0.06 + layer * 0.2).toFixed(2)}s`;
                      const floatAnimDuration = `${1.2 + (j % 3) * 0.4}s`;
                      
                      foamCircles.push(
                        <circle
                          key={`foam-${layer}-${j}`}
                          cx={rx}
                          cy={yOffset}
                          r={rRadius}
                          fill="#ffffff"
                          opacity={opacity}
                          style={{
                            animation: `foam-float ${floatAnimDuration} infinite ease-in-out`,
                            animationDelay: floatAnimDelay
                          }}
                        />
                      );
                    }
                  }
                  return foamCircles;
                })()}
              </g>
            )}

            {/* Rising Bubbles (Simulated Gas Sparger) */}
            {airBubbleCount > 0 && level_read > 0 && (
              <g>
                {[...Array(airBubbleCount)].map((_, i) => {
                  const rx = 80 + (i * 13) % 140;
                  const ry = 340 - (i * 17) % (Math.max(20, (level_read / physicalMaxVolume) * 210));
                  const rRadius = 1.5 + (i % 3);
                  return (
                    <circle
                      key={`bubble-${i}`}
                      cx={rx}
                      cy={ry}
                      r={rRadius}
                      fill="#ffffff"
                      opacity="0.6"
                      className="rising-bubble-element"
                      style={{
                        animation: `bubble-rise ${1 + (i % 2)}s infinite ease-in-out`,
                        animationDelay: `${i * 0.1}s`
                      }}
                    />
                  );
                })}
              </g>
            )}

            {/* Central Stirrer Impeller Shaft */}
            <line x1="150" y1="100" x2="150" y2="330" stroke="#1f2937" strokeWidth="6" strokeLinecap="round" />
            <line x1="150" y1="100" x2="150" y2="330" stroke="#d1d5db" strokeWidth="2" />

            {/* Sparger Ring at bottom (Static) */}
            <line x1="100" y1="340" x2="200" y2="340" stroke="#4b5563" strokeWidth="3" strokeDasharray="3,3" />

            {/* Agitator Blades (Impellers) - 3 Sets on Shaft */}
            {agit_read > 0 ? (
              <g style={agitSpinStyle}>
                {/* 1. TOP IMPELLER: Foam Cutter (ใบตัดฟอง) */}
                <rect x="140" y="132" width="20" height="16" rx="2" fill="#111827" />
                <path d="M 105 136 L 140 137 L 140 143 L 105 144 Z" fill="#4b5563" stroke="#1f2937" strokeWidth="1" />
                <path d="M 195 136 L 160 137 L 160 143 L 195 144 Z" fill="#4b5563" stroke="#1f2937" strokeWidth="1" />
                {/* Antifoam vertical tooth elements at the end of blades */}
                <rect x="104" y="130" width="3" height="18" fill="#ffb800" />
                <rect x="193" y="130" width="3" height="18" fill="#ffb800" />

                {/* 2. MIDDLE IMPELLER: Mixing (ใบกวนผสม) */}
                <rect x="135" y="212" width="30" height="15" rx="3" fill="#111827" />
                <path d="M 100 214 L 135 217 L 135 222 L 100 225 Z" fill="#374151" stroke="#1f2937" strokeWidth="1" />
                <path d="M 200 214 L 165 217 L 165 222 L 200 225 Z" fill="#374151" stroke="#1f2937" strokeWidth="1" />

                {/* 3. BOTTOM IMPELLER: Gas Dispersion (ใบกวนล่างสุด) */}
                <rect x="135" y="300" width="30" height="15" rx="3" fill="#111827" />
                <path d="M 100 302 L 135 305 L 135 310 L 100 313 Z" fill="#374151" stroke="#1f2937" strokeWidth="1" />
                <path d="M 200 302 L 165 305 L 165 310 L 200 313 Z" fill="#374151" stroke="#1f2937" strokeWidth="1" />
              </g>
            ) : (
              <g>
                {/* 1. TOP IMPELLER: Foam Cutter (ใบตัดฟอง) */}
                <rect x="140" y="132" width="20" height="16" rx="2" fill="#374151" />
                <path d="M 105 136 L 140 137 L 140 143 L 105 144 Z" fill="#4b5563" />
                <path d="M 195 136 L 160 137 L 160 143 L 195 144 Z" fill="#4b5563" />
                <rect x="104" y="130" width="3" height="18" fill="#d97706" />
                <rect x="193" y="130" width="3" height="18" fill="#d97706" />

                {/* 2. MIDDLE IMPELLER: Mixing (ใบกวนผสม) */}
                <rect x="135" y="212" width="30" height="15" rx="3" fill="#374151" />
                <path d="M 100 214 L 135 217 L 135 222 L 100 225 Z" fill="#4b5563" />
                <path d="M 200 214 L 165 217 L 165 222 L 200 225 Z" fill="#4b5563" />

                {/* 3. BOTTOM IMPELLER: Gas Dispersion (ใบกวนล่างสุด) */}
                <rect x="135" y="300" width="30" height="15" rx="3" fill="#374151" />
                <path d="M 100 302 L 135 305 L 135 310 L 100 313 Z" fill="#4b5563" />
                <path d="M 200 302 L 165 305 L 165 310 L 200 313 Z" fill="#4b5563" />
              </g>
            )}

            {/* Top Motor Drive Assembly */}
            <rect x="135" y="45" width="30" height="45" fill="url(#metal-grad)" stroke="#1e293b" />
            <rect x="130" y="35" width="40" height="10" fill="#1e293b" />
            <circle cx="150" cy="35" r="4" fill="#ef4444" />

            {/* Pipes and Valvings */}
            {/* Gas Inlet pipe (Top-Right) */}
            <path d="M 230 65 L 180 65 L 180 115" fill="none" stroke={air_read > 0 ? "#065f46" : "#9ca3af"} strokeWidth="6" strokeLinecap="round" />
            {/* Gas Inlet Flow Line */}
            {air_read > 0 && (
              <path
                d="M 230 65 L 180 65 L 180 115"
                fill="none"
                stroke="#10b981"
                strokeWidth="3.5"
                strokeLinecap="round"
                className="inlet-air-flow-line"
              />
            )}
            {/* Gas Inlet Flow Arrow */}
            {air_read > 0 && (
              <path d="M 185 65 L 195 60 L 195 70 Z" fill="#ffffff" className="flow-arrow-animation-left" />
            )}

            {/* Air Outlet pipe (Top-Left) */}
            <path d="M 120 115 L 120 80 L 60 80" fill="none" stroke={air_out_read > 0 ? "#3b82f6" : "#9ca3af"} strokeWidth="6" strokeLinecap="round" />
            {/* Air Outlet Flow Arrow */}
            {air_out_read > 0 && (
              <path d="M 65 80 L 75 75 L 75 85 Z" fill="#ffffff" className="flow-arrow-animation-left" />
            )}

            {/* Warning Liquid flowing out when warning is triggered */}
            {showAirOutWarning && (
              <>
                <path
                  d="M 120 115 L 120 80 L 60 80"
                  fill="none"
                  stroke="#ff7c00"
                  strokeWidth="3.5"
                  strokeLinecap="round"
                  className="warning-liquid-flow-out"
                />
                <g className="warning-splash-droplets">
                  <circle cx="55" cy="80" r="2" fill="#ff7c00" className="droplet-1" />
                  <circle cx="50" cy="81" r="1.5" fill="#ff7c00" className="droplet-2" />
                  <circle cx="45" cy="79" r="2.5" fill="#ff7c00" className="droplet-3" />
                </g>
              </>
            )}

            {/* Warning Sign near Air Outlet Pipe */}
            {showAirOutWarning && (
              <g className="svg-warning-pulse" transform="translate(68, 48)">
                <path d="M 12 0 L 24 20 L 0 20 Z" fill="#ffb800" stroke="#ff3e3e" strokeWidth="2" />
                <text x="12" y="16" fill="#000000" fontSize="13" fontWeight="bold" textAnchor="middle">!</text>
              </g>
            )}

            {/* Bottom Product Harvest pipe */}
            <path d="M 150 380 L 150 420 L 120 420" fill="none" stroke="#4b5563" strokeWidth="8" strokeLinecap="round" />
            {/* Gas Control Valve at bottom right */}
            <path d="M 200 370 L 200 420" fill="none" stroke="#4b5563" strokeWidth="4" />
            <polygon points="190,385 210,385 190,405 210,405" fill="#374151" stroke="#1f2937" strokeWidth="1" />

            {/* Volume Scale Ruler (Left Wall) — 50 L ticks from maxVolumeLiters down to 0 */}
            {(() => {
              const tankBottom = 350;
              const tankH = 230; // px for full volume
              const volToY = (vol) => tankBottom - (vol / physicalMaxVolume) * tankH;

              // Determine capacity, physical size, and custom ticks based on machineName
              let physicalMaxVolume = maxVolumeLiters;
              let ticks = [];
              let step = 50;
              let isCustom = false;

              const mName = (machineName || '').toLowerCase();
              if (mName.includes('750') || mName.includes('500')) {
                physicalMaxVolume = 750.0;
                step = 50;
                isCustom = true;
                for (let v = 100; v <= 750; v += 50) {
                  ticks.push(v);
                }
              } else if (mName.includes('70')) {
                physicalMaxVolume = 70.0;
                step = 5;
                isCustom = true;
                for (let v = 20; v <= 70; v += 5) {
                  ticks.push(v);
                }
              } else if (mName.includes('10')) {
                physicalMaxVolume = 10.0;
                step = 1;
                isCustom = true;
                for (let v = 1; v <= 10; v += 1) {
                  ticks.push(v);
                }
              } else {
                step = maxVolumeLiters <= 20 ? 2
                  : maxVolumeLiters <= 100 ? 10
                  : maxVolumeLiters <= 300 ? 25
                  : 50;
                for (let v = step; v < maxVolumeLiters; v += step) {
                  ticks.push(v);
                }
              }

              return (
                <g>
                  {/* Vertical ruler line */}
                  <line x1="63" y1={volToY(physicalMaxVolume)} x2="63" y2={volToY(0)} stroke="#64748b" strokeWidth="1" />
                  
                  {/* 0 label (always visible) */}
                  <line x1="58" y1={volToY(0)} x2="70" y2={volToY(0)} stroke="#64748b" strokeWidth="1" />
                  <text x="55" y={volToY(0) + 4} fill="#64748b" fontSize="7.5" textAnchor="end" fontFamily="monospace">0L</text>
                  
                  {/* For fallback: draw max volume label separately if not custom */}
                  {!isCustom && (
                    <>
                      <line x1="58" y1={volToY(maxVolumeLiters)} x2="70" y2={volToY(maxVolumeLiters)} stroke="#94a3b8" strokeWidth="1.5" />
                      <text x="55" y={volToY(maxVolumeLiters) + 4} fill="#94a3b8" fontSize="7.5" textAnchor="end" fontFamily="monospace">{maxVolumeLiters.toFixed(0)}L</text>
                    </>
                  )}

                  {/* Intermediate/Custom ticks */}
                  {ticks.map((v) => {
                    const y = volToY(v);
                    const isMajor =
                      physicalMaxVolume === 750.0 ? (v % 100 === 0 || v === 750) :
                      physicalMaxVolume === 70.0 ? (v % 10 === 0) :
                      physicalMaxVolume === 10.0 ? (v % 2 === 0 || v === 1) :
                      (v % (step * 2) === 0);
                    // Current level indicator line
                    const isNearLevel = Math.abs(v - level_read) < step * 0.5;
                    return (
                      <g key={v}>
                        <line
                          x1={isMajor ? "56" : "60"}
                          y1={y}
                          x2="70"
                          y2={y}
                          stroke={isNearLevel ? '#f59e0b' : (isMajor ? '#94a3b8' : '#475569')}
                          strokeWidth={isMajor ? "1.5" : "1"}
                        />
                        {isMajor && (
                          <text
                            x="53"
                            y={y + 4}
                            fill={isNearLevel ? '#f59e0b' : '#94a3b8'}
                            fontSize="7.5"
                            textAnchor="end"
                            fontFamily="monospace"
                            fontWeight={isNearLevel ? 'bold' : 'normal'}
                          >
                            {v}L
                          </text>
                        )}
                      </g>
                    );
                  })}
                  {/* Current liquid level indicator arrow */}
                  {level_read > 0 && level_read <= physicalMaxVolume && (
                    <g>
                      <line x1="56" y1={volToY(level_read)} x2="80" y2={volToY(level_read)} stroke="#f59e0b" strokeWidth="1.5" strokeDasharray="2 2" />
                      <polygon
                        points={`70,${volToY(level_read) - 4} 70,${volToY(level_read) + 4} 80,${volToY(level_read)}`}
                        fill="#f59e0b"
                      />
                    </g>
                  )}
                </g>
              );
            })()}
          </svg>

          {/* Motor Drive RPM Display Box Overlaid on Vessel Head */}
          <div className="motor-drive-display">
            <span className="display-label">MOTOR RPM</span>
            <div className="display-screen">
              <span className="screen-val">{isMachineStoppedVisual ? 0 : Math.round(agit_read)}</span>
              <span className="screen-unit">RPM</span>
            </div>
            <span className="display-sv">SV: {Math.round(agit_set)}</span>
          </div>

          {/* Gas Inlet Flow slpm Box Overlaid on Inlet Pipe */}
          <div className="gas-inlet-display">
            <span className="display-label">AIR FLOW IN</span>
            <div className="display-screen green-lcd" style={{ display: 'flex', flexDirection: 'column', alignItems: 'stretch', gap: '2px', minWidth: '75px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline' }}>
                <span className="screen-val" style={{ color: 'var(--accent-green)', textShadow: '0 0 5px rgba(16,185,129,0.8)' }}>{airDisplayLpm.toFixed(2)}</span>
                <span className="screen-unit">L/min</span>
              </div>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', borderTop: '1px solid rgba(16,185,129,0.2)', paddingTop: '2px' }}>
                <span className="screen-val" style={{ color: 'var(--accent-green)', textShadow: '0 0 5px rgba(16,185,129,0.8)' }}>{calculatedVvm}</span>
                <span className="screen-unit">vvm</span>
              </div>
            </div>
          </div>

          {/* Air Outlet Flow slpm Box Overlaid on Outlet Pipe */}
          <div className={`air-outlet-display ${showAirOutWarning ? 'air-outlet-warning-flashing' : ''}`}>
            <span className="display-label">AIR OUT (PMa)</span>
            <div className="display-screen blue-lcd">
              <span className="screen-val">{air_out_read.toFixed(1)}</span>
              <span className="screen-unit">PMa</span>
            </div>
            {showAirOutWarning && (
              <span style={{ fontSize: '0.6rem', color: '#ff3e3e', fontWeight: 'bold', marginTop: '2px', textAlign: 'center' }}>
                ⚠️ ระวัง! RPM & AIR สูง
              </span>
            )}
          </div>

        </div>

        {/* Right Side: Detailed Progress Bar Gauges */}
        <div className="diagram-gauges-col">
          <div className="gauges-panel-card">

            {/* Volume Gauge */}
            <div className="gauge-item">
              <div className="gauge-label-row">
                <span className="gauge-icon">🎚️</span>
                <span className="gauge-title">VOLUME (L)</span>
                <span className="gauge-value">{workingVolumeLiters.toFixed(1)} L</span>
              </div>
              <div className="gauge-track-bar">
                <div className="gauge-filled-bar green-bar" style={{ width: `${Math.min(100, Math.max(0, (level_read / physicalMaxVolume) * 100))}%` }}></div>
              </div>
              <div className="gauge-limits">
                <span>0 L</span>
                <span>{(physicalMaxVolume / 2).toFixed(1)} L</span>
                <span>{physicalMaxVolume.toFixed(1)} L</span>
              </div>
            </div>

            {/* Agitator RPM Gauge */}
            <div className="gauge-item">
              <div className="gauge-label-row">
                <span className="gauge-icon">⚙️</span>
                <span className="gauge-title">AGITATOR SPEED (RPM)</span>
                <span className="gauge-value">{agit_read.toFixed(1)} RPM</span>
              </div>
              <div className="gauge-track-bar">
                <div className="gauge-filled-bar purple-bar" style={{ width: `${Math.min(100, Math.max(0, (agit_read / 500) * 100))}%` }}></div>
              </div>
              <div className="gauge-limits">
                <span>0</span>
                <span>250</span>
                <span>500</span>
              </div>
            </div>

            {/* Air Flow In Gauge */}
            <div className="gauge-item">
              <div className="gauge-label-row">
                <span className="gauge-icon">🟢</span>
                <span className="gauge-title">AIR FLOW IN</span>
                <span className="gauge-value">{airDisplayLpm.toFixed(2)} L/min / {calculatedVvm} vvm</span>
              </div>
              <div className="gauge-track-bar">
                <div className="gauge-filled-bar green-bar" style={{ width: `${Math.min(100, Math.max(0, (airDisplayLpm / 1000) * 100))}%` }}></div>
              </div>
              <div className="gauge-limits">
                <span>0</span>
                <span>500 L/min</span>
                <span>1000 L/min</span>
              </div>
            </div>

            {/* Air Flow Out Gauge */}
            <div className="gauge-item">
              <div className="gauge-label-row">
                <span className="gauge-icon">🔵</span>
                <span className="gauge-title">AIR OUT (PMa)</span>
                <span className="gauge-value">{air_out_read.toFixed(1)} PMa</span>
              </div>
              <div className="gauge-track-bar">
                <div className="gauge-filled-bar blue-bar" style={{ width: `${Math.min(100, Math.max(0, (airUnit === 'mlmin' ? (air_out_read / 1000) : air_out_read) / 1000 * 100))}%` }}></div>
              </div>
              <div className="gauge-limits">
                <span>0</span>
                <span>500 L/min</span>
                <span>1000 L/min</span>
              </div>
            </div>

            {/* pH Gauge */}
            <div className="gauge-item">
              <div className="gauge-label-row">
                <span className="gauge-icon">pH</span>
                <span className="gauge-title">pH VALUE</span>
                <span className="gauge-value">{ph_read.toFixed(2)}</span>
              </div>
              <div className="gauge-track-bar">
                <div className="gauge-filled-bar green-bar" style={{ width: `${Math.min(100, Math.max(0, (ph_read / 14) * 100))}%` }}></div>
              </div>
              <div className="gauge-limits">
                <span>0</span>
                <span>7</span>
                <span>14</span>
              </div>
            </div>

            {/* DO Gauge */}
            <div className="gauge-item">
              <div className="gauge-label-row">
                <span className="gauge-icon">🫧</span>
                <span className="gauge-title">DISSOLVED OXYGEN (DO)</span>
                <span className="gauge-value">{do_read.toFixed(1)} %</span>
              </div>
              <div className="gauge-track-bar">
                <div className="gauge-filled-bar green-bar" style={{ width: `${Math.min(100, Math.max(0, (do_read / 200) * 100))}%` }}></div>
              </div>
              <div className="gauge-limits">
                <span>0</span>
                <span>100</span>
                <span>200</span>
              </div>
            </div>

            {/* Temperature Gauge */}
            <div className="gauge-item">
              <div className="gauge-label-row">
                <span className="gauge-icon">🌡️</span>
                <span className="gauge-title">TEMPERATURE</span>
                <span className="gauge-value">{temp_read.toFixed(1)} °C</span>
              </div>
              <div className="gauge-track-bar">
                <div className="gauge-filled-bar red-bar" style={{ width: `${Math.min(100, Math.max(0, (temp_read / 100) * 100))}%` }}></div>
              </div>
              <div className="gauge-limits">
                <span>0</span>
                <span>50</span>
                <span>100</span>
              </div>
            </div>

            {/* Heating Power Gauge */}
            <div className="gauge-item">
              <div className="gauge-label-row">
                <span className="gauge-icon">🔥</span>
                <span className="gauge-title">HEATING POWER</span>
                <span className="gauge-value">{heat_read.toFixed(1)} %</span>
              </div>
              <div className="gauge-track-bar">
                <div className="gauge-filled-bar orange-bar" style={{ width: `${Math.min(100, Math.max(0, heat_read))}%` }}></div>
              </div>
              <div className="gauge-limits">
                <span>0</span>
                <span>50</span>
                <span>100</span>
              </div>
            </div>

          </div>
        </div>

      </div>

      {/* 3. BOTTOM SPARKLINE TRENDS PANEL */}
      <div className="diagram-sparklines-row">

        {/* Sparkline 1: Volume */}
        <div className="sparkline-card">
          <span className="sparkline-title">VOLUME (L)</span>
          <div className="sparkline-chart-area">
            {renderSparkline(chartData, 'level_read', '#10b981', 0, physicalMaxVolume)}
          </div>
          <div className="sparkline-axis-labels">
            <span>{startLabel}</span>
            <span>{midLabel}</span>
            <span>{endLabel}</span>
          </div>
        </div>

        {/* Sparkline 2: RPM */}
        <div className="sparkline-card">
          <span className="sparkline-title">RPM</span>
          <div className="sparkline-chart-area">
            {renderSparkline(chartData, 'agit_read', '#8b5cf6', 0, 500)}
          </div>
          <div className="sparkline-axis-labels">
            <span>{startLabel}</span>
            <span>{midLabel}</span>
            <span>{endLabel}</span>
          </div>
        </div>

        {/* Sparkline 3: Air In */}
        <div className="sparkline-card">
          <span className="sparkline-title">AIR IN (L/M)</span>
          <div className="sparkline-chart-area">
            {renderSparkline(chartData, 'air_read', '#10b981', 0, 1000)}
          </div>
          <div className="sparkline-axis-labels">
            <span>{startLabel}</span>
            <span>{midLabel}</span>
            <span>{endLabel}</span>
          </div>
        </div>

        {/* Sparkline 4: Air Out */}
        <div className="sparkline-card">
          <span className="sparkline-title">AIR OUT (PMa)</span>
          <div className="sparkline-chart-area">
            {renderSparkline(chartData, 'air_out_read', '#3b82f6', 0, 1000)}
          </div>
          <div className="sparkline-axis-labels">
            <span>{startLabel}</span>
            <span>{midLabel}</span>
            <span>{endLabel}</span>
          </div>
        </div>

        {/* Sparkline 5: pH */}
        <div className="sparkline-card">
          <span className="sparkline-title">pH</span>
          <div className="sparkline-chart-area">
            {renderSparkline(chartData, 'ph_read', '#10b981', 0, 14)}
          </div>
          <div className="sparkline-axis-labels">
            <span>{startLabel}</span>
            <span>{midLabel}</span>
            <span>{endLabel}</span>
          </div>
        </div>

        {/* Sparkline 6: DO */}
        <div className="sparkline-card">
          <span className="sparkline-title">DO (%)</span>
          <div className="sparkline-chart-area">
            {renderSparkline(chartData, 'do_read', '#10b981', 0, 200)}
          </div>
          <div className="sparkline-axis-labels">
            <span>{startLabel}</span>
            <span>{midLabel}</span>
            <span>{endLabel}</span>
          </div>
        </div>

        {/* Sparkline 7: Temp */}
        <div className="sparkline-card">
          <span className="sparkline-title">TEMP (°C)</span>
          <div className="sparkline-chart-area">
            {renderSparkline(chartData, 'temp_read', '#ef4444', 0, 100)}
          </div>
          <div className="sparkline-axis-labels">
            <span>{startLabel}</span>
            <span>{midLabel}</span>
            <span>{endLabel}</span>
          </div>
        </div>

        {/* Sparkline 8: Heat */}
        <div className="sparkline-card">
          <span className="sparkline-title">HEAT (%)</span>
          <div className="sparkline-chart-area">
            {renderSparkline(chartData, 'heat_read', '#f59e0b', 0, 100)}
          </div>
          <div className="sparkline-axis-labels">
            <span>{startLabel}</span>
            <span>{midLabel}</span>
            <span>{endLabel}</span>
          </div>
        </div>

      </div>

      {/* 4. FOOTER STATUS BAR */}
      <div className="diagram-status-footer-bar">
        <div className="status-indicator-tag">
          STATUS: <span className={`status-badge-val ${isReplaying ? (isReplayingPlaying ? 'replaying-badge' : 'paused-badge') : (jobStatus === 'finished' ? 'stopped-badge' : isMachineStopped ? 'stopped-badge' : 'running-badge')}`}>
            {isReplaying
              ? (isReplayingPlaying ? 'PLAYBACK' : 'PAUSED')
              : (jobStatus === 'finished' ? 'FINISHED' : isMachineStopped ? 'STOPPED' : 'RUNNING')}
          </span>
        </div>
        <div className="action-buttons-indicator">
          <span className="indicator-btn-mock active-btn-mock">TREND</span>
          <span className="indicator-btn-mock">ALARM</span>
          <span className="indicator-btn-mock">SETTINGS</span>
        </div>
        <div className="user-role-indicator">
          USER: <strong>OPERATOR</strong>
        </div>
      </div>

    </div>
  );
};
const defaultMachine = { id: 'm1', name: 'Bioreactor 1' };
const defaultJob = {
  id: 'job-' + Date.now(),
  machineId: 'm1',
  name: 'Default Session',
  createdAt: new Date().toISOString(),
  data: []
};

function App() {
  // Helper: normalize time strings or Date to HH:MM for <input type="time">
  const pad2 = (n) => String(n).padStart(2, '0');
  const toHHMM = (v) => {
    if (!v) return '';
    if (v instanceof Date) return `${pad2(v.getHours())}:${pad2(v.getMinutes())}`;
    if (typeof v === 'string') {
      // If it looks like an ISO string or full date-time string, parse it as a Date first
      if (v.includes('-') || v.includes('T')) {
        const d = new Date(v);
        if (!isNaN(d.getTime())) return `${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
      }
      // try match HH:MM or HH:MM:SS
      const m = v.match(/^(\d{1,2}):(\d{2})/);
      if (m) return `${pad2(parseInt(m[1], 10))}:${pad2(parseInt(m[2], 10))}`;
    }
    return '';
  };

  const toYYYYMMDD = (v) => {
    if (!v) return '';
    const d = new Date(v);
    if (isNaN(d.getTime())) return '';
    return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
  };
  const formatDateTime = (v) => {
    if (!v) return '';
    const d = new Date(v);
    if (isNaN(d.getTime())) return v;
    return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())} ${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
  };
  const formatCreatedAt = (v) => {
    if (!v) return '';
    let d;
    if (v.includes(',')) {
      d = new Date(v.split(',')[0]);
    } else if (v.includes(' ')) {
      d = new Date(v.split(' ')[0]);
    } else {
      d = new Date(v);
    }
    if (isNaN(d.getTime())) {
      const parts = v.split(/[\s,T]/);
      if (parts.length > 0) return parts[0];
      return v;
    }
    const pad2 = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
  };
  // Authentication State
  const [userRole, setUserRole] = useState(() => {
    return localStorage.getItem('bioprocess-role') || null; // 'admin' | 'customer' | null
  });
  const [activeCustomerJobId, setActiveCustomerJobId] = useState(() => {
    return localStorage.getItem('bioprocess-customer-job-id') || null;
  });
  // Message shown on login screen when session expires
  const [sessionExpiredMessage, setSessionExpiredMessage] = useState(null);
  const [hasFetchedDB, setHasFetchedDB] = useState(false);

  // Developer & System Info states
  const [aboutSystem, setAboutSystem] = useState({
    systemName: 'DBMS (Bioprocess Data Logging)',
    systemVersion: 'v2.4.0 (SCADA Polish)',
    developer: 'ทีมวิศวกรรมข้อมูลชีวภาพ (Bioprocess Engineering Team)',
    techStack: 'React / Vite / Node.js / GCS',
    supportEmail: 'support@bioprocess-logging.local',
    supportPhone: '+66 2 123 4567',
    vvmCalcType: 'dynamic',
    maxVolumeLiters: 5.0,
    constantVolumeLiters: 3.5,
    airUnit: 'mlmin'
  });
  const [isEditingAbout, setIsEditingAbout] = useState(false);
  const [showEditMachineModal, setShowEditMachineModal] = useState(false);
  const [editingMachineId, setEditingMachineId] = useState('');
  const [editingMachineName, setEditingMachineName] = useState('');

  

  // PWA Install Prompt State
  const [deferredPrompt, setDeferredPrompt] = useState(null);
  const [isStandalone, setIsStandalone] = useState(false);

  // Unique client ID for active users status tracking
  const [clientId] = useState(() => {
    let id = sessionStorage.getItem('bioprocess-client-id');
    if (!id) {
      id = 'client-' + Math.random().toString(36).substring(2, 11);
      try {
        sessionStorage.setItem('bioprocess-client-id', id);
      } catch (e) {
        // ignore storage errors
      }
    }
    return id;
  });

  const [activeUsers, setActiveUsers] = useState([]);

  useEffect(() => {
    const checkStandalone = window.navigator.standalone === true ||
      window.matchMedia('(display-mode: standalone)').matches;
    setIsStandalone(checkStandalone);

    const handleBeforeInstallPrompt = (e) => {
      e.preventDefault();
      setDeferredPrompt(e);
    };

    window.addEventListener('beforeinstallprompt', handleBeforeInstallPrompt);
    return () => {
      window.removeEventListener('beforeinstallprompt', handleBeforeInstallPrompt);
    };
  }, []);

  const handleInstallApp = async () => {
    if (!deferredPrompt) return;
    deferredPrompt.prompt();
    const { outcome } = await deferredPrompt.userChoice;
    console.log(`User response to install: ${outcome}`);
    setDeferredPrompt(null);
  };

  useEffect(() => {
    if (userRole) {
      localStorage.setItem('bioprocess-role', userRole);
    } else {
      localStorage.removeItem('bioprocess-role');
    }
  }, [userRole]);

  useEffect(() => {
    if (activeCustomerJobId) {
      localStorage.setItem('bioprocess-customer-job-id', activeCustomerJobId);
    } else {
      localStorage.removeItem('bioprocess-customer-job-id');
    }
  }, [activeCustomerJobId]);

  // Global View State
  const [currentAppView, setCurrentAppView] = useState('monitoring'); // 'monitoring' | 'customers'
  const [isInstrumentsExpanded, setIsInstrumentsExpanded] = useState(true);
  const [isSessionsExpanded, setIsSessionsExpanded] = useState(true);
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);

  // Load States from Backend
  const [machines, setMachines] = useState([]);
  const [currentMachineId, setCurrentMachineId] = useState(null);
  const [jobs, setJobs] = useState([]);
  const [currentJobId, setCurrentJobId] = useState(null);
  const [isViewingHistory, setIsViewingHistory] = useState(false);
  const [customers, setCustomers] = useState([]);
  const [feedbacks, setFeedbacks] = useState([]);
  const [showFeedbackModal, setShowFeedbackModal] = useState(false);
  const [feedbackJobId, setFeedbackJobId] = useState('');
  const [feedbackSuccess, setFeedbackSuccess] = useState(false);
  // New survey state
  const SURVEY_QUESTIONS = [
    { id: 'q1', section: 'กระบวนการให้บริการ', text: 'ขั้นตอนการให้บริการมีความเหมาะสม คล่องตัว เข้าใจง่าย' },
    { id: 'q2', section: 'กระบวนการให้บริการ', text: 'มีช่องทางการติดต่อที่สะดวก เหมาะสม เช่น โทรศัพท์ เว็บไซต์ เครือข่ายสังคมออนไลน์' },
    { id: 'q3', section: 'กระบวนการให้บริการ', text: 'ระยะเวลาในการให้บริการเหมาะสม' },
    { id: 'q4', section: 'เจ้าหน้าที่ผู้ให้บริการ', text: 'เจ้าหน้าที่ผู้ให้บริการมีบุคลิกภาพที่ดี และใส่ใจในการให้บริการ' },
    { id: 'q5', section: 'เจ้าหน้าที่ผู้ให้บริการ', text: 'เจ้าหน้าที่ผู้ให้บริการให้คำแนะนำและตอบข้อซักถามอย่างชัดเจน' },
    { id: 'q6', section: 'เจ้าหน้าที่ผู้ให้บริการ', text: 'เจ้าหน้าที่ผู้ให้บริการมีความรู้ความสามารถในการให้บริการ' },
    { id: 'q7', section: 'การประชาสัมพันธ์', text: 'สื่อประชาสัมพันธ์ เช่น โบรชัวร์ มีการแจ้งรายละเอียดข้อมูลของเครื่องมืออย่างชัดเจนและครอบคลุม' },
    { id: 'q8', section: 'การประชาสัมพันธ์', text: 'มีการประชาสัมพันธ์หลากหลายช่องทาง' },
    { id: 'q9', section: 'สถานที่และสิ่งอำนวยความสะดวก', text: 'สถานที่ให้บริการมีความเหมาะสมและเข้าถึงได้สะดวก' },
    { id: 'q10', section: 'สถานที่และสิ่งอำนวยความสะดวก', text: 'สถานที่รอรับบริการมีสิ่งอำนวยความสะดวกเพียงพอ เช่น ที่นั่ง' },
    { id: 'q11', section: 'คุณภาพการให้บริการ', text: 'ผลิตภัณฑ์หรือบริการของเราตอบสนองต่อความคาดหวังของคุณได้ดีเพียงใด' },
  ];
  const CHANNEL_OPTIONS = ['Website', 'Facebook', 'แผ่นพับโบรชัวร์ (Brochure)', 'การแนะนำจากคนรู้จัก', 'อื่น ๆ'];
  const TOOL_OPTIONS = ['Fermentor 10 L', 'Fermentor 70 L', 'Fermentor 750 L', 'Continuous Centrifuge', 'Freeze Dryer', 'Ultrafiltration'];
  const initScores = () => Object.fromEntries(SURVEY_QUESTIONS.map(q => [q.id, 5]));
  const [feedbackScores, setFeedbackScores] = useState(initScores);
  const [feedbackChannels, setFeedbackChannels] = useState([]);
  const [feedbackTools, setFeedbackTools] = useState([]);
  const [feedbackSuggestion, setFeedbackSuggestion] = useState('');


  const [activeTab, setActiveTab] = useState('diagram'); // 'diagram' | 'dashboard' | 'combined' | 'table' | 'ai'
  // AI Assistant States
  const [aiReport, setAiReport] = useState('');
  const [isAiReportLoading, setIsAiReportLoading] = useState(false);
  const [aiChatMessages, setAiChatMessages] = useState([
    { role: 'model', content: 'สวัสดีครับ ผมคือผู้ช่วย AI วิเคราะห์ข้อมูลกระบวนการทางชีวภาพ (Bioprocess AI Co-pilot) คุณสามารถคลิกให้ผมช่วยวิเคราะห์แนวโน้มและจุดผิดปกติ หรือพิมพ์สอบถามเกี่ยวกับข้อมูลรันนี้ได้เลยครับ!' }
  ]);
  const [aiNewMessage, setAiNewMessage] = useState('');
  const [isAiChatLoading, setIsAiChatLoading] = useState(false);
  const [isAiReportMock, setIsAiReportMock] = useState(false);
  const [isAiChatMock, setIsAiChatMock] = useState(false);

  const handleSendChipMessage = async (text) => {
    await sendChatMessage(text);
  };

  const sendChatMessage = async (text) => {
    const updatedMessages = [...aiChatMessages, { role: 'user', content: text }];
    setAiChatMessages(updatedMessages);
    setIsAiChatLoading(true);

    try {
      const res = await fetch('/api/ai/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          jobId: currentJobId,
          messages: updatedMessages.slice(1), // Exclude greeting
          newMessage: text
        })
      });

      if (res.ok) {
        const data = await res.json();
        setAiChatMessages(prev => [...prev, { role: 'model', content: data.response }]);
        setIsAiChatMock(data.isMock);
      } else {
        const errData = await res.json().catch(() => ({}));
        const errMsg = errData.error ? ` (${errData.error})` : '';
        setAiChatMessages(prev => [...prev, { role: 'model', content: `ขออภัยครับ ไม่สามารถสร้างคำตอบได้ในขณะนี้เนื่องจากปัญหาทางระบบ${errMsg}` }]);
      }
    } catch (e) {
      console.error(e);
      setAiChatMessages(prev => [...prev, { role: 'model', content: 'ขออภัยครับ เกิดข้อผิดพลาดในการเชื่อมต่อเซิร์ฟเวอร์' }]);
    } finally {
      setIsAiChatLoading(false);
    }
  };

  // Theme: 'dark' | 'light'
  const [theme, setTheme] = useState(() => {
    return localStorage.getItem('bioprocess-theme') || 'dark';
  });
  // Customer notice modal state
  const [showCustomerNotice, setShowCustomerNotice] = useState(false);
  const [pendingJobCode, setPendingJobCode] = useState(null);
  const [dontShowAgain, setDontShowAgain] = useState(() => {
    try { return localStorage.getItem('bioprocess-customer-notice-ack') === 'true'; } catch (e) { return false; }
  });
  const [showCustomerBanner, setShowCustomerBanner] = useState(false);
  // Customer Login Warning Toast State
  const [showCustomerToast, setShowCustomerToast] = useState(false);
  const [isToastHiding, setIsToastHiding] = useState(false);
  // Replay recorded data (playback) state
  const [isReplay, setIsReplay] = useState(false);
  const [isReplayPlaying, setIsReplayPlaying] = useState(false);
  const [replayIndex, setReplayIndex] = useState(1);
  const [replaySpeed, setReplaySpeed] = useState(1500); // default speed (1x = 1.5s per point)

  // Combined Jobs view states
  const [combinedActiveTab, setCombinedActiveTab] = useState('list'); // 'list' | 'compare'
  const [selectedCompareJobIds, setSelectedCompareJobIds] = useState([]);
  const [compareParams, setCompareParams] = useState(['temp_read']);
  const COMPARE_PARAM_OPTIONS = [
    { value: 'temp_read', label: 'อุณหภูมิ PV (°C)' },
    { value: 'ph_read', label: 'ค่า pH PV' },
    { value: 'do_read', label: 'ค่า DO PV (%)' },
    { value: 'agit_read', label: 'ความเร็วการกวน AGIT (RPM)' },
    { value: 'air_read', label: 'อัตราไหลลม AIR (L/M)' },
  ];
  const [combinedSearchQuery, setCombinedSearchQuery] = useState('');


  const compareChartRef = useRef(null);

  // Export chart panel as PNG image
  const downloadChartImage = async () => {
    const el = compareChartRef.current;
    if (!el) return;
    try {
      const canvas = await html2canvas(el, { backgroundColor: '#0f172a', scale: 2, useCORS: true });
      const link = document.createElement('a');
      link.download = 'compare_chart_' + new Date().toISOString().slice(0, 10) + '.png';
      link.href = canvas.toDataURL('image/png');
      link.click();
    } catch (e) { console.error('Export image error', e); }
  };

  // Export chart panel as PDF
  const downloadChartPDF = async () => {
    const el = compareChartRef.current;
    if (!el) return;
    try {
      const canvas = await html2canvas(el, { backgroundColor: '#0f172a', scale: 2, useCORS: true });
      const imgData = canvas.toDataURL('image/png');
      const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
      const pdfW = pdf.internal.pageSize.getWidth();
      const pdfH = (canvas.height / canvas.width) * pdfW;
      pdf.addImage(imgData, 'PNG', 0, 0, pdfW, pdfH);
      pdf.save('compare_chart_' + new Date().toISOString().slice(0, 10) + '.pdf');
    } catch (e) { console.error('Export PDF error', e); }
  };
  // Download compare chart data as CSV
  const downloadCompareCSV = () => {
    if (selectedCompareJobIds.length === 0) return;
    const roundedPoints = {};
    const columnKeys = [];
    selectedCompareJobIds.forEach(jobId => {
      const job = jobs.find(j => j.id === jobId);
      if (!job) return;
      const machine = machines.find(m => m.id === job.machineId);
      const mPrefix = machine ? machine.name + ' - ' : '';
      const jobLabel = mPrefix + job.name;
      let minTimeMs = Infinity;
      if (job.data && job.data.length > 0) {
        job.data.forEach(row => {
          if (row.timestamp) {
            const t = new Date(row.timestamp).getTime();
            if (!isNaN(t) && t < minTimeMs) minTimeMs = t;
          }
        });
      }
      compareParams.forEach(param => {
        const paramLabel = COMPARE_PARAM_OPTIONS.find(o => o.value === param)?.label || param;
        const lineKey = compareParams.length > 1 ? jobLabel + ' [' + paramLabel + ']' : jobLabel;
        if (!columnKeys.includes(lineKey)) columnKeys.push(lineKey);
        if (job.data) {
          job.data.forEach(row => {
            const val = row[param];
            if (val === undefined || isNaN(parseFloat(val))) return;
            const t = row.timestamp ? new Date(row.timestamp).getTime() : NaN;
            const elapsed = isNaN(t) || minTimeMs === Infinity ? 0 : (t - minTimeMs) / 3600000;
            const rounded = Math.round(elapsed * 10) / 10;
            if (!roundedPoints[rounded]) roundedPoints[rounded] = { cultureHour: rounded };
            roundedPoints[rounded][lineKey] = parseFloat(val);
          });
        }
      });
    });
    const rows = Object.values(roundedPoints).sort((a, b) => a.cultureHour - b.cultureHour);
    const header = ['Culture Hour (ชม.)', ...columnKeys];
    const csvLines = [header.join(',')];
    rows.forEach(row => {
      const vals = [row.cultureHour, ...columnKeys.map(k => row[k] !== undefined ? row[k] : '')];
      csvLines.push(vals.join(','));
    });
    const bom = '\uFEFF';
    const blob = new Blob([bom + csvLines.join('\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'compare_chart_' + new Date().toISOString().slice(0, 10) + '.csv';
    a.click();
    URL.revokeObjectURL(url);
  };

  const [formData, setFormData] = useState({
    temp_set: 38.0, temp_read: 38.0,
    ph_set: 7.00, ph_read: 7.00,
    do_set: 50, do_read: 50,
    agit_set: 200, agit_read: 200,
    air_set: 2.0, air_read: 2.0,
    level_set: 65.0, level_read: 65.0,
    air_out_set: 1.9, air_out_read: 1.9,
    heat_set: 50, heat_read: 50,
    remark: '',
    date: new Date().toLocaleDateString('en-CA'),
    time: toHHMM(new Date())
  });

  const [visibleParameters, setVisibleParameters] = useState({
    temp: true,
    ph: true,
    do: true,
    agit: true,
    air: true,
    level: true,
    air_out: true,
    heat: true
  });

  const [customerFormData, setCustomerFormData] = useState({
    companyName: '',
    machineId: '',
    email: ''
  });

  const [selectedCustomerJobs, setSelectedCustomerJobs] = useState({});

  const [showCustomerShareModal, setShowCustomerShareModal] = useState(false);
  const [shareModalJobId, setShareModalJobId] = useState(null);
  const [shareInvitationText, setShareInvitationText] = useState('');

  const [showAddSessionModal, setShowAddSessionModal] = useState(false);
  const [newSessionName, setNewSessionName] = useState('');
  const [newSessionMachineId, setNewSessionMachineId] = useState('');

  const [showAddMachineModal, setShowAddMachineModal] = useState(false);
  const [newMachineName, setNewMachineName] = useState('');

  const [newSessionCustomerId, setNewSessionCustomerId] = useState('');
  const [newSessionCustomerName, setNewSessionCustomerName] = useState('');
  const [newSessionCustomerEmail, setNewSessionCustomerEmail] = useState('');
  const [newSessionExpiryHours, setNewSessionExpiryHours] = useState(0);
  const [showWizardSuccess, setShowWizardSuccess] = useState(false);
  const [wizardSuccessJobId, setWizardSuccessJobId] = useState(null);

  // Fetch Database from Backend Helper
  const fetchDB = async (shouldAutoSelect = false) => {
    try {
      const role = userRole || 'guest';
      const params = new URLSearchParams({
        clientId,
        role,
        machineId: currentMachineId || '',
        jobId: currentJobId || ''
      });
      const res = await fetch(`/api/db?${params.toString()}`);
      if (res.ok) {
        const data = await res.json();
        applyDBUpdate(data);
        if (data.activeUsers) {
          setActiveUsers(data.activeUsers);
        }
        setHasFetchedDB(true);

        if (shouldAutoSelect) {
          const savedMachineId = localStorage.getItem('bioprocess-current-machine');
          const savedJobId = localStorage.getItem('bioprocess-current-job');

          const targetMachineId = savedMachineId && data.machines.some(m => m.id === savedMachineId)
            ? savedMachineId
            : (data.machines[0]?.id || null);

          const targetJobId = savedJobId && data.jobs.some(j => j.id === savedJobId)
            ? savedJobId
            : (data.jobs.filter(j => j.machineId === targetMachineId)[0]?.id || null);

          setCurrentMachineId(targetMachineId);
          setCurrentJobId(targetJobId);
        }
      }
    } catch (e) {
      console.error("Error fetching database:", e);
    }
  };

  const handleAutoCustomerLogin = async (jobCode) => {
    try {
      const role = userRole || 'guest';
      const params = new URLSearchParams({
        clientId,
        role,
        machineId: currentMachineId || '',
        jobId: jobCode
      });
      const res = await fetch(`/api/db?${params.toString()}`);
      if (res.ok) {
        const data = await res.json();
        const targetJob = data.jobs.find(j => j.id === jobCode);
        if (!targetJob) {
          alert('ไม่พบรหัสงานนี้ในระบบ หรือ ลิงก์ไม่ถูกต้อง');
          setUserRole(null);
          setActiveCustomerJobId(null);
          return;
        }

        const isExpired = targetJob.expiresAt && new Date() > new Date(targetJob.expiresAt);
        const isFinished = targetJob.status === 'finished';
        if ((isExpired || isFinished) && !targetJob.allowHistoryView) {
          alert('สิทธิ์การเข้าใช้งานหมดอายุหรือเสร็จสิ้นแล้ว (หากต้องการดูย้อนหลัง กรุณาติดต่อเจ้าหน้าที่เพื่อขอเปิดสิทธิ์)');
          setUserRole(null);
          setActiveCustomerJobId(null);
          return;
        }

        // Apply fresh DB
        applyDBUpdate(data);
        if (data.activeUsers) {
          setActiveUsers(data.activeUsers);
        }

        const ack = localStorage.getItem('bioprocess-customer-notice-ack') === 'true';
        if (ack) {
          setActiveCustomerJobId(jobCode);
          setUserRole('customer');
          setCurrentAppView('monitoring');
        } else {
          setPendingJobCode(jobCode);
          setShowCustomerNotice(true);
        }
      }
    } catch (err) {
      console.error("Error auto logging in customer:", err);
    }
  };

  useEffect(() => {
    fetchDB(true);

    // Fetch public settings / about info
    fetch('/api/settings')
      .then(res => {
        if (res.ok) return res.json();
        throw new Error('Failed to fetch settings');
      })
      .then(data => {
        if (data) {
          setAboutSystem({
            systemName: data.systemName || 'DBMS (Bioprocess Data Logging)',
            systemVersion: data.systemVersion || 'v2.4.0 (SCADA Polish)',
            developer: data.developer || 'ทีมวิศวกรรมข้อมูลชีวภาพ (Bioprocess Engineering Team)',
            techStack: data.techStack || 'React / Vite / Node.js / GCS',
            supportEmail: data.supportEmail || 'support@bioprocess-logging.local',
            supportPhone: data.supportPhone || '+66 2 123 4567',
            vvmCalcType: data.vvmCalcType || 'dynamic',
            maxVolumeLiters: data.maxVolumeLiters !== undefined ? data.maxVolumeLiters : 5.0,
            constantVolumeLiters: data.constantVolumeLiters !== undefined ? data.constantVolumeLiters : 3.5,
            airUnit: data.airUnit || 'mlmin'
          });
        }
      })
      .catch(err => console.error('Error fetching about system settings:', err));

    const params = new URLSearchParams(window.location.search);
    const urlJobId = params.get('job');
    if (urlJobId) {
      window.history.replaceState({}, document.title, window.location.pathname);
      handleAutoCustomerLogin(urlJobId);
    }
  }, []);

  // Poll server for real-time synchronization across multiple users
  useEffect(() => {
    const interval = setInterval(() => {
      fetchDB(false);
    }, 5000);
    return () => clearInterval(interval);
  }, [currentMachineId, currentJobId, userRole]);

  // Active Customer Session Expiry Check
  useEffect(() => {
    if (hasFetchedDB && userRole === 'customer' && activeCustomerJobId) {
      const activeJob = jobs.find(j => j.id === activeCustomerJobId);

      let shouldLogout = false;
      let reason = '';

      if (!activeJob) {
        shouldLogout = true;
        reason = 'ไม่พบข้อมูลของรอบรันนี้ในระบบ หรือข้อมูลอาจถูกลบไปแล้ว';
      } else {
        const isExpired = activeJob.expiresAt && new Date() > new Date(activeJob.expiresAt);
        const isFinished = activeJob.status === 'finished';
        if ((isExpired || isFinished) && !activeJob.allowHistoryView) {
          shouldLogout = true;
          reason = `สิทธิ์การเข้าใช้งานสำหรับรอบรัน "${activeJob.name || activeCustomerJobId}" หมดอายุหรือเสร็จสิ้นแล้ว และไม่ได้รับอนุญาตให้ดูย้อนหลัง`;
        }
      }

      if (shouldLogout) {
        setSessionExpiredMessage(reason + ' กรุณาติดต่อเจ้าหน้าที่เพื่อขออนุญาตหรือขอลิงก์ใหม่');
        setUserRole(null);
        setActiveCustomerJobId(null);
        setCurrentAppView('monitoring');
      }
    }
  }, [hasFetchedDB, jobs, userRole, activeCustomerJobId]);

  // Save current active selections locally in browser
  useEffect(() => {
    if (currentMachineId) {
      localStorage.setItem('bioprocess-current-machine', currentMachineId);
    }
  }, [currentMachineId]);

  useEffect(() => {
    if (currentJobId) {
      localStorage.setItem('bioprocess-current-job', currentJobId);
    }
  }, [currentJobId]);

  // Apply theme to document and persist
  useEffect(() => {
    try {
      document.documentElement.setAttribute('data-theme', theme);
      localStorage.setItem('bioprocess-theme', theme);
    } catch (e) {
      // server-side or non-browser env ignore
    }
  }, [theme]);

  useEffect(() => {
    try { localStorage.setItem('bioprocess-customer-notice-ack', dontShowAgain ? 'true' : 'false'); } catch (e) { }
  }, [dontShowAgain]);

  // Ensure customer form has a valid machineId selected
  useEffect(() => {
    if (!customerFormData.machineId && machines.length > 0) {
      setCustomerFormData(prev => ({ ...prev, machineId: machines[0].id }));
    }
  }, [machines, customerFormData.machineId]);

  // Derived state
  const jobsForMachine = jobs.filter(j => j.machineId === currentMachineId);
  const currentMachine = machines.find(m => m.id === currentMachineId);

  const getDynamicMaxVolForMachine = (mId) => {
    const machine = machines.find(m => m.id === mId);
    if (!machine) return aboutSystem?.maxVolumeLiters !== undefined ? Number(aboutSystem.maxVolumeLiters) : 5.0;
    const name = machine.name.toLowerCase();
    if (name.includes('750') || name.includes('500')) return 750.0;
    if (name.includes('70')) return 70.0;
    if (name.includes('10')) return 10.0;
    return aboutSystem?.maxVolumeLiters !== undefined ? Number(aboutSystem.maxVolumeLiters) : 5.0;
  };

  const getDynamicMaxVol = () => {
    if (!currentMachine) return aboutSystem?.maxVolumeLiters !== undefined ? Number(aboutSystem.maxVolumeLiters) : 5.0;
    const name = currentMachine.name.toLowerCase();
    if (name.includes('750') || name.includes('500')) return 750.0;
    if (name.includes('70')) return 70.0;
    if (name.includes('10')) return 10.0;
    return aboutSystem?.maxVolumeLiters !== undefined ? Number(aboutSystem.maxVolumeLiters) : 5.0;
  };
  const maxVol = getDynamicMaxVol();
  const calcVolumeLiters = (percentVal) => {
    if (percentVal === undefined || percentVal === null || typeof percentVal !== 'number') return 0;
    return parseFloat(((percentVal / 100) * maxVol).toFixed(1));
  };

  const mapDataPointsToLiters = (dataPoints, job) => {
    if (!job) return [];
    const jobMaxVol = getDynamicMaxVolForMachine(job.machineId);
    const calcVol = (percentVal) => {
      if (percentVal === undefined || percentVal === null || typeof percentVal !== 'number') return 0;
      return parseFloat(((percentVal / 100) * jobMaxVol).toFixed(1));
    };
    return dataPoints.map((row, idx) => {
      const airSetVal = row.air_set !== undefined ? row.air_set : row.air;
      const airReadVal = row.air_read !== undefined ? row.air_read : row.air;
      return {
        ...row,
        originalIndex: idx,
        cultureHour: getElapsedHours(job, row.timestamp),
        temp_read: row.temp_read !== undefined ? row.temp_read : row.temp,
        temp_set: row.temp_set !== undefined ? row.temp_set : row.temp,
        ph_read: row.ph_read !== undefined ? row.ph_read : row.ph,
        ph_set: row.ph_set !== undefined ? row.ph_set : row.ph,
        do_read: row.do_read !== undefined ? row.do_read : row.do,
        do_set: row.do_set !== undefined ? row.do_set : row.do,
        agit_read: row.agit_read !== undefined ? row.agit_read : row.agit,
        agit_set: row.agit_set !== undefined ? row.agit_set : row.agit,
        air_read: airReadVal,
        air_set: airSetVal,
        level_set: row.level_set !== undefined && row.level_set !== null ? calcVol(row.level_set) : calcVol(65.0),
        level_read: row.level_read !== undefined && row.level_read !== null ? calcVol(row.level_read) : calcVol(65.0),
        air_out_set: row.air_out_set !== undefined && row.air_out_set !== null ? row.air_out_set : parseFloat(((airSetVal || 0) * 0.96).toFixed(2)),
        air_out_read: row.air_out_read !== undefined && row.air_out_read !== null ? row.air_out_read : parseFloat(((airReadVal || 0) * 0.96).toFixed(2)),
        heat_set: row.heat_set !== undefined && row.heat_set !== null ? row.heat_set : 0.0,
        heat_read: row.heat_read !== undefined && row.heat_read !== null ? row.heat_read : 0.0,
        remark: row.remark !== undefined ? row.remark : ''
      };
    });
  };

  const currentJob = jobs.find(j => j.id === currentJobId) || jobsForMachine[0];
  const currentJobData = currentJob?.data || [];

  // Sort full data chronologically
  const sortedFullData = React.useMemo(() => {
    return [...currentJobData].sort((a, b) => {
      const aTime = a.timestamp ? new Date(a.timestamp).getTime() : 0;
      const bTime = b.timestamp ? new Date(b.timestamp).getTime() : 0;
      return aTime - bTime;
    });
  }, [currentJobData]);

  // Choose which data set to render: live/full stored data or replayed slice
  const displayData = React.useMemo(() => {
    if (isReplay) {
      return sortedFullData.slice(0, replayIndex);
    }
    return currentJobData;
  }, [isReplay, sortedFullData, replayIndex, currentJobData]);

  const chartData = mapDataPointsToLiters(displayData, currentJob).sort((a, b) => {
    const aTime = a.timestamp ? new Date(a.timestamp).getTime() : 0;
    const bTime = b.timestamp ? new Date(b.timestamp).getTime() : 0;
    return aTime - bTime;
  });
  // Sorting state for table
  const [sortField, setSortField] = useState('timestamp'); // default sort by timestamp
  const [sortAsc, setSortAsc] = useState(false);

  const toggleSort = (field) => {
    if (sortField === field) {
      setSortAsc(!sortAsc);
    } else {
      setSortField(field);
      setSortAsc(true);
    }
  };

  const getSortedRows = () => {
    const rows = [...chartData];
    rows.sort((a, b) => {
      if (sortField === 'date' || sortField === 'time' || sortField === 'timestamp' || sortField === 'cultureHour') {
        const aTime = a.timestamp ? new Date(a.timestamp).getTime() : 0;
        const bTime = b.timestamp ? new Date(b.timestamp).getTime() : 0;
        return sortAsc ? aTime - bTime : bTime - aTime;
      }
      const aVal = a[sortField] !== undefined ? a[sortField] : '';
      const bVal = b[sortField] !== undefined ? b[sortField] : '';
      if (aVal < bVal) return sortAsc ? -1 : 1;
      if (aVal > bVal) return sortAsc ? 1 : -1;
      return 0;
    });
    return rows;
  };
  const getEditingRowCultureHour = () => {
    if (!editingRowData) return 0;
    const { date, time } = editingRowData;
    if (date && time) {
      const dt = new Date(`${date}T${time}`);
      if (!isNaN(dt.getTime())) {
        return getElapsedHours(currentJob, dt.toISOString());
      }
      const dt2 = new Date(`${date} ${time}`);
      if (!isNaN(dt2.getTime())) {
        return getElapsedHours(currentJob, dt2.toISOString());
      }
    }
    return getElapsedHours(currentJob, editingRowData.timestamp);
  };
  const lastDataPoint = chartData.length > 0 ? chartData[chartData.length - 1] : null;
  const showStandby = currentJob?.status === 'finished' && !isViewingHistory && userRole !== 'customer';
  const lastDataPointForDisplay = (showStandby || !lastDataPoint) ? {
    temp_set: 0, temp_read: 0, temp: 0,
    ph_set: 0, ph_read: 0, ph: 0,
    do_set: 0, do_read: 0, do: 0,
    agit_set: 0, agit_read: 0, agit: 0,
    air_set: 0, air_read: 0, air: 0,
    level_set: 0, level_read: 0,
    air_out_set: 0, air_out_read: 0,
    heat_set: 0, heat_read: 0,
    remark: '',
    date: toYYYYMMDD(new Date()),
    time: toHHMM(new Date())
  } : lastDataPoint;

  // Enforce customer constraints: lock into active customer job and machine
  useEffect(() => {
    if (userRole === 'customer' && activeCustomerJobId) {
      setCurrentJobId(activeCustomerJobId);
      const targetJob = jobs.find(j => j.id === activeCustomerJobId);
      if (targetJob && targetJob.machineId) {
        setCurrentMachineId(targetJob.machineId);
        const isExpired = targetJob.expiresAt && new Date() > new Date(targetJob.expiresAt);
        const isFinished = targetJob.status === 'finished';
        setIsViewingHistory(isFinished || isExpired);
      }
    }
  }, [userRole, activeCustomerJobId, jobs]);

  // Show transient banner for customers on login (unless they've chosen don't-show-again)
  useEffect(() => {
    if (userRole === 'customer') {
      try {
        const ack = localStorage.getItem('bioprocess-customer-notice-ack') === 'true';
        if (!ack) {
          setShowCustomerBanner(true);
          const t = setTimeout(() => setShowCustomerBanner(false), 6000);
          return () => clearTimeout(t);
        }
      } catch (e) {
        // ignore
      }
    }
    return undefined;
  }, [userRole, activeCustomerJobId]);

  // Customer Warning Toast Effect
  useEffect(() => {
    if (userRole === 'customer') {
      setShowCustomerToast(true);
      setIsToastHiding(false);

      const hideTimer = setTimeout(() => {
        setIsToastHiding(true);
      }, 14400);

      const closeTimer = setTimeout(() => {
        setShowCustomerToast(false);
        setIsToastHiding(false);
      }, 15000);

      return () => {
        clearTimeout(hideTimer);
        clearTimeout(closeTimer);
      };
    } else {
      setShowCustomerToast(false);
      setIsToastHiding(false);
    }
  }, [userRole]);

  const dismissCustomerToast = () => {
    setIsToastHiding(true);
    setTimeout(() => {
      setShowCustomerToast(false);
      setIsToastHiding(false);
    }, 600);
  };

  // Reset replay when job changes
  useEffect(() => {
    if (!currentJobId) {
      setIsReplay(false);
      setIsReplayPlaying(false);
      setReplayIndex(1);
    } else {
      // If replay active but job changed, restart replay
      if (isReplay) {
        setReplayIndex(1);
      }
    }
  }, [currentJobId]);

  useEffect(() => {
    if (currentJob && currentJob.machineId !== currentMachineId) {
      const firstJobForMachine = jobsForMachine[0];
      setCurrentJobId(firstJobForMachine ? firstJobForMachine.id : null);
    }
  }, [currentMachineId, currentJob, jobsForMachine]);

  // Pre-populate manual entry form with the last data point of the active session
  useEffect(() => {
    if (showStandby) {
      setFormData({
        temp_set: 0, temp_read: 0,
        ph_set: 0, ph_read: 0,
        do_set: 0, do_read: 0,
        agit_set: 0, agit_read: 0,
        air_set: 0, air_read: 0,
        level_set: 0, level_read: 0,
        air_out_set: 0, air_out_read: 0,
        heat_set: 0, heat_read: 0,
        remark: '',
        date: toYYYYMMDD(new Date()),
        time: toHHMM(new Date())
      });
      return;
    }

    // determine last timestamp -> date and time fields
    const lastTs = lastDataPoint && (lastDataPoint.timestamp || lastDataPoint.time) ? (lastDataPoint.timestamp || lastDataPoint.time) : null;
    const lastDate = lastTs ? new Date(lastTs) : null;
    const lastDateValid = lastDate && !isNaN(lastDate.getTime());
    if (lastDataPoint) {
      setFormData({
        temp_set: lastDataPoint.temp_set,
        temp_read: lastDataPoint.temp_read,
        ph_set: lastDataPoint.ph_set,
        ph_read: lastDataPoint.ph_read,
        do_set: lastDataPoint.do_set,
        do_read: lastDataPoint.do_read,
        agit_set: lastDataPoint.agit_set,
        agit_read: lastDataPoint.agit_read,
        air_set: lastDataPoint.air_set,
        air_read: lastDataPoint.air_read,
        level_set: lastDataPoint.level_set !== undefined && lastDataPoint.level_set !== null ? (lastDataPoint.level_set / maxVol) * 100 : 65.0,
        level_read: lastDataPoint.level_read !== undefined && lastDataPoint.level_read !== null ? (lastDataPoint.level_read / maxVol) * 100 : 65.0,
        air_out_set: lastDataPoint.air_out_set !== undefined && lastDataPoint.air_out_set !== null ? lastDataPoint.air_out_set : parseFloat(((lastDataPoint.air_set || 0) * 0.96).toFixed(2)),
        air_out_read: lastDataPoint.air_out_read !== undefined && lastDataPoint.air_out_read !== null ? lastDataPoint.air_out_read : parseFloat(((lastDataPoint.air_read || 0) * 0.96).toFixed(2)),
        heat_set: lastDataPoint.heat_set !== undefined && lastDataPoint.heat_set !== null ? lastDataPoint.heat_set : 0,
        heat_read: lastDataPoint.heat_read !== undefined && lastDataPoint.heat_read !== null ? lastDataPoint.heat_read : 0,
        remark: '',
        date: new Date().toLocaleDateString('en-CA'),
        time: toHHMM(new Date())
      });
    } else {
      setFormData({
        temp_set: 37.0,
        temp_read: 37.0,
        ph_set: 7.00,
        ph_read: 7.00,
        do_set: 50,
        do_read: 50,
        agit_set: 200,
        agit_read: 200,
        air_set: 2.0,
        air_read: 2.0,
        level_set: 65.0,
        level_read: 65.0,
        air_out_set: 1.9,
        air_out_read: 1.9,
        heat_set: 0,
        heat_read: 0,
        remark: '',
        date: toYYYYMMDD(new Date()),
        time: toHHMM(new Date())
      });
    }
  }, [currentJobId, showStandby, lastDataPoint === null]);

  // Helper to apply updated DB state
  const applyDBUpdate = (data) => {
    setMachines(data.machines);
    setJobs(data.jobs);
    setCustomers(data.customers);
    if (data.feedbacks) {
      setFeedbacks(data.feedbacks);
    }
  };

  const handleMachineChange = (e) => {
    const value = e.target.value;
    if (value === 'ADD_NEW') {
      setNewMachineName('');
      setShowAddMachineModal(true);
    } else {
      setCurrentMachineId(value);
    }
  };

  const submitNewMachine = async (e) => {
    if (e) e.preventDefault();
    if (!newMachineName.trim()) {
      alert("กรุณากรอกชื่อเครื่องมือ");
      return;
    }
    try {
      const res = await fetch('/api/machines', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: newMachineName.trim() })
      });
      if (res.ok) {
        const data = await res.json();
        applyDBUpdate(data);

        const newMachine = data.machines[data.machines.length - 1];
        if (newMachine) {
          setCurrentMachineId(newMachine.id);
          const jobRes = await fetch('/api/jobs', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ machineId: newMachine.id, name: 'Session 1' })
          });
          if (jobRes.ok) {
            const jobData = await jobRes.json();
            applyDBUpdate(jobData);
            const newJob = jobData.jobs.find(j => j.machineId === newMachine.id);
            if (newJob) {
              setCurrentJobId(newJob.id);
            }
          }
        }
        setShowAddMachineModal(false);
        setNewMachineName('');
      } else {
        alert("ไม่สามารถสร้างเครื่องมือใหม่ได้");
      }
    } catch (err) {
      console.error(err);
      alert("เกิดข้อผิดพลาดในการเชื่อมต่อเซิร์ฟเวอร์");
    }
  };

  const renameMachine = () => {
    const activeMachine = machines.find(m => m.id === currentMachineId);
    if (!activeMachine) return;
    setEditingMachineId(activeMachine.id);
    setEditingMachineName(activeMachine.name);
    setShowEditMachineModal(true);
  };

  const submitEditMachine = async (e) => {
    if (e) e.preventDefault();
    if (!editingMachineName.trim()) {
      alert("กรุณากรอกชื่อเครื่องมือ");
      return;
    }
    try {
      const res = await fetch(`/api/machines/${editingMachineId}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: editingMachineName.trim() })
      });
      if (res.ok) {
        const data = await res.json();
        applyDBUpdate(data);
        setShowEditMachineModal(false);
      } else {
        alert("ไม่สามารถแก้ไขชื่อเครื่องมือได้");
      }
    } catch (err) {
      console.error(err);
      alert("เกิดข้อผิดพลาดในการเชื่อมต่อเซิร์ฟเวอร์");
    }
  };

  const deleteMachine = async () => {
    const activeMachine = machines.find(m => m.id === currentMachineId);
    if (!activeMachine) return;

    if (machines.length <= 1) {
      alert("You must keep at least one machine.");
      return;
    }

    if (window.confirm(`Are you sure you want to delete "${activeMachine.name}"? All its sessions and data will be permanently deleted.`)) {
      try {
        const res = await fetch(`/api/machines/${currentMachineId}`, {
          method: 'DELETE'
        });
        if (res.ok) {
          const data = await res.json();
          applyDBUpdate(data);

          const nextMachineId = data.machines[0]?.id || null;
          setCurrentMachineId(nextMachineId);

          const nextMachineJobs = data.jobs.filter(j => j.machineId === nextMachineId);
          setCurrentJobId(nextMachineJobs.length > 0 ? nextMachineJobs[0].id : null);
        }
      } catch (err) {
        console.error(err);
      }
    }
  };

  const createNewJob = () => {
    if (!currentMachineId) return alert("Please select or create a machine first.");
    setNewSessionName(`Session ${jobsForMachine.length + 1}`);
    setNewSessionMachineId(currentMachineId);
    setShowAddSessionModal(true);
  };

  const submitNewJob = async (e) => {
    if (e) e.preventDefault();
    if (!newSessionName.trim()) {
      alert("กรุณากรอกชื่อรอบบันทึก");
      return;
    }
    if (!newSessionMachineId) {
      alert("กรุณาเลือกเครื่องมือ");
      return;
    }
    // Validate new customer name if ADD_NEW selected
    if (newSessionCustomerId === 'ADD_NEW' && !newSessionCustomerName.trim()) {
      alert("กรุณากรอกชื่อลูกค้าใหม่");
      return;
    }
    try {
      // Step 1: Create the job
      const res = await fetch('/api/jobs', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ machineId: newSessionMachineId, name: newSessionName.trim() })
      });
      if (!res.ok) {
        alert("ไม่สามารถสร้างงานใหม่ได้ กรุณาลองใหม่อีกครั้ง");
        return;
      }
      let data = await res.json();
      applyDBUpdate(data);

      // Find the newly created job
      const newJob = data.jobs.find(j => j.machineId === newSessionMachineId && j.name === newSessionName.trim());
      if (!newJob) {
        alert("ไม่พบงานที่สร้างใหม่ กรุณาลองใหม่อีกครั้ง");
        return;
      }

      // Step 2: Set expiry if specified
      if (newSessionExpiryHours > 0) {
        const expiresAt = new Date(Date.now() + newSessionExpiryHours * 60 * 60 * 1000).toISOString();
        try {
          const expiryRes = await fetch(`/api/jobs/${newJob.id}/expiry`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ expiresAt })
          });
          if (expiryRes.ok) {
            data = await expiryRes.json();
            applyDBUpdate(data);
          }
        } catch (expiryErr) {
          console.error("Error setting expiry:", expiryErr);
        }
      }

      // Step 3: Create new customer if ADD_NEW selected
      if (newSessionCustomerId === 'ADD_NEW' && newSessionCustomerName.trim()) {
        try {
          const custRes = await fetch('/api/customers', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              companyName: newSessionCustomerName.trim(),
              machineId: newSessionMachineId,
              email: newSessionCustomerEmail.trim() || undefined
            })
          });
          if (custRes.ok) {
            data = await custRes.json();
            applyDBUpdate(data);
          }
        } catch (custErr) {
          console.error("Error creating customer:", custErr);
        }
      }

      // Step 4: Update machine selection and show success screen
      if (newSessionMachineId !== currentMachineId) {
        setCurrentMachineId(newSessionMachineId);
      }
      setCurrentJobId(newJob.id);
      setIsViewingHistory(false);

      // Show success screen inside wizard
      setWizardSuccessJobId(newJob.id);
      setShowWizardSuccess(true);

    } catch (err) {
      console.error("Error creating job:", err);
      alert("เกิดข้อผิดพลาดในการเชื่อมต่อเซิร์ฟเวอร์");
    }
  };

  const deleteJob = async (id, e) => {
    e.stopPropagation();
    if (window.confirm("Are you sure you want to delete this session? All its data will be lost.")) {
      try {
        const res = await fetch(`/api/jobs/${id}`, {
          method: 'DELETE'
        });
        if (res.ok) {
          const data = await res.json();
          applyDBUpdate(data);
          if (currentJobId === id) {
            const remainingForMachine = data.jobs.filter(j => j.machineId === currentMachineId);
            setCurrentJobId(remainingForMachine.length > 0 ? remainingForMachine[0].id : null);
          }
        }
      } catch (err) {
        console.error(err);
      }
    }
  };

  const handleToggleJobStatus = async (jobId, currentStatus) => {
    // ปุ่มนี้คือ "เสร็จสิ้นงาน" — Finish job permanently (cannot restart)
    if (currentStatus === 'finished') return; // already finished
    const confirmMessage = 'ยืนยันการเสร็จสิ้นงานนี้?\n\nเมื่อยืนยันแล้ว งานจะสิ้นสุดถาวรและไม่สามารถเริ่มใหม่ได้';
    if (window.confirm(confirmMessage)) {
      try {
        const res = await fetch(`/api/jobs/${jobId}/finish`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' }
        });
        if (res.ok) {
          const data = await res.json();
          applyDBUpdate(data);
          setIsViewingHistory(false);
        } else {
          const errData = await res.json().catch(() => ({}));
          alert(errData.error || 'ไม่สามารถเสร็จสิ้นงานได้ กรุณาลองใหม่อีกครั้ง');
        }
      } catch (err) {
        console.error('Error finishing job:', err);
        alert('เกิดข้อผิดพลาดในการเชื่อมต่อกับเซิร์ฟเวอร์');
      }
    }
  };

  // Replay (playback) timer logic
  useEffect(() => {
    if (!isReplay || !isReplayPlaying) return undefined;
    if (sortedFullData.length === 0) {
      setIsReplay(false);
      setIsReplayPlaying(false);
      return undefined;
    }

    const interval = setInterval(() => {
      setReplayIndex((ri) => {
        if (ri >= sortedFullData.length) {
          setIsReplayPlaying(false);
          return ri;
        }
        return ri + 1;
      });
    }, replaySpeed);

    return () => clearInterval(interval);
  }, [isReplay, isReplayPlaying, sortedFullData.length, replaySpeed]);

  const deleteDataPoint = async (index) => {
    if (window.confirm("Delete this recorded value?")) {
      try {
        const res = await fetch(`/api/jobs/${currentJobId}/data/${index}`, {
          method: 'DELETE'
        });
        if (res.ok) {
          const data = await res.json();
          applyDBUpdate(data);
        }
      } catch (err) {
        console.error(err);
      }
    }
  };

  // Row Editing States
  const [editingRowIndex, setEditingRowIndex] = useState(null);
  const [editingRowData, setEditingRowData] = useState(null);

  const startEditRow = (index, row) => {
    setEditingRowIndex(index);
    setEditingRowData({
      timestamp: row.timestamp || new Date().toISOString(),
      date: row.timestamp ? toYYYYMMDD(row.timestamp) : (row.date || ''),
      time: row.timestamp ? toHHMM(row.timestamp) : (row.time || ''),
      temp_set: row.temp_set !== undefined ? row.temp_set : 0,
      temp_read: row.temp_read !== undefined ? row.temp_read : 0,
      ph_set: row.ph_set !== undefined ? row.ph_set : 0,
      ph_read: row.ph_read !== undefined ? row.ph_read : 0,
      do_set: row.do_set !== undefined ? row.do_set : 0,
      do_read: row.do_read !== undefined ? row.do_read : 0,
      agit_set: row.agit_set !== undefined ? row.agit_set : 0,
      agit_read: row.agit_read !== undefined ? row.agit_read : 0,
      air_set: row.air_set !== undefined ? row.air_set : 0,
      air_read: row.air_read !== undefined ? row.air_read : 0,
      level_set: row.level_set !== undefined ? (maxVol > 0 ? (row.level_set / maxVol) * 100 : 65.0) : 65.0,
      level_read: row.level_read !== undefined ? (maxVol > 0 ? (row.level_read / maxVol) * 100 : 65.0) : 65.0,
      air_out_set: row.air_out_set !== undefined ? row.air_out_set : parseFloat(((row.air_set || 0) * 0.96).toFixed(2)),
      air_out_read: row.air_out_read !== undefined ? row.air_out_read : parseFloat(((row.air_read || 0) * 0.96).toFixed(2)),
      heat_set: row.heat_set !== undefined ? row.heat_set : 0,
      heat_read: row.heat_read !== undefined ? row.heat_read : 0,
      remark: row.remark !== undefined ? row.remark : ''
    });
  };

  const handleEditChange = (field, value) => {
    setEditingRowData(prev => ({
      ...prev,
      [field]: value
    }));
  };

  const cancelEditRow = () => {
    setEditingRowIndex(null);
    setEditingRowData(null);
  };

  const saveEditRow = async () => {
    if (editingRowIndex === null || !editingRowData) return;
    try {
      const { date, time } = editingRowData;
      let timestamp = editingRowData.timestamp || new Date().toISOString();
      if (date && time) {
        const dateParts = date.split('-');
        const timeParts = time.split(':');
        if (dateParts.length === 3 && timeParts.length >= 2) {
          const year = parseInt(dateParts[0], 10);
          const month = parseInt(dateParts[1], 10) - 1;
          const day = parseInt(dateParts[2], 10);
          const hours = parseInt(timeParts[0], 10);
          const minutes = parseInt(timeParts[1], 10);
          const localDate = new Date(year, month, day, hours, minutes, 0);
          if (!isNaN(localDate.getTime())) {
            timestamp = localDate.toISOString();
          }
        }
      }

      const payload = {
        ...editingRowData,
        timestamp
      };

      const res = await fetch(`/api/jobs/${currentJobId}/data/${editingRowIndex}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      if (res.ok) {
        const data = await res.json();
        applyDBUpdate(data);
        cancelEditRow();
      } else {
        alert("Failed to save changes. Please try again.");
      }
    } catch (err) {
      console.error("Error saving edited row:", err);
      alert("An error occurred while saving.");
    }
  };

  const clearAllData = async () => {
    if (window.confirm("Are you sure you want to delete ALL data in this session?")) {
      try {
        const res = await fetch(`/api/jobs/${currentJobId}/clear`, {
          method: 'POST'
        });
        if (res.ok) {
          const data = await res.json();
          applyDBUpdate(data);
        }
      } catch (err) {
        console.error(err);
      }
    }
  };

  const handleInputChange = (e) => {
    const { name, value } = e.target;
    setFormData(prev => ({
      ...prev,
      [name]: value
    }));
  };

  const handleInputFocus = (e) => {
    // Select all text on focus so user can immediately type a new value
    e.target.select();
    const { name, value } = e.target;
    if (value === '0' || parseFloat(value) === 0) {
      setFormData(prev => ({
        ...prev,
        [name]: ''
      }));
    }
  };

  const handleInputBlur = (e) => {
    const { name, value } = e.target;
    if (String(value).trim() === '') {
      setFormData(prev => ({
        ...prev,
        [name]: 0
      }));
    }
  };

  const handleExportFeedback = () => {
    if (feedbacks.length === 0) {
      alert('ไม่มีข้อมูลแบบประเมินความพึงพอใจสำหรับส่งออก');
      return;
    }

    const headers = [
      'วันที่บันทึก',
      'ชื่องาน',
      'คะแนนเฉลี่ย (CSAT)',
      'q1_ขั้นตอนการให้บริการ',
      'q2_ช่องทางการติดต่อ',
      'q3_ระยะเวลาให้บริการ',
      'q4_บุคลิกภาพเจ้าหน้าที่',
      'q5_คำแนะนำของเจ้าหน้าที่',
      'q6_ความรู้เจ้าหน้าที่',
      'q7_สื่อประชาสัมพันธ์',
      'q8_ช่องทางPR',
      'q9_สถานที่ให้บริการ',
      'q10_สิ่งอำนวยความสะดวก',
      'q11_คุณภาพบริการ',
      'ช่องทางที่รู้จัก',
      'เครื่องมือที่ใช้บริการ',
      'ข้อเสนอแนะเพิ่มเติม'
    ];

    const rows = feedbacks.map(fb => {
      const dateStr = fb.createdAt ? new Date(fb.createdAt).toLocaleString('th-TH') : '';
      const scores = fb.scores || {};
      const channels = (fb.channels || []).join('; ');
      const tools = (fb.tools || []).join('; ');
      const suggestion = fb.suggestion || fb.comment || '';
      const avgScore = fb.avgScore || fb.rating || 0;

      return [
        `"${dateStr}"`,
        `"${fb.jobName || ''}"`,
        avgScore.toFixed(2),
        scores.q1 || '',
        scores.q2 || '',
        scores.q3 || '',
        scores.q4 || '',
        scores.q5 || '',
        scores.q6 || '',
        scores.q7 || '',
        scores.q8 || '',
        scores.q9 || '',
        scores.q10 || '',
        scores.q11 || '',
        `"${channels.replace(/"/g, '""')}"`,
        `"${tools.replace(/"/g, '""')}"`,
        `"${suggestion.replace(/"/g, '""')}"`
      ];
    });

    const csvContent = "\uFEFF" + [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `customer_satisfaction_feedbacks_${new Date().toISOString().slice(0, 10)}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleManualSubmit = (e) => {
    e.preventDefault();
    const now = new Date();
    const inputDate = formData.date || now.toLocaleDateString('en-CA');
    const inputTime = formData.time || toHHMM(now);
    addDataPoint(
      formData.temp_set, formData.temp_read,
      formData.ph_set, formData.ph_read,
      formData.do_set, formData.do_read,
      formData.agit_set, formData.agit_read,
      formData.air_set, formData.air_read,
      formData.level_set, formData.level_read,
      formData.air_out_set, formData.air_out_read,
      formData.heat_set, formData.heat_read,
      formData.remark,
      inputTime,
      inputDate
    );
    // Reset remark and update date/time to now
    const resetNow = new Date();
    setFormData(prev => ({
      ...prev,
      remark: '',
      date: resetNow.toLocaleDateString('en-CA'),
      time: toHHMM(resetNow)
    }));
  };

  const addDataPoint = async (
    temp_set, temp_read,
    ph_set, ph_read,
    do_set, do_read,
    agit_set, agit_read,
    air_set, air_read,
    level_set, level_read,
    air_out_set, air_out_read,
    heat_set, heat_read,
    remark = '',
    time = '',
    date = ''
  ) => {
    if (!currentJobId) return;

    // Construct local timestamp on client side to avoid server-side timezone shifts
    let timestamp = new Date().toISOString();
    if (date && time) {
      const dateParts = date.split('-');
      const timeParts = time.split(':');
      if (dateParts.length === 3 && timeParts.length >= 2) {
        const year = parseInt(dateParts[0], 10);
        const month = parseInt(dateParts[1], 10) - 1;
        const day = parseInt(dateParts[2], 10);
        const hours = parseInt(timeParts[0], 10);
        const minutes = parseInt(timeParts[1], 10);
        const localDate = new Date(year, month, day, hours, minutes, 0);
        if (!isNaN(localDate.getTime())) {
          timestamp = localDate.toISOString();
        }
      }
    }

    try {
      const res = await fetch(`/api/jobs/${currentJobId}/data`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          temp_set, temp_read,
          ph_set, ph_read,
          do_set, do_read,
          agit_set, agit_read,
          air_set, air_read,
          level_set, level_read,
          air_out_set, air_out_read,
          heat_set, heat_read,
          remark,
          time,
          date,
          timestamp
        })
      });
      if (res.ok) {
        const data = await res.json();
        applyDBUpdate(data);
      }
    } catch (err) {
      console.error(err);
    }
  };

  // Customer Management Actions
  const handleCustomerInputChange = (e) => {
    const { name, value } = e.target;
    setCustomerFormData(prev => ({ ...prev, [name]: value }));
  };

  const handleAddCustomer = async (e) => {
    e.preventDefault();
    if (!customerFormData.companyName.trim()) return alert("Please enter a company name.");

    try {
      const res = await fetch('/api/customers', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          companyName: customerFormData.companyName.trim(),
          machineId: customerFormData.machineId,
          email: customerFormData.email.trim()
        })
      });
      if (res.ok) {
        const data = await res.json();
        applyDBUpdate(data);
        setCustomerFormData(prev => ({ ...prev, companyName: '', email: '' }));
      }
    } catch (err) {
      console.error(err);
    }
  };

  const deleteCustomer = async (id) => {
    if (window.confirm("Are you sure you want to delete this customer?")) {
      try {
        const res = await fetch(`/api/customers/${id}`, {
          method: 'DELETE'
        });
        if (res.ok) {
          const data = await res.json();
          applyDBUpdate(data);
        }
      } catch (err) {
        console.error(err);
      }
    }
  };

  const submitFeedback = async (e) => {
    if (e) e.preventDefault();
    if (!feedbackJobId) return;

    try {
      const res = await fetch('/api/feedbacks', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          jobId: feedbackJobId,
          scores: feedbackScores,
          channels: feedbackChannels,
          tools: feedbackTools,
          suggestion: feedbackSuggestion.trim()
        })
      });
      if (res.ok) {
        const data = await res.json();
        applyDBUpdate(data);
        setFeedbackSuccess(true);
        setTimeout(() => {
          setShowFeedbackModal(false);
          setFeedbackScores(initScores());
          setFeedbackChannels([]);
          setFeedbackTools([]);
          setFeedbackSuggestion('');
          setFeedbackSuccess(false);
          // If logout was pending (customer clicked logout), do logout now
          if (window._pendingLogout) {
            window._pendingLogout = false;
            setUserRole(null);
            setActiveCustomerJobId(null);
            setIsMobileMenuOpen(false);
          }
        }, 1800);
      }
    } catch (err) {
      console.error(err);
    }
  };

  const deleteFeedback = async (id) => {
    if (window.confirm("ยืนยันการลบความคิดเห็นประเมินผลนี้?")) {
      try {
        const res = await fetch(`/api/feedbacks/${id}`, {
          method: 'DELETE'
        });
        if (res.ok) {
          const data = await res.json();
          applyDBUpdate(data);
        }
      } catch (err) {
        console.error(err);
      }
    }
  };

  const getMachineName = (id) => {
    const m = machines.find(m => m.id === id);
    return m ? m.name : 'Unknown Machine';
  };

  // Export
  const exportToCSV = (passedJob) => {
    const job = (passedJob && passedJob.id && passedJob.data) ? passedJob : currentJob;
    if (!job || !job.data || job.data.length === 0) {
      alert("ไม่มีข้อมูลที่จะส่งออก!");
      return;
    }
    const headers = [
      'Date',
      'Time',
      'Culture Hour (Hr)',
      'Temp SV (C)', 'Temp PV (C)',
      'pH SV', 'pH PV',
      'DO SV (%)', 'DO PV (%)',
      'Agit SV (RPM)', 'Agit PV (RPM)',
      'Air Flow SV (L/M)', 'Air Flow PV (L/M)',
      'Volume SV (L)', 'Volume PV (L)',
      'Air Out SV (L/M)', 'Air Out PV (L/M)',
      'Heat SV (%)', 'Heat PV (%)',
      'Remarks'
    ];
    const csvRows = [headers.join(',')];
    const mappedRows = mapDataPointsToLiters(job.data, job).sort((a, b) => {
      const aTime = a.timestamp ? new Date(a.timestamp).getTime() : 0;
      const bTime = b.timestamp ? new Date(b.timestamp).getTime() : 0;
      return bTime - aTime;
    });
    const rowsToExport = mappedRows;
    rowsToExport.forEach(row => {
      const dateVal = row.timestamp ? toYYYYMMDD(row.timestamp) : (row.date || '');
      const timeVal = row.timestamp ? toHHMM(row.timestamp) : (row.time || '');
      const hrVal = row.cultureHour !== undefined ? row.cultureHour : getElapsedHours(job, row.timestamp);
      const t_s = row.temp_set !== undefined ? row.temp_set : row.temp;
      const t_r = row.temp_read !== undefined ? row.temp_read : row.temp;
      const p_s = row.ph_set !== undefined ? row.ph_set : row.ph;
      const p_r = row.ph_read !== undefined ? row.ph_read : row.ph;
      const d_s = row.do_set !== undefined ? row.do_set : row.do;
      const d_r = row.do_read !== undefined ? row.do_read : row.do;
      const ag_s = row.agit_set !== undefined ? row.agit_set : row.agit;
      const ag_r = row.agit_read !== undefined ? row.agit_read : row.agit;
      const ai_s = row.air_set !== undefined ? row.air_set : row.air;
      const ai_r = row.air_read !== undefined ? row.air_read : row.air;
      const lv_s = row.level_set !== undefined && row.level_set !== null ? (typeof row.level_set === 'number' ? row.level_set.toFixed(1) : row.level_set) : '-';
      const lv_r = row.level_read !== undefined && row.level_read !== null ? (typeof row.level_read === 'number' ? row.level_read.toFixed(1) : row.level_read) : '-';
      const ao_s = row.air_out_set !== undefined && row.air_out_set !== null ? row.air_out_set : parseFloat(((ai_s || 0) * 0.96).toFixed(2));
      const ao_r = row.air_out_read !== undefined && row.air_out_read !== null ? row.air_out_read : parseFloat(((ai_r || 0) * 0.96).toFixed(2));
      const ht_s = row.heat_set !== undefined && row.heat_set !== null ? row.heat_set : 0.0;
      const ht_r = row.heat_read !== undefined && row.heat_read !== null ? row.heat_read : 0.0;
      const rem = row.remark !== undefined ? row.remark : '';

      const values = [
        `"${dateVal}"`,
        `"${timeVal}"`,
        hrVal,
        t_s, t_r,
        p_s, p_r,
        d_s, d_r,
        ag_s, ag_r,
        ai_s, ai_r,
        lv_s, lv_r,
        ao_s, ao_r,
        ht_s, ht_r,
        `"${rem.replace(/"/g, '""')}"`
      ];
      csvRows.push(values.join(','));
    });
    const csvContent = csvRows.join('\n');
    const BOM = '\uFEFF';
    const blob = new Blob([BOM + csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    const safeName = job.name.replace(/[^a-z0-9]/gi, '_').toLowerCase();
    link.setAttribute('download', `${safeName}_data.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    if (userRole === 'customer') {
      const alreadyReviewed = feedbacks.some(f => f.jobId === job.id);
      if (!alreadyReviewed) {
        setTimeout(() => {
          setFeedbackJobId(job.id);
          setFeedbackScores(initScores());
          setFeedbackChannels([]);
          setFeedbackTools([]);
          setFeedbackSuggestion('');
          setFeedbackSuccess(false);
          setShowFeedbackModal(true);
        }, 800);
      }
    }
  };

  const exportToExcel = (passedJob) => {
    const job = (passedJob && passedJob.id && passedJob.data) ? passedJob : currentJob;
    if (!job || !job.data || job.data.length === 0) {
      alert("ไม่มีข้อมูลที่จะส่งออก!");
      return;
    }

    const mappedRows = mapDataPointsToLiters(job.data, job).sort((a, b) => {
      const aTime = a.timestamp ? new Date(a.timestamp).getTime() : 0;
      const bTime = b.timestamp ? new Date(b.timestamp).getTime() : 0;
      return bTime - aTime;
    });
    const rowsToExport = mappedRows;

    const sheetData = rowsToExport.map(row => {
      const ai_s = row.air_set !== undefined ? row.air_set : row.air;
      const ai_r = row.air_read !== undefined ? row.air_read : row.air;
      return {
        'Date': row.timestamp ? toYYYYMMDD(row.timestamp) : (row.date || ''),
        'Time': row.timestamp ? toHHMM(row.timestamp) : (row.time || ''),
        'Culture Hour (Hr)': row.cultureHour !== undefined ? row.cultureHour : getElapsedHours(job, row.timestamp),
        'Temp SV (°C)': row.temp_set !== undefined ? row.temp_set : row.temp,
        'Temp PV (°C)': row.temp_read !== undefined ? row.temp_read : row.temp,
        'pH SV': row.ph_set !== undefined ? row.ph_set : row.ph,
        'pH PV': row.ph_read !== undefined ? row.ph_read : row.ph,
        'DO SV (%)': row.do_set !== undefined ? row.do_set : row.do,
        'DO PV (%)': row.do_read !== undefined ? row.do_read : row.do,
        'Agit SV (RPM)': row.agit_set !== undefined ? row.agit_set : row.agit,
        'Agit PV (RPM)': row.agit_read !== undefined ? row.agit_read : row.agit,
        'Air Flow SV (L/M)': ai_s,
        'Air Flow PV (L/M)': ai_r,
        'Volume SV (L)': row.level_set !== undefined && row.level_set !== null ? (typeof row.level_set === 'number' ? parseFloat(row.level_set.toFixed(1)) : row.level_set) : '-',
        'Volume PV (L)': row.level_read !== undefined && row.level_read !== null ? (typeof row.level_read === 'number' ? parseFloat(row.level_read.toFixed(1)) : row.level_read) : '-',
        'Air Out SV (L/M)': row.air_out_set !== undefined && row.air_out_set !== null ? row.air_out_set : parseFloat(((ai_s || 0) * 0.96).toFixed(2)),
        'Air Out PV (L/M)': row.air_out_read !== undefined && row.air_out_read !== null ? row.air_out_read : parseFloat(((ai_r || 0) * 0.96).toFixed(2)),
        'Heat SV (%)': row.heat_set !== undefined && row.heat_set !== null ? row.heat_set : 0.0,
        'Heat PV (%)': row.heat_read !== undefined && row.heat_read !== null ? row.heat_read : 0.0,
        'Remarks': row.remark !== undefined ? row.remark : ''
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(sheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Bioprocess Records");

    // Auto-fit column widths
    const max_widths = [];
    sheetData.forEach(row => {
      Object.keys(row).forEach((key, colIndex) => {
        const value = row[key] ? String(row[key]) : "";
        const length = Math.max(value.length, key.length) + 2;
        max_widths[colIndex] = Math.max(max_widths[colIndex] || 0, length);
      });
    });
    worksheet['!cols'] = max_widths.map(w => ({ wch: w }));

    const safeName = job.name.replace(/[^a-z0-9]/gi, '_').toLowerCase();
    XLSX.writeFile(workbook, `${safeName}_data.xlsx`);
    if (userRole === 'customer') {
      const alreadyReviewed = feedbacks.some(f => f.jobId === job.id);
      if (!alreadyReviewed) {
        setTimeout(() => {
          setFeedbackJobId(job.id);
          setFeedbackScores(initScores());
          setFeedbackChannels([]);
          setFeedbackTools([]);
          setFeedbackSuggestion('');
          setFeedbackSuccess(false);
          setShowFeedbackModal(true);
        }, 800);
      }
    }
  };

  const exportToPDF = () => {
    if (!currentJob) return;
    window.print();
  };

  // Auto Simulation removed

  if (!userRole) {
    return (
      <div className="login-overlay">
        {/* Floating orbs for background decoration */}
        <div className="login-orb login-orb-1" />
        <div className="login-orb login-orb-2" />
        <div className="login-orb login-orb-3" />

        {/* Session Expired Banner */}
        {sessionExpiredMessage && (
          <div style={{
            position: 'fixed', top: '1.25rem', left: '50%', transform: 'translateX(-50%)',
            zIndex: 9999, width: 'min(480px, 92vw)',
            background: 'linear-gradient(135deg, rgba(239,68,68,0.18), rgba(185,28,28,0.22))',
            border: '1px solid rgba(239,68,68,0.45)',
            borderRadius: '14px',
            padding: '0.85rem 1.1rem',
            display: 'flex', alignItems: 'flex-start', gap: '0.75rem',
            boxShadow: '0 8px 32px rgba(239,68,68,0.22)',
            animation: 'fadeInDown 0.4s ease'
          }}>
            <span style={{ fontSize: '1.3rem', flexShrink: 0, marginTop: '1px' }}>⏰</span>
            <div style={{ flex: 1 }}>
              <p style={{ margin: 0, fontWeight: 700, color: '#fca5a5', fontSize: '0.88rem', marginBottom: '2px' }}>
                สิทธิ์การเข้าใช้งานหมดอายุแล้ว
              </p>
              <p style={{ margin: 0, color: '#fecaca', fontSize: '0.8rem', lineHeight: 1.5 }}>
                {sessionExpiredMessage}
              </p>
            </div>
            <button
              onClick={() => setSessionExpiredMessage(null)}
              style={{
                background: 'none', border: 'none', cursor: 'pointer',
                color: '#fca5a5', fontSize: '1.1rem', padding: '2px 4px',
                flexShrink: 0, lineHeight: 1
              }}
              title="ปิด"
            >×</button>
          </div>
        )}

        <div className="login-card-wrapper">
          <div className="login-card-inner">

            {/* Logo / Title */}
            <div style={{ textAlign: 'center', marginBottom: '2rem' }}>
              {/* Icon circle */}
              <div style={{
                width: '64px', height: '64px', margin: '0 auto 1rem',
                borderRadius: '16px',
                background: 'linear-gradient(135deg, rgba(0,240,255,0.15), rgba(124,58,237,0.15))',
                border: '1px solid rgba(0,240,255,0.3)',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                fontSize: '2rem',
                boxShadow: '0 0 30px rgba(0,240,255,0.15)'
              }}>🧫</div>

              <h1 style={{
                fontSize: '2.2rem', fontWeight: 800, letterSpacing: '0.06em',
                background: 'linear-gradient(135deg, #00f0ff 0%, #a78bfa 100%)',
                WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent',
                marginBottom: '0.4rem'
              }}>DBMS</h1>
              <p style={{ color: '#7bacc4', fontSize: '0.88rem', letterSpacing: '0.03em' }}>
                เครื่องมือบันทึกและจัดเก็บข้อมูล
              </p>
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
              {/* Admin Login Block */}
              <div style={{
                padding: '1.25rem',
                borderRadius: '14px',
                background: 'rgba(0, 240, 255, 0.04)',
                border: '1px solid rgba(0, 240, 255, 0.12)'
              }}>
                <h3 style={{
                  marginBottom: '1rem', fontSize: '0.95rem', fontWeight: 700,
                  color: '#00f0ff', display: 'flex', alignItems: 'center', gap: '0.5rem',
                  letterSpacing: '0.02em'
                }}>🔐 สำหรับแอดมิน (Admin)</h3>
                <form onSubmit={async (e) => {
                  e.preventDefault();
                  const password = e.target.adminPassword.value;
                  try {
                    const res = await fetch('/api/settings/verify-password', {
                      method: 'POST',
                      headers: { 'Content-Type': 'application/json' },
                      body: JSON.stringify({ password })
                    });
                    if (res.ok) {
                      setUserRole('admin');
                    } else {
                      alert('รหัสผ่านไม่ถูกต้อง กรุณาลองใหม่อีกครั้ง');
                    }
                  } catch (err) {
                    console.error('Error verifying password:', err);
                    alert('ไม่สามารถเชื่อมต่อเซิร์ฟเวอร์ได้');
                  }
                }} style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>
                  <input
                    type="password"
                    name="adminPassword"
                    placeholder="รหัสผ่านแอดมิน"
                    style={{
                      width: '100%', padding: '11px 14px', borderRadius: '10px',
                      border: '1px solid rgba(0,240,255,0.2)',
                      background: 'rgba(0, 20, 40, 0.6)',
                      color: 'white', fontSize: '0.9rem',
                      outline: 'none', transition: 'border-color 0.2s, box-shadow 0.2s',
                      boxSizing: 'border-box'
                    }}
                    onFocus={e => {
                      e.target.style.borderColor = 'rgba(0,240,255,0.5)';
                      e.target.style.boxShadow = '0 0 0 3px rgba(0,240,255,0.1)';
                    }}
                    onBlur={e => {
                      e.target.style.borderColor = 'rgba(0,240,255,0.2)';
                      e.target.style.boxShadow = 'none';
                    }}
                  />
                  <button type="submit" style={{
                    width: '100%', padding: '11px', borderRadius: '10px', border: 'none',
                    background: 'linear-gradient(135deg, #0ea5e9, #0284c7)',
                    color: 'white', fontWeight: 700, fontSize: '0.9rem', cursor: 'pointer',
                    transition: 'transform 0.15s, box-shadow 0.15s',
                    boxShadow: '0 4px 15px rgba(14,165,233,0.35)'
                  }}
                    onMouseEnter={e => { e.currentTarget.style.transform = 'translateY(-1px)'; e.currentTarget.style.boxShadow = '0 6px 20px rgba(14,165,233,0.45)'; }}
                    onMouseLeave={e => { e.currentTarget.style.transform = 'translateY(0)'; e.currentTarget.style.boxShadow = '0 4px 15px rgba(14,165,233,0.35)'; }}
                  >เข้าสู่ระบบแอดมิน</button>
                </form>
              </div>

              {/* Divider */}
              <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
                <div style={{ flex: 1, height: '1px', background: 'rgba(255,255,255,0.08)' }} />
                <span style={{ color: '#4a6a7a', fontSize: '0.78rem', letterSpacing: '0.05em' }}>หรือ</span>
                <div style={{ flex: 1, height: '1px', background: 'rgba(255,255,255,0.08)' }} />
              </div>

              {/* Customer Login Block */}
              <div style={{
                padding: '1.25rem',
                borderRadius: '14px',
                background: 'rgba(124, 58, 237, 0.04)',
                border: '1px solid rgba(124, 58, 237, 0.15)'
              }}>
                <h3 style={{
                  marginBottom: '1rem', fontSize: '0.95rem', fontWeight: 700,
                  color: '#a78bfa', display: 'flex', alignItems: 'center', gap: '0.5rem',
                  letterSpacing: '0.02em'
                }}>👥 สำหรับลูกค้า (Customer)</h3>
                <form onSubmit={async (e) => {
                  e.preventDefault();
                  const jobCode = e.target.jobCode.value.trim();
                  try {
                    const res = await fetch('/api/db');
                    if (!res.ok) throw new Error('Failed to fetch DB');
                    const data = await res.json();
                    const jobExists = data.jobs.find(j => j.id === jobCode);
                    if (!jobExists) {
                      alert('ไม่พบรหัสงานนี้ในระบบ กรุณาตรวจสอบรหัสอีกครั้ง');
                      return;
                    }

                    const isExpired = jobExists.expiresAt && new Date() > new Date(jobExists.expiresAt);
                    const isFinished = jobExists.status === 'finished';
                    if ((isExpired || isFinished) && !jobExists.allowHistoryView) {
                      alert('สิทธิ์การเข้าใช้งานหมดอายุหรือเสร็จสิ้นแล้ว (หากต้องการดูย้อนหลัง กรุณาติดต่อเจ้าหน้าที่เพื่อขอเปิดสิทธิ์)');
                      return;
                    }
                    try {
                      const ack = localStorage.getItem('bioprocess-customer-notice-ack') === 'true';
                      if (ack) {
                        applyDBUpdate(data);
                        setActiveCustomerJobId(jobCode);
                        setUserRole('customer');
                        setCurrentAppView('monitoring');
                        return;
                      }
                    } catch (err) { }
                    setPendingJobCode(jobCode);
                    setShowCustomerNotice(true);
                  } catch (err) {
                    console.error(err);
                    alert('เกิดข้อผิดพลาดขณะตรวจสอบรหัสงาน กรุณาลองใหม่');
                  }
                }} style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>
                  <input
                    type="text"
                    name="jobCode"
                    placeholder="ป้อนรหัสงานของคุณ (เช่น job-...)"
                    style={{
                      width: '100%', padding: '11px 14px', borderRadius: '10px',
                      border: '1px solid rgba(124,58,237,0.25)',
                      background: 'rgba(10, 5, 30, 0.6)',
                      color: 'white', fontSize: '0.9rem',
                      outline: 'none', transition: 'border-color 0.2s, box-shadow 0.2s',
                      boxSizing: 'border-box'
                    }}
                    onFocus={e => {
                      e.target.style.borderColor = 'rgba(167,139,250,0.6)';
                      e.target.style.boxShadow = '0 0 0 3px rgba(124,58,237,0.12)';
                    }}
                    onBlur={e => {
                      e.target.style.borderColor = 'rgba(124,58,237,0.25)';
                      e.target.style.boxShadow = 'none';
                    }}
                  />
                  <button type="submit" style={{
                    width: '100%', padding: '11px', borderRadius: '10px', border: 'none',
                    background: 'linear-gradient(135deg, #7c3aed, #6d28d9)',
                    color: 'white', fontWeight: 700, fontSize: '0.9rem', cursor: 'pointer',
                    transition: 'transform 0.15s, box-shadow 0.15s',
                    boxShadow: '0 4px 15px rgba(124,58,237,0.35)'
                  }}
                    onMouseEnter={e => { e.currentTarget.style.transform = 'translateY(-1px)'; e.currentTarget.style.boxShadow = '0 6px 20px rgba(124,58,237,0.45)'; }}
                    onMouseLeave={e => { e.currentTarget.style.transform = 'translateY(0)'; e.currentTarget.style.boxShadow = '0 4px 15px rgba(124,58,237,0.35)'; }}
                  >เข้าดูข้อมูลงาน</button>
                </form>
              </div>
            </div>

            {/* Footer */}
            <p style={{ textAlign: 'center', marginTop: '1.5rem', fontSize: '0.75rem', color: '#3a5a6a', letterSpacing: '0.04em' }}>
              Bioprocess Data Logging System © {new Date().getFullYear()}
            </p>
          </div>
        </div>

        {showCustomerNotice && (
          <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }}>
            <div style={{ width: 520, maxWidth: '95%', background: 'var(--panel-bg)', border: '1px solid var(--border-color)', borderRadius: 12, padding: '1.25rem' }}>
              <h3 style={{ marginBottom: '0.5rem' }}>⚠️ แจ้งเตือนการเข้าใช้งาน</h3>
              <div style={{ color: 'var(--text-secondary)', marginBottom: '1rem' }}>
                รหัสผ่านนี้ใช้เข้าดูได้เฉพาะข้อมูลในส่วนงานของลูกค้าที่ได้รับอนุญาตเท่านั้น.
                ข้อมูลทั้งหมดจะถูกจัดเก็บเป็นความลับ และระบบจะสำรองข้อมูลไว้เป็นเวลา 7 วันหลังเสร็จงาน.
              </div>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '0.5rem', marginTop: '8px' }}>
                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', color: 'var(--text-secondary)' }}>
                  <input type="checkbox" checked={dontShowAgain} onChange={(e) => setDontShowAgain(e.target.checked)} />
                  อย่าขึ้นเตือนอีก
                </label>
                <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '0.5rem' }}>
                  <button onClick={() => { setShowCustomerNotice(false); setPendingJobCode(null); }} className="btn btn-secondary" style={{ margin: 0 }}>ยกเลิก</button>
                  <button onClick={async () => {
                    if (dontShowAgain) {
                      try { localStorage.setItem('bioprocess-customer-notice-ack', 'true'); } catch (e) { }
                    }
                    if (pendingJobCode) {
                      try {
                        const res = await fetch('/api/db');
                        if (res.ok) {
                          const data = await res.json();
                          applyDBUpdate(data);
                        }
                      } catch (e) {
                        console.error('Failed to refresh DB before customer access', e);
                      }
                      setActiveCustomerJobId(pendingJobCode);
                      setUserRole('customer');
                      setCurrentAppView('monitoring');
                    }
                    setShowCustomerNotice(false);
                    setPendingJobCode(null);
                  }} className="btn btn-blue" style={{ margin: 0 }}>ยอมรับและเข้าสู่ระบบ</button>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>

    );
  }


  return (
    <div className="app-layout">
      {/* Customer Warning Toast Notification */}
      {showCustomerToast && (
        <div className={`login-toast-container ${isToastHiding ? 'hiding' : ''}`}>
          <div className="login-toast">
            <div className="login-toast-header">
              <span>⚠️ แจ้งเตือนการใช้งานระบบ</span>
              <button className="login-toast-close" onClick={dismissCustomerToast}>&times;</button>
            </div>
            <div className="login-toast-body">
              <div className="login-toast-item">
                <span className="login-toast-bullet">•</span>
                <span>ระบบนี้จัดทำขึ้นเพื่อบันทึกและจัดเก็บข้อมูลสำหรับการปฏิบัติงานและการจัดทำรายงานเท่านั้น</span>
              </div>
              <div className="login-toast-item">
                <span className="login-toast-bullet">•</span>
                <span>ID และรหัสผ่านเป็นข้อมูลสำคัญ ห้ามเปิดเผยหรือแชร์ให้บุคคลที่ไม่เกี่ยวข้อง</span>
              </div>
              <div className="login-toast-item">
                <span className="login-toast-bullet">•</span>
                <span>สิทธิ์การเข้าใช้งานระบบถูกกำหนดเฉพาะผู้ได้รับอนุญาตเท่านั้น</span>
              </div>
              <div className="login-toast-item">
                <span className="login-toast-bullet">•</span>
                <span>ข้อมูลจะถูกสำรองไว้ 7 วัน นับจากวันที่บันทึกข้อมูลล่าสุด</span>
              </div>
            </div>
          </div>
        </div>
      )}
      {showCustomerNotice && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }}>
          <div style={{ width: 520, maxWidth: '95%', background: 'var(--panel-bg)', border: '1px solid var(--border-color)', borderRadius: 12, padding: '1.25rem' }}>
            <h3 style={{ marginBottom: '0.5rem' }}>⚠️ แจ้งเตือนการเข้าใช้งาน</h3>
            <div style={{ color: 'var(--text-secondary)', marginBottom: '1rem' }}>
              รหัสผ่านนี้ใช้เข้าดูได้เฉพาะข้อมูลในส่วนงานของลูกค้าที่ได้รับอนุญาตเท่านั้น.
              ข้อมูลทั้งหมดจะถูกจัดเก็บเป็นความลับ และระบบจะสำรองข้อมูลไว้เป็นเวลา 7 วันหลังเสร็จงาน.
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '0.5rem', marginTop: '8px' }}>
              <label style={{ display: 'flex', alignItems: 'center', gap: '8px', color: 'var(--text-secondary)' }}>
                <input type="checkbox" checked={dontShowAgain} onChange={(e) => setDontShowAgain(e.target.checked)} />
                อย่าขึ้นเตือนอีก
              </label>
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '0.5rem' }}>
                <button onClick={() => { setShowCustomerNotice(false); setPendingJobCode(null); }} className="btn btn-secondary" style={{ margin: 0 }}>ยกเลิก</button>
                <button onClick={async () => {
                  if (dontShowAgain) {
                    try { localStorage.setItem('bioprocess-customer-notice-ack', 'true'); } catch (e) { }
                  }
                  if (pendingJobCode) {
                    // Refresh DB before granting access to ensure currentJobId/machine are set
                    try {
                      const res = await fetch('/api/db');
                      if (res.ok) {
                        const data = await res.json();
                        applyDBUpdate(data);
                      }
                    } catch (e) {
                      console.error('Failed to refresh DB before customer access', e);
                    }
                    setActiveCustomerJobId(pendingJobCode);
                    setUserRole('customer');
                    setCurrentAppView('monitoring');
                  }
                  setShowCustomerNotice(false);
                  setPendingJobCode(null);
                }} className="btn btn-blue" style={{ margin: 0 }}>ยอมรับและเข้าสู่ระบบ</button>
              </div>
            </div>
          </div>
        </div>
      )}
      {/* Mobile Header Bar */}
      <div className="mobile-topbar">
        <button
          onClick={() => setIsMobileMenuOpen(true)}
          style={{ background: 'transparent', border: 'none', color: 'white', cursor: 'pointer', display: 'flex', alignItems: 'center' }}
        >
          <Menu size={24} />
        </button>
        <span className="mobile-topbar-title">
          <ActivityIcon size={20} style={{ color: '#00f0ff', filter: 'drop-shadow(0 0 6px rgba(0, 240, 255, 0.5))' }} />
          BIOPROCESS
        </span>
        <span style={{ fontSize: '0.8rem', color: '#84b2bc', fontWeight: 600 }}>
          {currentMachine?.name || 'Bioprocess'}
        </span>
      </div>

      {/* Backdrop overlay for mobile menu */}
      {isMobileMenuOpen && (
        <div
          className="mobile-overlay"
          onClick={() => setIsMobileMenuOpen(false)}
        />
      )}

      {/* SIDEBAR */}
      <aside className={`sidebar ${isMobileMenuOpen ? 'open' : ''}`}>
        {/* Header */}
        <div className="sidebar-header">
          <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem', width: '100%', justifyContent: 'space-between' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
              <ActivityIcon size={28} style={{ color: '#00f0ff', filter: 'drop-shadow(0 0 6px rgba(0, 240, 255, 0.5))' }} />
              <div>
                <div className="sidebar-logo-text">DBMS</div>
                <div className="sidebar-logo-subtext">SYSTEM ULTRA</div>
              </div>
            </div>
            {/* Close button on mobile sidebar header */}
            <button
              onClick={() => setIsMobileMenuOpen(false)}
              style={{ background: 'transparent', border: 'none', color: '#84b2bc', cursor: 'pointer', display: 'none' }}
              className="mobile-close-btn"
            >
              <X size={20} />
            </button>
          </div>
        </div>

        {/* Menu */}
        <div className="sidebar-menu">
          {userRole === 'admin' ? (
            <>
              {/* Active Users Status Component */}
              <div className="sidebar-active-users" style={{ padding: '0 1.25rem 1rem 1.25rem', borderBottom: '1px solid rgba(255, 255, 255, 0.05)', marginBottom: '0.5rem' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '0.75rem', color: '#84b2bc', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: '0.5rem', fontWeight: 700 }}>
                  <span style={{ width: '8px', height: '8px', borderRadius: '50%', backgroundColor: '#10b981', display: 'inline-block', boxShadow: '0 0 8px #10b981' }} className="status-dot"></span>
                  ผู้ใช้งานออนไลน์ ({activeUsers.length})
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                  {activeUsers.map((u, i) => {
                    const isSelf = u.clientId === clientId;
                    const mName = machines.find(m => m.id === u.machineId)?.name || '';
                    const jName = jobs.find(j => j.id === u.jobId)?.name || '';
                    return (
                      <div key={u.clientId || i} style={{ fontSize: '0.8rem', padding: '6px 8px', borderRadius: '6px', background: isSelf ? 'rgba(0, 240, 255, 0.03)' : 'rgba(255, 255, 255, 0.02)', border: isSelf ? '1px solid rgba(0, 240, 255, 0.15)' : '1px solid rgba(255, 255, 255, 0.05)' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', fontWeight: 600, color: isSelf ? '#00f0ff' : 'var(--text-primary)' }}>
                          <span>
                            {u.role === 'admin' ? '🔐 แอดมิน' : '👥 ลูกค้า'}
                            {isSelf && ' (คุณ)'}
                          </span>
                          <span style={{ fontSize: '0.7rem', color: '#10b981', fontWeight: 600 }}>Active</span>
                        </div>
                        {(mName || jName) && (
                          <div style={{ fontSize: '0.7rem', color: '#84b2bc', marginTop: '2px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                            กำลังดู: {mName} {jName && `› ${jName}`}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Category: Monitoring */}
              <div className="sidebar-menu-header">Monitoring</div>

              {/* Menu Dashboard */}
              <div
                className={`sidebar-menu-item ${currentAppView === 'monitoring' ? 'active' : ''}`}
                onClick={() => {
                  setCurrentAppView('monitoring');
                  setIsMobileMenuOpen(false);
                }}
              >
                <span className="sidebar-menu-link">
                  <LayoutDashboard size={18} />
                  Dashboard / Monitor
                </span>
              </div>

              {/* Menu Sessions */}
              <div
                className={`sidebar-menu-item ${currentAppView === 'sessions' ? 'active' : ''}`}
                onClick={() => {
                  setCurrentAppView('sessions');
                  setIsMobileMenuOpen(false);
                }}
              >
                <span className="sidebar-menu-link">
                  <FolderOpen size={18} />
                  รอบรันทั้งหมด (Sessions)
                </span>
                <span className="sidebar-badge">{jobs.length}</span>
              </div>

              {/* Menu Instruments */}
              <div
                className={`sidebar-menu-item ${currentAppView === 'instruments' ? 'active' : ''}`}
                onClick={() => {
                  setCurrentAppView('instruments');
                  setIsMobileMenuOpen(false);
                }}
              >
                <span className="sidebar-menu-link">
                  <Cpu size={18} />
                  เครื่องมือ (Instruments)
                </span>
                <span className="sidebar-badge">{machines.length}</span>
              </div>

              {/* Category: Management */}
              <div className="sidebar-menu-header">DATA MANAGEMENT</div>

              {/* Menu Manual Data Entry */}
              <div
                className={`sidebar-menu-item ${currentAppView === 'manual_entry' ? 'active' : ''}`}
                onClick={() => {
                  setCurrentAppView('manual_entry');
                  setIsMobileMenuOpen(false);
                }}
              >
                <span className="sidebar-menu-link">
                  <PlusCircle size={18} />
                  Manual Data Entry
                </span>
              </div>

              {/* Menu Customer Database */}
              <div
                className={`sidebar-menu-item ${currentAppView === 'customers' ? 'active' : ''}`}
                onClick={() => {
                  setCurrentAppView('customers');
                  setIsMobileMenuOpen(false);
                }}
              >
                <span className="sidebar-menu-link">
                  <Users size={18} />
                  Customer Database
                </span>
                <span className="sidebar-badge">{customers.length}</span>
              </div>

              {/* Menu Customer Feedbacks */}
              <div
                className={`sidebar-menu-item ${currentAppView === 'feedbacks' ? 'active' : ''}`}
                onClick={() => {
                  setCurrentAppView('feedbacks');
                  setIsMobileMenuOpen(false);
                }}
              >
                <span className="sidebar-menu-link">
                  <Star size={18} />
                  Customer Feedbacks
                </span>
                <span className="sidebar-badge">{feedbacks.length}</span>
              </div>

              {/* Menu System Settings */}
              <div
                className={`sidebar-menu-item ${currentAppView === 'settings' ? 'active' : ''}`}
                onClick={() => {
                  setCurrentAppView('settings');
                  setIsMobileMenuOpen(false);
                }}
              >
                <span className="sidebar-menu-link">
                  <Settings size={18} />
                  System Settings
                </span>
              </div>
            </>
          ) : (
            /* Customer Portal Navigation */
            <div style={{ padding: '0 1.25rem', display: 'flex', flexDirection: 'column', gap: '1rem' }}>
              <div style={{ borderBottom: '1px solid rgba(255,255,255,0.05)', paddingBottom: '1rem' }}>
                <div style={{ fontSize: '0.75rem', color: '#84b2bc', textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: '0.25rem' }}>👤 Customer Portal</div>
                <div style={{ fontSize: '0.85rem', fontWeight: 600, color: '#00f0ff', wordBreak: 'break-all' }}>ID: {activeCustomerJobId}</div>
              </div>

              <div>
                <div style={{ fontSize: '0.75rem', color: '#84b2bc', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Instrument</div>
                <div style={{ fontSize: '0.95rem', fontWeight: 600, color: '#ffffff', marginTop: '0.25rem' }}>{currentMachine?.name}</div>
              </div>

              <div>
                <div style={{ fontSize: '0.75rem', color: '#84b2bc', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Session</div>
                <div style={{ fontSize: '0.95rem', fontWeight: 600, color: '#ffffff', marginTop: '0.25rem' }}>{currentJob?.name}</div>
              </div>

              {showCustomerBanner && (
                <div style={{ padding: '12px', borderRadius: '8px', background: 'rgba(255,255,255,0.02)', borderLeft: '3px solid #f59e0b', fontSize: '0.8rem', lineHeight: 1.4, color: '#cbdce0' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                    <div>
                      <div style={{ fontWeight: 700, color: '#f59e0b', marginBottom: '4px' }}>⚠️ บันทึกข้อมูลและสิทธิ์</div>
                      <div>• สิทธิ์ในการอ่านข้อมูลเท่านั้น</div>
                      <div>• ข้อมูลถูกเก็บเป็นความลับ</div>
                      <div>• สำรองข้อมูล 7 วันหลังจบงาน</div>
                    </div>
                    <button onClick={() => setShowCustomerBanner(false)} style={{ background: 'transparent', border: 'none', color: '#84b2bc', cursor: 'pointer' }}>✕</button>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="sidebar-footer">
          {/* Feedback Button - customer only */}
          {userRole === 'customer' && (
            <button
              className="sidebar-logout-btn"
              style={{ marginBottom: '0.5rem', background: 'linear-gradient(135deg, rgba(234,179,8,0.18), rgba(234,179,8,0.06))', borderColor: 'rgba(234,179,8,0.4)', color: '#fde68a' }}
              onClick={() => {
                setFeedbackJobId(activeCustomerJobId || null);
                setFeedbackScores(initScores());
                setFeedbackChannels([]);
                setFeedbackTools([]);
                setFeedbackSuggestion('');
                setFeedbackSuccess(false);
                setShowFeedbackModal(true);
              }}
            >
              <Star size={16} />
              ประเมินความพึงพอใจ
            </button>
          )}
          <button
            className="sidebar-logout-btn"
            onClick={() => {
              if (userRole === 'customer') {
                const alreadyReviewed = feedbacks.some(f => f.jobId === activeCustomerJobId);
                if (alreadyReviewed) {
                  setUserRole(null);
                  setActiveCustomerJobId(null);
                  setIsMobileMenuOpen(false);
                } else {
                  // Show feedback before logout for customer
                  setFeedbackJobId(activeCustomerJobId || null);
                  setFeedbackScores(initScores());
                  setFeedbackChannels([]);
                  setFeedbackTools([]);
                  setFeedbackSuggestion('');
                  setFeedbackSuccess(false);
                  setShowFeedbackModal(true);
                  // Set a flag to logout after modal closes
                  window._pendingLogout = true;
                }
              } else {
                if (window.confirm('ยืนยันการออกจากระบบ?')) {
                  setUserRole(null);
                  setActiveCustomerJobId(null);
                  setIsMobileMenuOpen(false);
                }
              }
            }}
          >
            <LogOut size={16} />
            Logout
          </button>
        </div>
      </aside>

      {/* MAIN CONTENT */}
      <main className="main-content">
        {/* Print-only Header */}
        <div className="print-only-header">
          <h1>Bioprocess Data Report</h1>
          <p><strong>Session Name:</strong> {currentJob?.name || '-'}</p>
          <p><strong>Instrument / Machine:</strong> {currentMachine?.name || '-'}</p>
          <p><strong>Date Generated:</strong> {formatDateTime(new Date())}</p>
        </div>
        {currentAppView === 'customers' ? (
          /* CUSTOMER DATA VIEW */
          <div className="customers-view">
            <header className="dashboard-header">
              <h2>Customer Database</h2>
            </header>

            <div className="glass-panel" style={{ padding: '1.5rem', marginBottom: '1rem' }}>
              <h2 className="form-title">Add New Customer</h2>
              <form onSubmit={handleAddCustomer} className="customer-add-form">
                <div className="form-group" style={{ flex: 2 }}>
                  <label>Company / Customer Name</label>
                  <input
                    type="text"
                    name="companyName"
                    placeholder="บริษัท/ชื่อลูกค้า"
                    value={customerFormData.companyName}
                    onChange={handleCustomerInputChange}
                  />
                </div>

                <div className="form-group" style={{ flex: 1 }}>
                  <label>Email (อีเมล)</label>
                  <input
                    type="email"
                    name="email"
                    placeholder="customer@example.com"
                    value={customerFormData.email}
                    onChange={handleCustomerInputChange}
                  />
                </div>

                <div className="form-group" style={{ flex: 1 }}>
                  <label>Instrument / Machine Used</label>
                  <select
                    name="machineId"
                    value={customerFormData.machineId}
                    onChange={handleCustomerInputChange}
                    className="machine-dropdown"
                    style={{ backgroundImage: 'none' }}
                  >
                    {machines.map(m => (
                      <option key={m.id} value={m.id}>{m.name}</option>
                    ))}
                  </select>
                </div>

                <div className="form-group" style={{ flex: 'none' }}>
                  <label style={{ visibility: 'hidden' }}>Submit</label>
                  <button type="submit" className="btn btn-primary" style={{ margin: 0, height: '42px' }}>
                    <Users size={18} style={{ marginRight: '8px' }} /> Add Customer
                  </button>
                </div>
              </form>
            </div>

            <div className="glass-panel" style={{ padding: '2rem' }}>
              <div className="table-header-controls">
                <h2 className="chart-title">Customer List</h2>
                <span style={{ color: 'var(--text-secondary)' }}>Total: {customers.length}</span>
              </div>
              <div className="data-table-container">
                <table className="data-table">
                  <thead>
                    <tr>
                      <th>Company / Customer Name</th>
                      <th>Email</th>
                      <th>Machine Assigned</th>
                      <th>Session & Share Link</th>
                      <th>Date Added</th>
                      <th style={{ width: '80px', textAlign: 'center' }}>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {customers.length === 0 ? (
                      <tr>
                        <td colSpan="6" style={{ textAlign: 'center', padding: '2rem', color: 'var(--text-secondary)' }}>
                          No customers recorded yet.
                        </td>
                      </tr>
                    ) : (
                      customers.map((c) => (
                        <tr key={c.id}>
                          <td style={{ fontWeight: 600 }}>{c.companyName}</td>
                          <td style={{ color: 'var(--text-secondary)' }}>{c.email || '-'}</td>
                          <td>
                            <span style={{ background: 'rgba(59, 130, 246, 0.2)', color: 'var(--accent-blue)', padding: '4px 8px', borderRadius: '4px', fontSize: '0.85rem' }}>
                              {getMachineName(c.machineId)}
                            </span>
                          </td>
                          <td>
                            {(() => {
                              const custJobs = jobs.filter(j => j.machineId === c.machineId);
                              if (custJobs.length === 0) {
                                return <span style={{ color: 'var(--text-secondary)', fontSize: '0.85rem' }}>ไม่มีเซสชันในเครื่องนี้</span>;
                              }
                              const selectedJobId = selectedCustomerJobs[c.id] || custJobs[0].id;
                              return (
                                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                  <select
                                    value={selectedJobId}
                                    onChange={(e) => setSelectedCustomerJobs(prev => ({ ...prev, [c.id]: e.target.value }))}
                                    className="machine-dropdown"
                                    style={{ padding: '4px 8px', height: '32px', fontSize: '0.85rem', width: 'auto', display: 'inline-block', minWidth: '120px', backgroundImage: 'none' }}
                                  >
                                    {custJobs.map(j => (
                                      <option key={j.id} value={j.id}>{j.name}</option>
                                    ))}
                                  </select>
                                  <button
                                    className="btn-outline-cyan"
                                    onClick={() => {
                                      const targetJob = custJobs.find(j => j.id === selectedJobId) || custJobs[0];
                                      setShareModalJobId(targetJob.id);
                                      setShowCustomerShareModal(true);
                                    }}
                                  >
                                    ลิงก์แชร์
                                  </button>
                                </div>
                              );
                            })()}
                          </td>
                          <td style={{ color: 'var(--text-secondary)' }}>{formatCreatedAt(c.createdAt)}</td>
                          <td style={{ textAlign: 'center' }}>
                            <button
                              className="delete-row-btn"
                              onClick={() => deleteCustomer(c.id)}
                              title="Delete customer"
                            >
                              <Trash2 size={16} />
                            </button>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        ) : currentAppView === 'instruments' ? (
          /* INSTRUMENTS VIEW */
          <div className="instruments-view">
            <header className="dashboard-header" style={{ marginBottom: '1.5rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '1rem' }}>
              <div>
                <h2>จัดการเครื่องมือ (Instruments Management)</h2>
                <p style={{ color: 'var(--text-secondary)', fontSize: '0.9rem', marginTop: '4px' }}>
                  แสดงรายการ ลบ แก้ไขชื่อ และเพิ่มเครื่องมือ/อุปกรณ์ bioreactor ใหม่ทั้งหมด
                </p>
              </div>
              <button
                className="btn btn-blue"
                style={{ margin: 0 }}
                onClick={() => {
                  setNewMachineName('');
                  setShowAddMachineModal(true);
                }}
              >
                <PlusCircle size={18} />
                + เพิ่มเครื่องมือใหม่
              </button>
            </header>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: '1.5rem' }}>
              {machines.map(m => {
                const machineJobs = jobs.filter(j => j.machineId === m.id);
                const isActive = m.id === currentMachineId;
                const imgSrc = m.imageData || '/bioreactor.png';
                const fileInputId = `machine-img-input-${m.id}`;

                return (
                  <div key={m.id} className="glass-panel" style={{ padding: '1.5rem', border: isActive ? '1px solid rgba(0, 240, 255, 0.4)' : '1px solid var(--border-color)', boxShadow: isActive ? '0 0 15px rgba(0, 240, 255, 0.15)' : 'none', position: 'relative' }}>

                    {/* Hidden File Input */}
                    <input
                      id={fileInputId}
                      type="file"
                      accept="image/*"
                      style={{ display: 'none' }}
                      onChange={(e) => {
                        const file = e.target.files?.[0];
                        if (!file) return;
                        compressImage(file).then((base64) => {
                          if (!base64) return;
                          fetch(`/api/machines/${m.id}/image`, {
                            method: 'PUT',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ imageData: base64 })
                          })
                            .then(res => res.ok && res.json())
                            .then(data => data && applyDBUpdate(data))
                            .catch(err => {
                              console.error('Error uploading machine image:', err);
                              alert('เกิดข้อผิดพลาดในการอัปโหลดรูปภาพ');
                            });
                        });
                        // reset input so same file can be re-selected
                        e.target.value = '';
                      }}
                    />

                    {/* Image area — click to upload */}
                    <div
                      style={{ marginBottom: '0.75rem', position: 'relative', display: 'inline-block', cursor: 'pointer' }}
                      title="คลิกเพื่อเปลี่ยนรูปเครื่องมือ"
                      onClick={() => document.getElementById(fileInputId)?.click()}
                    >
                      <img
                        src={imgSrc}
                        alt={m.name}
                        style={{
                          width: '80px',
                          height: '80px',
                          objectFit: m.imageData ? 'cover' : 'contain',
                          borderRadius: m.imageData ? '10px' : '0',
                          filter: isActive ? 'drop-shadow(0 0 8px rgba(0,240,255,0.6))' : 'drop-shadow(0 2px 6px rgba(0,0,0,0.5))',
                          transition: 'filter 0.3s, transform 0.2s',
                          border: m.imageData ? '2px solid rgba(0,240,255,0.25)' : 'none',
                          display: 'block'
                        }}
                        onMouseEnter={e => e.currentTarget.style.transform = 'scale(1.05)'}
                        onMouseLeave={e => e.currentTarget.style.transform = 'scale(1)'}
                      />
                      {/* Camera overlay icon */}
                      <div style={{
                        position: 'absolute', bottom: 0, right: 0,
                        background: 'rgba(0,240,255,0.18)', backdropFilter: 'blur(4px)',
                        border: '1px solid rgba(0,240,255,0.35)',
                        borderRadius: '6px', padding: '2px 5px',
                        fontSize: '0.65rem', color: '#00f0ff', lineHeight: 1.4,
                        pointerEvents: 'none'
                      }}>📷</div>
                    </div>

                    <h3 style={{ fontSize: '1.2rem', fontWeight: 700, marginBottom: '0.5rem', color: isActive ? '#00f0ff' : 'var(--text-primary)' }}>{m.name}</h3>
                    <p style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: '1.5rem' }}>
                      มีรันบันทึกข้อมูลทั้งหมด: <strong>{machineJobs.length} รอบ</strong>
                    </p>

                    <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap', borderTop: '1px solid var(--border-color)', paddingTop: '1rem' }}>
                      <button
                        className="replay-close-btn"
                        style={{
                          background: 'rgba(0, 240, 255, 0.1)',
                          borderColor: 'rgba(0, 240, 255, 0.2)',
                          color: '#00f0ff',
                          padding: '6px 12px',
                          fontSize: '0.8rem',
                          margin: 0
                        }}
                        onClick={() => {
                          setCurrentMachineId(m.id);
                          if (machineJobs.length > 0) {
                            setCurrentJobId(machineJobs[0].id);
                          } else {
                            setCurrentJobId(null);
                          }
                          setIsViewingHistory(false);
                          setCurrentAppView('monitoring');
                          setActiveTab('dashboard');
                        }}
                      >
                        เปิดบอร์ดข้อมูล
                      </button>
                      <button
                        className="export-btn"
                        style={{ padding: '6px 12px', fontSize: '0.8rem', margin: 0 }}
                        onClick={() => {
                          setEditingMachineId(m.id);
                          setEditingMachineName(m.name);
                          setShowEditMachineModal(true);
                        }}
                      >
                        แก้ไขชื่อ
                      </button>
                      {/* Upload image button */}
                      <button
                        style={{
                          padding: '6px 10px', fontSize: '0.8rem', margin: 0,
                          background: 'rgba(139,92,246,0.12)',
                          border: '1px solid rgba(139,92,246,0.35)',
                          color: '#a78bfa', borderRadius: '8px', cursor: 'pointer',
                          display: 'flex', alignItems: 'center', gap: '4px',
                          transition: 'background 0.18s'
                        }}
                        onMouseEnter={e => e.currentTarget.style.background = 'rgba(139,92,246,0.25)'}
                        onMouseLeave={e => e.currentTarget.style.background = 'rgba(139,92,246,0.12)'}
                        title="อัปโหลดรูปเครื่องมือจริง"
                        onClick={() => document.getElementById(fileInputId)?.click()}
                      >
                        📷 อัปโหลดรูป
                      </button>
                      {/* Remove custom image button (only if custom image is set) */}
                      {m.imageData && (
                        <button
                          style={{
                            padding: '6px 10px', fontSize: '0.8rem', margin: 0,
                            background: 'rgba(239,68,68,0.08)',
                            border: '1px solid rgba(239,68,68,0.3)',
                            color: '#f87171', borderRadius: '8px', cursor: 'pointer',
                            transition: 'background 0.18s'
                          }}
                          onMouseEnter={e => e.currentTarget.style.background = 'rgba(239,68,68,0.2)'}
                          onMouseLeave={e => e.currentTarget.style.background = 'rgba(239,68,68,0.08)'}
                          title="ลบรูปที่อัปโหลด กลับไปใช้รูป default"
                          onClick={() => {
                            if (window.confirm('ลบรูปที่อัปโหลด แล้วกลับใช้รูป default?')) {
                              fetch(`/api/machines/${m.id}/image`, {
                                method: 'PUT',
                                headers: { 'Content-Type': 'application/json' },
                                body: JSON.stringify({ imageData: null })
                              }).then(res => res.ok && res.json()).then(data => data && applyDBUpdate(data));
                            }
                          }}
                        >
                          🗑 ลบรูป
                        </button>
                      )}
                      {machines.length > 1 && (
                        <button
                          className="delete-row-btn"
                          style={{ margin: 0, padding: '6px', width: '32px', height: '32px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}
                          onClick={() => {
                            if (window.confirm(`คุณแน่ใจว่าต้องการลบเครื่องมือ "${m.name}"? รอบข้อมูลและประวัติรันทั้งหมดจะสูญหายอย่างถาวร!`)) {
                              fetch(`/api/machines/${m.id}`, {
                                method: 'DELETE'
                              }).then(res => res.ok && res.json()).then(data => {
                                if (data) {
                                  applyDBUpdate(data);
                                  if (currentMachineId === m.id) {
                                    setCurrentMachineId(data.machines[0]?.id || null);
                                  }
                                }
                              });
                            }
                          }}
                        >
                          <Trash2 size={14} />
                        </button>
                      )}
                    </div>
                  </div>
                );
              })}

            </div>
          </div>
        ) : currentAppView === 'sessions' ? (
          /* SESSIONS LIST VIEW */
          <div className="combined-jobs-view">
            <header className="dashboard-header" style={{ marginBottom: '1.5rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '1rem' }}>
              <div>
                <h2>จัดการรอบรันทั้งหมด (Sessions Management)</h2>
                <p style={{ color: 'var(--text-secondary)', fontSize: '0.9rem', marginTop: '4px' }}>
                  รายงานสรุป ค้นหา ดาวน์โหลดเอกสาร เปรียบเทียบกราฟ และเริ่มเปิดรันข้อมูลใหม่
                </p>
              </div>

              {/* Header Actions */}
              <div style={{ display: 'flex', gap: '1rem', alignItems: 'center', flexWrap: 'wrap' }}>
                <button
                  className="btn btn-blue"
                  style={{ margin: 0 }}
                  onClick={() => {
                    setNewSessionName(`Session ${jobs.length + 1}`);
                    setNewSessionMachineId(currentMachineId || machines[0]?.id);
                    setNewSessionCustomerId('');
                    setNewSessionCustomerName('');
                    setNewSessionCustomerEmail('');
                    setNewSessionExpiryHours(0);
                    setShowWizardSuccess(false);
                    setShowAddSessionModal(true);
                  }}
                >
                  <PlusCircle size={18} />
                  ตั้งค่าเปิดงานใหม่ (Setup New Run)
                </button>

                {/* Tab Selector */}
                <div style={{ display: 'flex', gap: '0.5rem', background: 'rgba(15, 23, 42, 0.3)', padding: '4px', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                  <button
                    className={`nav-tab ${combinedActiveTab === 'list' ? 'active' : ''}`}
                    onClick={() => setCombinedActiveTab('list')}
                    style={{ padding: '6px 16px', fontSize: '0.85rem' }}
                  >
                    ตารางงานทั้งหมด ({jobs.length})
                  </button>
                  <button
                    className={`nav-tab ${combinedActiveTab === 'compare' ? 'active' : ''}`}
                    onClick={() => setCombinedActiveTab('compare')}
                    style={{ padding: '6px 16px', fontSize: '0.85rem' }}
                  >
                    วิเคราะห์เปรียบเทียบกราฟ
                  </button>
                </div>
              </div>
            </header>

            {combinedActiveTab === 'list' ? (
              /* TAB 1: ALL SESSIONS LIST */
              <div className="glass-panel" style={{ padding: '1.5rem' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.25rem', gap: '1rem', flexWrap: 'wrap' }}>
                  <h3 style={{ fontSize: '1.1rem', fontWeight: 600 }}>รายการรอบบันทึกทั้งหมด</h3>

                  {/* Search Box */}
                  <div style={{ position: 'relative', width: '100%', maxWidth: '300px' }}>
                    <input
                      type="text"
                      placeholder="ค้นหาเครื่องมือ หรือชื่อรอบบันทึก..."
                      value={combinedSearchQuery}
                      onChange={(e) => setCombinedSearchQuery(e.target.value)}
                      style={{
                        width: '100%',
                        padding: '8px 12px 8px 36px',
                        borderRadius: '8px',
                        border: '1px solid var(--border-color)',
                        background: 'rgba(15, 23, 42, 0.5)',
                        color: 'white',
                        fontSize: '0.9rem'
                      }}
                    />
                    <span style={{ position: 'absolute', left: '12px', top: '50%', transform: 'translateY(-50%)', color: 'var(--text-secondary)' }}>🔍</span>
                  </div>
                </div>

                <div className="table-responsive">
                  <table className="data-table">
                    <thead>
                      <tr>
                        <th>เครื่องมือ (Instrument)</th>
                        <th>ชื่อรอบบันทึก (Session Name)</th>
                        <th>วันที่สร้าง (Created At)</th>
                        <th>จำนวนจุดข้อมูล (Data Points)</th>
                        <th style={{ textAlign: 'center' }}>การจัดการ</th>
                      </tr>
                    </thead>
                    <tbody>
                      {(() => {
                        const filteredJobs = jobs.filter(job => {
                          const machineName = machines.find(m => m.id === job.machineId)?.name || '';
                          const query = combinedSearchQuery.toLowerCase();
                          return job.name.toLowerCase().includes(query) || machineName.toLowerCase().includes(query) || job.id.toLowerCase().includes(query);
                        });

                        if (filteredJobs.length === 0) {
                          return (
                            <tr>
                              <td colSpan="6" style={{ textAlign: 'center', padding: '3rem', color: 'var(--text-secondary)' }}>
                                ไม่พบข้อมูลรอบการรันที่ค้นหา
                              </td>
                            </tr>
                          );
                        }

                        return filteredJobs.map(job => {
                          const machine = machines.find(m => m.id === job.machineId);
                          return (
                            <tr key={job.id}>
                              <td style={{ fontWeight: 600 }}>
                                <span style={{ display: 'inline-flex', alignItems: 'center', gap: '8px' }}>
                                  <img src="/bioreactor.png" alt="Bioreactor" style={{ width: '24px', height: '24px', objectFit: 'contain', flexShrink: 0 }} />
                                  {machine?.name || 'Unknown'}
                                </span>
                              </td>
                              <td style={{ color: 'var(--accent-blue)', fontWeight: 600 }}>{job.name}</td>
                              <td>{formatCreatedAt(job.createdAt)}</td>
                              <td>{job.data?.length || 0} จุด</td>
                              <td>
                                <div style={{ display: 'flex', gap: '8px', justifyContent: 'center', alignItems: 'center' }}>
                                  <button
                                    className="btn-outline-yellow"
                                    onClick={() => {
                                      setCurrentMachineId(job.machineId);
                                      setCurrentJobId(job.id);
                                      setIsViewingHistory(job.status === 'finished');
                                      setCurrentAppView('monitoring');
                                      setActiveTab('dashboard');
                                    }}
                                  >
                                    เปิดบอร์ดข้อมูล
                                  </button>
                                  <button
                                    className="btn-outline-green"
                                    onClick={() => exportToExcel(job)}
                                    disabled={!job.data || job.data.length === 0}
                                    title="ดาวน์โหลด Excel สำหรับรอบนี้"
                                  >
                                    <Download size={14} style={{ marginRight: '4px' }} /> Excel
                                  </button>
                                  <button
                                    className="btn-outline-blue"
                                    onClick={() => exportToCSV(job)}
                                    disabled={!job.data || job.data.length === 0}
                                    title="ดาวน์โหลด CSV สำหรับรอบนี้"
                                  >
                                    <Download size={14} style={{ marginRight: '4px' }} /> CSV
                                  </button>
                                  <button
                                    className="delete-row-btn"
                                    onClick={(e) => deleteJob(job.id, e)}
                                    title="Delete Session"
                                    style={{ margin: 0 }}
                                  >
                                    <Trash2 size={16} />
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        });
                      })()}
                    </tbody>
                  </table>
                </div>
              </div>
            ) : (
              /* TAB 2: MULTI-SESSION COMPARE CHART */
              <div style={{ display: 'flex', gap: '1.5rem', flexDirection: 'column' }}>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: '1.5rem' }}>

                  {/* Select Sessions Panel */}
                  <div className="glass-panel" style={{ padding: '1.5rem' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '1rem', marginBottom: '1.25rem' }}>
                      <h3 style={{ fontSize: '1.1rem', fontWeight: 600, margin: 0 }}>1. เลือกรอบการรันเพื่อเปรียบเทียบ</h3>

                      {/* Select Parameter Multi-select Pills */}
                      <div style={{ display: 'flex', alignItems: 'flex-start', gap: '0.75rem', flexWrap: 'wrap' }}>
                        <span style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', whiteSpace: 'nowrap', paddingTop: '5px' }}>ตัวแปรที่วิเคราะห์:</span>
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.5rem' }}>
                          {COMPARE_PARAM_OPTIONS.map(opt => {
                            const isSelected = compareParams.includes(opt.value);
                            return (
                              <button
                                key={opt.value}
                                onClick={() => {
                                  if (isSelected) {
                                    if (compareParams.length > 1) setCompareParams(prev => prev.filter(p => p !== opt.value));
                                  } else {
                                    setCompareParams(prev => [...prev, opt.value]);
                                  }
                                }}
                                style={{
                                  padding: '4px 12px',
                                  borderRadius: '20px',
                                  border: isSelected ? '1px solid var(--accent-blue)' : '1px solid var(--border-color)',
                                  background: isSelected ? 'rgba(59, 130, 246, 0.2)' : 'rgba(255,255,255,0.04)',
                                  color: isSelected ? '#60a5fa' : 'var(--text-secondary)',
                                  fontSize: '0.8rem',
                                  cursor: 'pointer',
                                  fontWeight: isSelected ? 600 : 400,
                                  transition: 'all 0.18s',
                                  whiteSpace: 'nowrap',
                                  display: 'flex',
                                  alignItems: 'center',
                                  gap: '4px',
                                }}
                              >
                                {isSelected && <span style={{ fontSize: '0.7rem' }}>✓</span>}{opt.label}
                              </button>
                            );
                          })}
                        </div>
                      </div>
                    </div>

                    {/* Session Selection Checkbox list */}
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: '0.75rem', maxHeight: '200px', overflowY: 'auto', padding: '0.5rem', background: 'rgba(15, 23, 42, 0.2)', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                      {jobs.map(job => {
                        const machine = machines.find(m => m.id === job.machineId);
                        const isChecked = selectedCompareJobIds.includes(job.id);
                        return (
                          <label
                            key={job.id}
                            style={{
                              display: 'flex',
                              alignItems: 'center',
                              gap: '0.75rem',
                              padding: '8px 12px',
                              borderRadius: '6px',
                              background: isChecked ? 'rgba(59, 130, 246, 0.1)' : 'rgba(255, 255, 255, 0.02)',
                              border: isChecked ? '1px solid rgba(59, 130, 246, 0.3)' : '1px solid var(--border-color)',
                              cursor: 'pointer',
                              fontSize: '0.85rem'
                            }}
                          >
                            <input
                              type="checkbox"
                              checked={isChecked}
                              onChange={() => {
                                if (isChecked) {
                                  setSelectedCompareJobIds(prev => prev.filter(id => id !== job.id));
                                } else {
                                  setSelectedCompareJobIds(prev => [...prev, job.id]);
                                }
                              }}
                              style={{ width: '16px', height: '16px' }}
                            />
                            <div style={{ minWidth: 0 }}>
                              <div style={{ fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{job.name}</div>
                              <div style={{ fontSize: '0.7rem', color: 'var(--text-secondary)' }}>{machine?.name} ({job.data?.length || 0} จุด)</div>
                            </div>
                          </label>
                        );
                      })}
                    </div>
                  </div>

                  {/* Chart Display Panel */}
                  <div ref={compareChartRef} className="glass-panel" style={{ padding: '1.5rem', minHeight: '500px' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.25rem', flexWrap: 'wrap', gap: '0.5rem' }}>
                      <h3 style={{ fontSize: '1.1rem', fontWeight: 600, margin: 0 }}>2. กราฟเปรียบเทียบแนวโน้ม (เทียบชั่วโมงเลี้ยงเชื้อ Culture Hour)</h3>
                      {selectedCompareJobIds.length > 0 && (
                        <div style={{ display: 'flex', gap: '0.5rem' }}>
                          {/* CSV button */}
                          <button onClick={downloadCompareCSV} title="ดาวน์โหลดข้อมูลเป็น CSV"
                            style={{ display: 'flex', alignItems: 'center', gap: '5px', padding: '6px 12px', borderRadius: '8px', border: '1px solid rgba(34,197,94,0.4)', background: 'rgba(34,197,94,0.1)', color: '#4ade80', fontSize: '0.8rem', fontWeight: 600, cursor: 'pointer', transition: 'all 0.18s' }}
                            onMouseEnter={e => { e.currentTarget.style.background = 'rgba(34,197,94,0.25)' }} onMouseLeave={e => { e.currentTarget.style.background = 'rgba(34,197,94,0.1)' }}
                          >
                            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="7 10 12 15 17 10" /><line x1="12" y1="15" x2="12" y2="3" /></svg>
                            CSV
                          </button>
                          {/* PNG button */}
                          <button onClick={downloadChartImage} title="บันทึกกราฟเป็นรูปภาพ PNG"
                            style={{ display: 'flex', alignItems: 'center', gap: '5px', padding: '6px 12px', borderRadius: '8px', border: '1px solid rgba(139,92,246,0.4)', background: 'rgba(139,92,246,0.1)', color: '#a78bfa', fontSize: '0.8rem', fontWeight: 600, cursor: 'pointer', transition: 'all 0.18s' }}
                            onMouseEnter={e => { e.currentTarget.style.background = 'rgba(139,92,246,0.25)' }} onMouseLeave={e => { e.currentTarget.style.background = 'rgba(139,92,246,0.1)' }}
                          >
                            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="3" width="18" height="18" rx="2" /><circle cx="8.5" cy="8.5" r="1.5" /><polyline points="21 15 16 10 5 21" /></svg>
                            PNG
                          </button>
                          {/* PDF button */}
                          <button onClick={downloadChartPDF} title="บันทึกกราฟเป็น PDF"
                            style={{ display: 'flex', alignItems: 'center', gap: '5px', padding: '6px 12px', borderRadius: '8px', border: '1px solid rgba(239,68,68,0.4)', background: 'rgba(239,68,68,0.1)', color: '#f87171', fontSize: '0.8rem', fontWeight: 600, cursor: 'pointer', transition: 'all 0.18s' }}
                            onMouseEnter={e => { e.currentTarget.style.background = 'rgba(239,68,68,0.25)' }} onMouseLeave={e => { e.currentTarget.style.background = 'rgba(239,68,68,0.1)' }}
                          >
                            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" /><line x1="16" y1="13" x2="8" y2="13" /><line x1="16" y1="17" x2="8" y2="17" /><polyline points="10 9 9 9 8 9" /></svg>
                            PDF
                          </button>
                        </div>
                      )}
                    </div>
                    {selectedCompareJobIds.length === 0 ? (
                      <div style={{ display: 'flex', height: '80%', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', color: 'var(--text-secondary)', gap: '0.5rem' }}>
                        <span style={{ fontSize: '2rem' }}>📊</span>
                        <span>กรุณาเลือกอย่างน้อยหนึ่งรอบบันทึกเพื่อแสดงผลเปรียบเทียบกราฟ</span>
                      </div>
                    ) : (
                      <ResponsiveContainer width="100%" height="85%">
                        <LineChart data={(() => {
                          const roundedPoints = {};
                          selectedCompareJobIds.forEach(jobId => {
                            const job = jobs.find(j => j.id === jobId);
                            if (!job) return;
                            const machine = machines.find(m => m.id === job.machineId);
                            const mPrefix = machine ? `${machine.name} - ` : '';
                            const jobLabel = `${mPrefix}${job.name}`;

                            // Find earliest timestamp for cultureHour starting point
                            let minTimeMs = Infinity;
                            if (job.data && job.data.length > 0) {
                              job.data.forEach(row => {
                                if (row.timestamp) {
                                  const t = new Date(row.timestamp).getTime();
                                  if (!isNaN(t) && t < minTimeMs) minTimeMs = t;
                                }
                              });
                            }

                            if (job.data) {
                              job.data.forEach(row => {
                                const t = row.timestamp ? new Date(row.timestamp).getTime() : NaN;
                                const elapsed = isNaN(t) || minTimeMs === Infinity ? 0 : (t - minTimeMs) / 3600000;
                                const roundedElapsed = Math.round(elapsed * 10) / 10;

                                if (!roundedPoints[roundedElapsed]) {
                                  roundedPoints[roundedElapsed] = { cultureHour: roundedElapsed };
                                }

                                compareParams.forEach(param => {
                                  const val = row[param];
                                  if (val === undefined || isNaN(parseFloat(val))) return;
                                  const paramLabel = COMPARE_PARAM_OPTIONS.find(o => o.value === param)?.label || param;
                                  const lineKey = compareParams.length > 1
                                    ? `${jobLabel} [${paramLabel}]`
                                    : jobLabel;
                                  roundedPoints[roundedElapsed][lineKey] = parseFloat(val);
                                });
                              });
                            }
                          });

                          return Object.values(roundedPoints).sort((a, b) => a.cultureHour - b.cultureHour);
                        })()}>
                          <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.1)" />
                          <XAxis dataKey="cultureHour" stroke="var(--text-secondary)" tick={{ fontSize: 12 }} label={{ value: 'ชั่วโมงการเลี้ยงเชื้อ (Culture Hour)', position: 'insideBottomRight', offset: -10, fill: 'var(--text-secondary)', fontSize: 12 }} />
                          <YAxis stroke="var(--text-secondary)" tick={{ fontSize: 12 }} />
                          <Tooltip
                            contentStyle={{ background: '#1e293b', border: '1px solid rgba(255,255,255,0.1)', borderRadius: '8px' }}
                            labelFormatter={(val) => `ชั่วโมงเลี้ยงเชื้อ: ${val} ชม.`}
                          />
                          <Legend wrapperStyle={{ fontSize: '13px', paddingTop: '10px' }} />
                          {(() => {
                            const COMPARE_COLORS = [
                              'var(--accent-red)',
                              'var(--accent-blue)',
                              'var(--accent-green)',
                              'var(--accent-yellow)',
                              'var(--accent-purple)',
                              '#00f0ff',
                              '#ec4899',
                              '#f43f5e',
                              '#14b8a6',
                              '#84cc16',
                              '#fb923c',
                              '#a78bfa',
                            ];
                            const lines = [];
                            let colorIdx = 0;
                            selectedCompareJobIds.forEach(jobId => {
                              const job = jobs.find(j => j.id === jobId);
                              if (!job) return;
                              const machine = machines.find(m => m.id === job.machineId);
                              const mPrefix = machine ? `${machine.name} - ` : '';
                              const jobLabel = `${mPrefix}${job.name}`;
                              compareParams.forEach(param => {
                                const paramLabel = COMPARE_PARAM_OPTIONS.find(o => o.value === param)?.label || param;
                                const lineKey = compareParams.length > 1
                                  ? `${jobLabel} [${paramLabel}]`
                                  : jobLabel;
                                lines.push(
                                  <Line
                                    key={`${jobId}-${param}`}
                                    type="monotone"
                                    dataKey={lineKey}
                                    name={lineKey}
                                    stroke={COMPARE_COLORS[colorIdx % COMPARE_COLORS.length]}
                                    strokeWidth={3}
                                    dot={true}
                                    connectNulls={true}
                                  />
                                );
                                colorIdx++;
                              });
                            });
                            return lines;
                          })()}
                        </LineChart>
                      </ResponsiveContainer>
                    )}
                  </div>

                </div>
              </div>
            )}
          </div>
        ) : currentAppView === 'settings' ? (
          /* SYSTEM SETTINGS VIEW */
          <div className="settings-view">
            <header className="dashboard-header">
              <h2>System Settings (การตั้งค่าระบบ)</h2>
            </header>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: '1.5rem', alignItems: 'start' }}>
              <div className="glass-panel" style={{ padding: '2rem' }}>
                <h3 style={{ marginBottom: '1.5rem', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                  <Settings size={20} color="var(--accent-blue)" /> ตั้งค่ารหัสผ่านผู้ดูแลระบบ (Change Admin Password)
                </h3>

                <form onSubmit={async (e) => {
                  e.preventDefault();
                  const currentPassword = e.target.currentPassword.value;
                  const newPassword = e.target.newPassword.value;
                  const confirmPassword = e.target.confirmPassword.value;

                  if (!currentPassword || !newPassword || !confirmPassword) {
                    alert('กรุณากรอกข้อมูลให้ครบถ้วน');
                    return;
                  }

                  if (newPassword !== confirmPassword) {
                    alert('รหัสผ่านใหม่และยืนยันรหัสผ่านใหม่ไม่ตรงกัน');
                    return;
                  }

                  try {
                    const res = await fetch('/api/settings/update-password', {
                      method: 'POST',
                      headers: { 'Content-Type': 'application/json' },
                      body: JSON.stringify({ currentPassword, newPassword })
                    });
                    const result = await res.json();
                    if (res.ok) {
                      alert('เปลี่ยนรหัสผ่านสำเร็จเรียบร้อยแล้ว');
                      e.target.reset();
                    } else {
                      alert(`ผิดพลาด: ${result.error}`);
                    }
                  } catch (err) {
                    console.error('Error changing password:', err);
                    alert('ไม่สามารถเชื่อมต่อเซิร์ฟเวอร์เพื่อเปลี่ยนรหัสผ่านได้');
                  }
                }} className="data-form" style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>

                  <div className="form-group" style={{ width: '100%' }}>
                    <label>รหัสผ่านปัจจุบัน (Current Password)</label>
                    <input
                      type="password"
                      name="currentPassword"
                      placeholder="ป้อนรหัสผ่านปัจจุบัน"
                      style={{ padding: '10px', borderRadius: '8px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', width: '100%' }}
                    />
                  </div>

                  <div className="form-group" style={{ width: '100%' }}>
                    <label>รหัสผ่านใหม่ (New Password)</label>
                    <input
                      type="password"
                      name="newPassword"
                      placeholder="ป้อนรหัสผ่านใหม่"
                      style={{ padding: '10px', borderRadius: '8px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', width: '100%' }}
                    />
                  </div>

                  <div className="form-group" style={{ width: '100%' }}>
                    <label>ยืนยันรหัสผ่านใหม่ (Confirm New Password)</label>
                    <input
                      type="password"
                      name="confirmPassword"
                      placeholder="ป้อนยืนยันรหัสผ่านใหม่"
                      style={{ padding: '10px', borderRadius: '8px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', width: '100%' }}
                    />
                  </div>

                  <button type="submit" className="btn btn-primary" style={{ width: '100%', margin: 0, marginTop: '0.5rem' }}>
                    บันทึกการตั้งค่า (Update Password)
                  </button>
                </form>
              </div>

              {/* Developer & Version Information Panel */}
              <div className="glass-panel" style={{ padding: '2rem' }}>
                <h3 style={{ marginBottom: '1.5rem', display: 'flex', alignItems: 'center', gap: '0.5rem', color: 'var(--accent-green)' }}>
                  ⚙️ เกี่ยวกับผู้พัฒนา & ระบบ (About System)
                </h3>

                {!isEditingAbout ? (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem', color: 'var(--text-primary)' }}>
                    <div style={{ paddingBottom: '10px', borderBottom: '1px solid rgba(255,255,255,0.06)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <span style={{ color: 'var(--text-secondary)', fontSize: '0.88rem' }}>ชื่อระบบ:</span>
                      <span style={{ fontWeight: 700 }}>{aboutSystem.systemName}</span>
                    </div>
                    <div style={{ paddingBottom: '10px', borderBottom: '1px solid rgba(255,255,255,0.06)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <span style={{ color: 'var(--text-secondary)', fontSize: '0.88rem' }}>เวอร์ชันปัจจุบัน:</span>
                      <span style={{ fontWeight: 700, color: 'var(--accent-blue)' }}>{aboutSystem.systemVersion}</span>
                    </div>
                    <div style={{ paddingBottom: '10px', borderBottom: '1px solid rgba(255,255,255,0.06)', display: 'flex', flexDirection: 'column', gap: '4px' }}>
                      <span style={{ color: 'var(--text-secondary)', fontSize: '0.88rem' }}>ผู้พัฒนาระบบ:</span>
                      <span style={{ fontWeight: 700 }}>{aboutSystem.developer}</span>
                    </div>
                    <div style={{ paddingBottom: '10px', borderBottom: '1px solid rgba(255,255,255,0.06)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <span style={{ color: 'var(--text-secondary)', fontSize: '0.88rem' }}>เทคโนโลยีหลัก:</span>
                      <span style={{ fontWeight: 700, fontSize: '0.85rem' }}>{aboutSystem.techStack}</span>
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                      <span style={{ color: 'var(--text-secondary)', fontSize: '0.88rem' }}>รายละเอียดติดต่อ / สนับสนุนเทคนิค:</span>
                      <span style={{ fontSize: '0.85rem' }}>📧 {aboutSystem.supportEmail}</span>
                      <span style={{ fontSize: '0.85rem' }}>📞 {aboutSystem.supportPhone}</span>
                    </div>

                    {userRole === 'admin' && (
                      <button
                        onClick={() => setIsEditingAbout(true)}
                        className="btn btn-secondary"
                        style={{ width: '100%', margin: 0, marginTop: '1rem' }}
                      >
                        ✏️ แก้ไขข้อมูลผู้พัฒนา & ระบบ
                      </button>
                    )}
                  </div>
                ) : (
                  <form onSubmit={async (e) => {
                    e.preventDefault();
                    const systemName = e.target.systemName.value;
                    const systemVersion = e.target.systemVersion.value;
                    const developer = e.target.developer.value;
                    const techStack = e.target.techStack.value;
                    const supportEmail = e.target.supportEmail.value;
                    const supportPhone = e.target.supportPhone.value;

                    try {
                      const res = await fetch('/api/settings/update-about', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                          systemName,
                          systemVersion,
                          developer,
                          techStack,
                          supportEmail,
                          supportPhone
                        })
                      });
                      const result = await res.json();
                      if (res.ok) {
                        alert('อัปเดตข้อมูลผู้พัฒนาและระบบสำเร็จเรียบร้อยแล้ว');
                        setAboutSystem(result.settings);
                        setIsEditingAbout(false);
                      } else {
                        alert(`ผิดพลาด: ${result.error}`);
                      }
                    } catch (err) {
                      console.error('Error updating developer settings:', err);
                      alert('ไม่สามารถเชื่อมต่อเซิร์ฟเวอร์เพื่อบันทึกข้อมูลได้');
                    }
                  }} style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>

                    <div className="form-group" style={{ width: '100%' }}>
                      <label style={{ fontSize: '0.78rem' }}>ชื่อระบบ (System Name)</label>
                      <input
                        type="text"
                        name="systemName"
                        defaultValue={aboutSystem.systemName}
                        style={{ padding: '8px', borderRadius: '6px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', width: '100%' }}
                        required
                      />
                    </div>

                    <div className="form-group" style={{ width: '100%' }}>
                      <label style={{ fontSize: '0.78rem' }}>เวอร์ชันปัจจุบัน (System Version)</label>
                      <input
                        type="text"
                        name="systemVersion"
                        defaultValue={aboutSystem.systemVersion}
                        style={{ padding: '8px', borderRadius: '6px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', width: '100%' }}
                        required
                      />
                    </div>

                    <div className="form-group" style={{ width: '100%' }}>
                      <label style={{ fontSize: '0.78rem' }}>ผู้พัฒนาระบบ (Developer Team)</label>
                      <input
                        type="text"
                        name="developer"
                        defaultValue={aboutSystem.developer}
                        style={{ padding: '8px', borderRadius: '6px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', width: '100%' }}
                        required
                      />
                    </div>

                    <div className="form-group" style={{ width: '100%' }}>
                      <label style={{ fontSize: '0.78rem' }}>เทคโนโลยีหลัก (Core Tech)</label>
                      <input
                        type="text"
                        name="techStack"
                        defaultValue={aboutSystem.techStack}
                        style={{ padding: '8px', borderRadius: '6px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', width: '100%' }}
                        required
                      />
                    </div>

                    <div className="form-group" style={{ width: '100%' }}>
                      <label style={{ fontSize: '0.78rem' }}>อีเมลสนับสนุน (Support Email)</label>
                      <input
                        type="email"
                        name="supportEmail"
                        defaultValue={aboutSystem.supportEmail}
                        style={{ padding: '8px', borderRadius: '6px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', width: '100%' }}
                        required
                      />
                    </div>

                    <div className="form-group" style={{ width: '100%' }}>
                      <label style={{ fontSize: '0.78rem' }}>เบอร์โทรศัพท์ (Support Phone)</label>
                      <input
                        type="text"
                        name="supportPhone"
                        defaultValue={aboutSystem.supportPhone}
                        style={{ padding: '8px', borderRadius: '6px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', width: '100%' }}
                        required
                      />
                    </div>

                    <div style={{ display: 'flex', gap: '0.5rem', marginTop: '0.5rem' }}>
                      <button
                        type="button"
                        onClick={() => setIsEditingAbout(false)}
                        className="btn btn-secondary"
                        style={{ flex: 1, margin: 0 }}
                      >
                        ยกเลิก
                      </button>
                      <button
                        type="submit"
                        className="btn btn-blue"
                        style={{ flex: 1, margin: 0 }}
                      >
                        💾 บันทึกข้อมูล
                      </button>
                    </div>
                  </form>
                )}
              </div>

              {/* VVM Calculation Settings Panel */}
              <div className="glass-panel" style={{ padding: '2rem' }}>
                <h3 style={{ marginBottom: '1.5rem', display: 'flex', alignItems: 'center', gap: '0.5rem', color: 'var(--accent-blue)' }}>
                  💨 ตั้งค่าการคำนวณ VVM (VVM Calculation)
                </h3>

                <form onSubmit={async (e) => {
                  e.preventDefault();
                  const vvmCalcType = e.target.vvmCalcType.value;
                  const maxVolumeLiters = e.target.maxVolumeLiters ? e.target.maxVolumeLiters.value : aboutSystem.maxVolumeLiters;
                  const constantVolumeLiters = e.target.constantVolumeLiters ? e.target.constantVolumeLiters.value : aboutSystem.constantVolumeLiters;
                  const airUnit = e.target.airUnit.value;

                  try {
                    const res = await fetch('/api/settings/update-vvm', {
                      method: 'POST',
                      headers: { 'Content-Type': 'application/json' },
                      body: JSON.stringify({
                        vvmCalcType,
                        maxVolumeLiters,
                        constantVolumeLiters,
                        airUnit
                      })
                    });
                    const result = await res.json();
                    if (res.ok) {
                      alert('บันทึกการตั้งค่า VVM สำเร็จเรียบร้อยแล้ว');
                      setAboutSystem(result.settings);
                    } else {
                      alert(`ผิดพลาด: ${result.error}`);
                    }
                  } catch (err) {
                    console.error('Error saving VVM settings:', err);
                    alert('ไม่สามารถเชื่อมต่อเซิร์ฟเวอร์เพื่อบันทึกข้อมูลได้');
                  }
                }} style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>

                  <div className="form-group" style={{ width: '100%' }}>
                    <label>หน่วยของอัตราไหลลมในฐานข้อมูล (Air Flow Input Unit)</label>
                    <select
                      name="airUnit"
                      value={aboutSystem.airUnit || 'mlmin'}
                      onChange={(e) => {
                        setAboutSystem(prev => ({ ...prev, airUnit: e.target.value }));
                      }}
                      style={{ padding: '10px', borderRadius: '8px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', width: '100%' }}
                    >
                      <option value="mlmin">mL/min (มิลลิลิตร/นาที) - เช่น ป้อน 400 = 0.4 ลิตร/นาที</option>
                      <option value="lmin">L/min หรือ SLPM (ลิตร/นาที) - เช่น ป้อน 4.0 = 4.0 ลิตร/นาที</option>
                    </select>
                    <span style={{ fontSize: '0.72rem', color: 'var(--text-secondary)', marginTop: '4px' }}>
                      หากเลือก mL/min ระบบจะหาร 1000 ให้อัตโนมัติเพื่อใช้เป็นลิตร/นาทีในสูตรคำนวณ VVM
                    </span>
                  </div>

                  <div className="form-group" style={{ width: '100%' }}>
                    <label>วิธีการคำนวณปริมาตรน้ำหมัก (Fermentation Volume Method)</label>
                    <select
                      name="vvmCalcType"
                      value={aboutSystem.vvmCalcType || 'dynamic'}
                      onChange={(e) => {
                        setAboutSystem(prev => ({ ...prev, vvmCalcType: e.target.value }));
                      }}
                      style={{ padding: '10px', borderRadius: '8px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', width: '100%' }}
                    >
                      <option value="dynamic">คำนวณตามระดับน้ำจริงในถัง (Dynamic Level %)</option>
                      <option value="constant">คิดเป็นค่าคงที่ (Constant Volume)</option>
                    </select>
                  </div>

                  {(aboutSystem.vvmCalcType || 'dynamic') === 'dynamic' ? (
                    <div className="form-group" style={{ width: '100%' }}>
                      <label>ปริมาตรของน้ำหมักสูงสุดที่ระดับ 100% (ลิตร)</label>
                      <input
                        type="number"
                        name="maxVolumeLiters"
                        step="0.1"
                        min="0.1"
                        value={aboutSystem.maxVolumeLiters !== undefined && aboutSystem.maxVolumeLiters !== null ? aboutSystem.maxVolumeLiters : ''}
                        onChange={(e) => {
                          const val = e.target.value;
                          setAboutSystem(prev => ({ ...prev, maxVolumeLiters: val }));
                        }}
                        style={{ padding: '10px', borderRadius: '8px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', width: '100%' }}
                        required
                      />
                      <span style={{ fontSize: '0.72rem', color: 'var(--text-secondary)', marginTop: '4px' }}>
                        สูตร: SLPM ÷ ((Level % ÷ 100) × ปริมาตรสูงสุด)
                      </span>
                    </div>
                  ) : (
                    <div className="form-group" style={{ width: '100%' }}>
                      <label>ปริมาตรน้ำหมักคงที่ (Constant Volume in Liters)</label>
                      <input
                        type="number"
                        name="constantVolumeLiters"
                        step="0.1"
                        min="0.1"
                        value={aboutSystem.constantVolumeLiters !== undefined && aboutSystem.constantVolumeLiters !== null ? aboutSystem.constantVolumeLiters : ''}
                        onChange={(e) => {
                          const val = e.target.value;
                          setAboutSystem(prev => ({ ...prev, constantVolumeLiters: val }));
                        }}
                        style={{ padding: '10px', borderRadius: '8px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', width: '100%' }}
                        required
                      />
                      <span style={{ fontSize: '0.72rem', color: 'var(--text-secondary)', marginTop: '4px' }}>
                        สูตร: SLPM ÷ ปริมาตรน้ำหมักคงที่
                      </span>
                    </div>
                  )}

                  {userRole === 'admin' && (
                    <button type="submit" className="btn btn-primary" style={{ width: '100%', margin: 0, marginTop: '0.5rem' }}>
                      💾 บันทึกการตั้งค่า VVM
                    </button>
                  )}
                </form>
              </div>
            </div>
          </div>
        ) : currentAppView === 'feedbacks' ? (
          /* CUSTOMER FEEDBACKS VIEW FOR ADMIN */
          <div className="feedbacks-view">
            <header className="dashboard-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '1rem', marginBottom: '1.5rem' }}>
              <h2 style={{ margin: 0 }}>แบบประเมินความพึงพอใจลูกค้า (Customer Satisfaction Survey)</h2>
              <button
                onClick={handleExportFeedback}
                className="btn btn-primary"
                style={{ margin: 0, display: 'flex', alignItems: 'center', gap: '6px' }}
              >
                📥 ส่งออกข้อมูล (Export CSV)
              </button>
            </header>

            {(() => {
              const total = feedbacks.length;
              const avg = total > 0 ? (feedbacks.reduce((sum, f) => sum + (f.avgScore || f.rating || 0), 0) / total).toFixed(2) : '0.00';

              // Section averages
              const sectionDefs = [
                { name: 'ด้านกระบวนการให้บริการ', keys: ['q1', 'q2', 'q3'] },
                { name: 'ด้านเจ้าหน้าที่ผู้ให้บริการ', keys: ['q4', 'q5', 'q6'] },
                { name: 'ด้านการประชาสัมพันธ์', keys: ['q7', 'q8'] },
                { name: 'ด้านสถานที่และสิ่งอำนวยความสะดวก', keys: ['q9', 'q10'] },
                { name: 'ด้านคุณภาพการให้บริการ', keys: ['q11'] },
              ];

              const sectionAvgs = sectionDefs.map(sd => {
                const allVals = [];
                feedbacks.forEach(f => {
                  if (f.scores) {
                    sd.keys.forEach(k => { if (f.scores[k]) allVals.push(Number(f.scores[k])); });
                  }
                });
                const avg = allVals.length > 0 ? (allVals.reduce((a, b) => a + b, 0) / allVals.length) : null;
                return { name: sd.name, avg };
              });

              // Channel counts
              const channelCounts = {};
              feedbacks.forEach(f => { (f.channels || []).forEach(ch => { channelCounts[ch] = (channelCounts[ch] || 0) + 1; }); });
              const maxChCount = Math.max(1, ...Object.values(channelCounts));

              // Tool counts
              const toolCounts = {};
              feedbacks.forEach(f => { (f.tools || []).forEach(t => { toolCounts[t] = (toolCounts[t] || 0) + 1; }); });
              const maxToolCount = Math.max(1, ...Object.values(toolCounts));

              const scoreColor = (v) => {
                if (!v) return 'var(--text-secondary)';
                if (v >= 4.5) return '#22c55e';
                if (v >= 3.5) return '#84cc16';
                if (v >= 2.5) return '#eab308';
                if (v >= 1.5) return '#f97316';
                return '#ef4444';
              };

              return (
                <>
                  {/* Summary Stats */}
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '1.25rem', marginBottom: '1.75rem' }}>
                    <div className="stat-card" style={{ padding: '1.25rem', display: 'flex', flexDirection: 'column', gap: '6px' }}>
                      <span style={{ fontSize: '0.82rem', color: 'var(--text-secondary)' }}>คะแนนเฉลี่ยรวม (CSAT)</span>
                      <span style={{ fontSize: '2.2rem', fontWeight: 700, color: 'var(--accent-yellow)', display: 'flex', alignItems: 'baseline', gap: '6px' }}>
                        {avg} <span style={{ fontSize: '1.1rem', color: 'var(--text-secondary)', fontWeight: 500 }}>/ 5.00</span>
                      </span>
                      <span style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>
                        {"★".repeat(Math.round(avg)) + "☆".repeat(5 - Math.round(avg))}
                      </span>
                    </div>

                    <div className="stat-card" style={{ padding: '1.25rem', display: 'flex', flexDirection: 'column', gap: '6px' }}>
                      <span style={{ fontSize: '0.82rem', color: 'var(--text-secondary)' }}>จำนวนผู้ประเมิน</span>
                      <span style={{ fontSize: '2.2rem', fontWeight: 700, color: 'var(--accent-blue)' }}>
                        {total} <span style={{ fontSize: '1rem', fontWeight: 500, color: 'var(--text-secondary)' }}>ราย</span>
                      </span>
                      <span style={{ fontSize: '0.78rem', color: 'var(--text-secondary)' }}>อัปเดตแบบเรียลไทม์ (Real-time Sync)</span>
                    </div>

                    {/* Section Averages */}
                    <div className="stat-card" style={{ padding: '1.25rem', display: 'flex', flexDirection: 'column', gap: '8px', gridColumn: 'span 2' }}>
                      <span style={{ fontSize: '0.82rem', fontWeight: 700, color: 'var(--text-secondary)', marginBottom: '4px' }}>คะแนนเฉลี่ยรายด้าน</span>
                      {sectionAvgs.map((s, i) => (
                        <div key={i} style={{ display: 'flex', alignItems: 'center', gap: '10px', fontSize: '0.8rem' }}>
                          <span style={{ flex: 1, color: '#d1d5db', fontSize: '0.78rem' }}>{s.name}</span>
                          <div style={{ width: '120px', height: '6px', background: 'rgba(255,255,255,0.06)', borderRadius: '3px', overflow: 'hidden' }}>
                            <div style={{ width: `${s.avg ? (s.avg / 5) * 100 : 0}%`, height: '100%', background: scoreColor(s.avg), borderRadius: '3px', transition: 'width 0.4s' }} />
                          </div>
                          <span style={{ width: '36px', textAlign: 'right', fontWeight: 700, color: scoreColor(s.avg) }}>
                            {s.avg ? s.avg.toFixed(2) : '—'}
                          </span>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Channel & Tool Summary */}
                  {(Object.keys(channelCounts).length > 0 || Object.keys(toolCounts).length > 0) && (
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1.25rem', marginBottom: '1.75rem' }}>
                      {/* Channels */}
                      {Object.keys(channelCounts).length > 0 && (
                        <div className="stat-card" style={{ padding: '1.25rem' }}>
                          <div style={{ fontSize: '0.82rem', fontWeight: 700, color: 'var(--accent-blue)', marginBottom: '10px' }}>📡 ช่องทางที่รู้จัก FTC</div>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                            {Object.entries(channelCounts).sort((a, b) => b[1] - a[1]).map(([ch, cnt]) => (
                              <div key={ch} style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '0.8rem' }}>
                                <span style={{ flex: 1, color: '#d1d5db' }}>{ch}</span>
                                <div style={{ width: '80px', height: '6px', background: 'rgba(255,255,255,0.06)', borderRadius: '3px', overflow: 'hidden' }}>
                                  <div style={{ width: `${(cnt / maxChCount) * 100}%`, height: '100%', background: 'var(--accent-blue)', borderRadius: '3px' }} />
                                </div>
                                <span style={{ width: '24px', textAlign: 'right', color: 'var(--accent-blue)', fontWeight: 700 }}>{cnt}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                      {/* Tools */}
                      {Object.keys(toolCounts).length > 0 && (
                        <div className="stat-card" style={{ padding: '1.25rem' }}>
                          <div style={{ fontSize: '0.82rem', fontWeight: 700, color: 'var(--accent-green)', marginBottom: '10px' }}>🔬 เครื่องมือที่ใช้บริการ</div>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                            {Object.entries(toolCounts).sort((a, b) => b[1] - a[1]).map(([t, cnt]) => (
                              <div key={t} style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '0.8rem' }}>
                                <span style={{ flex: 1, color: '#d1d5db' }}>{t}</span>
                                <div style={{ width: '80px', height: '6px', background: 'rgba(255,255,255,0.06)', borderRadius: '3px', overflow: 'hidden' }}>
                                  <div style={{ width: `${(cnt / maxToolCount) * 100}%`, height: '100%', background: 'var(--accent-green)', borderRadius: '3px' }} />
                                </div>
                                <span style={{ width: '24px', textAlign: 'right', color: 'var(--accent-green)', fontWeight: 700 }}>{cnt}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
                  )}
                </>
              );
            })()}

            {feedbacks.length === 0 ? (
              <div style={{ textAlign: 'center', padding: '3rem 0', color: 'var(--text-secondary)' }}>
                ยังไม่มีลูกค้าประเมินความพึงพอใจเข้ามาในระบบ
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
                {feedbacks.map((fb) => {
                  const displayScore = fb.avgScore || fb.rating || 0;
                  const scoreC = displayScore >= 4.5 ? '#22c55e' : displayScore >= 3.5 ? '#84cc16' : displayScore >= 2.5 ? '#eab308' : '#f97316';
                  const scoreLabel = displayScore >= 4.5 ? 'มากที่สุด' : displayScore >= 3.5 ? 'มาก' : displayScore >= 2.5 ? 'ปานกลาง' : displayScore >= 1.5 ? 'น้อย' : 'น้อยที่สุด';
                  return (
                    <div key={fb.id} style={{
                      background: 'rgba(255,255,255,0.02)',
                      border: '1px solid rgba(255,255,255,0.06)',
                      borderRadius: '12px',
                      padding: '1.25rem',
                      display: 'flex', flexDirection: 'column', gap: '10px'
                    }}>
                      {/* Card Header */}
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', flexWrap: 'wrap', gap: '8px' }}>
                        <div>
                          <span style={{ fontWeight: 600, color: 'var(--accent-blue)', marginRight: '10px' }}>
                            📁 {fb.jobName || 'Unknown'}
                          </span>
                          <span style={{ fontSize: '0.78rem', color: 'var(--text-secondary)' }}>{formatCreatedAt(fb.createdAt)}</span>
                        </div>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                          <div style={{
                            display: 'flex', alignItems: 'center', gap: '6px',
                            background: scoreC + '22', border: `1px solid ${scoreC}55`,
                            borderRadius: '20px', padding: '3px 10px'
                          }}>
                            <span style={{ color: scoreC, fontWeight: 700, fontSize: '0.9rem' }}>★ {Number(displayScore).toFixed(2)}</span>
                            <span style={{ fontSize: '0.75rem', color: scoreC }}>{scoreLabel}</span>
                          </div>
                          <button
                            className="delete-row-btn"
                            style={{ margin: 0, padding: '4px' }}
                            onClick={() => deleteFeedback(fb.id)}
                            title="Delete feedback"
                          >
                            <Trash2 size={15} />
                          </button>
                        </div>
                      </div>

                      {/* Section Scores (if new format) */}
                      {fb.scores && (
                        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(180px, 1fr))', gap: '4px' }}>
                          {Object.entries(fb.scores).map(([k, v]) => {
                            const qIdx = parseInt(k.replace('q', '')) - 1;
                            const qObj = [
                              'ขั้นตอนการให้บริการ', 'ช่องทางการติดต่อ', 'ระยะเวลาให้บริการ',
                              'บุคลิกภาพเจ้าหน้าที่', 'คำแนะนำของเจ้าหน้าที่', 'ความรู้เจ้าหน้าที่',
                              'สื่อประชาสัมพันธ์', 'ช่องทาง PR', 'สถานที่ให้บริการ',
                              'สิ่งอำนวยความสะดวก', 'คุณภาพบริการ'
                            ];
                            const label = qObj[qIdx] || k;
                            const score = Number(v);
                            const c = score >= 5 ? '#22c55e' : score >= 4 ? '#84cc16' : score >= 3 ? '#eab308' : score >= 2 ? '#f97316' : '#ef4444';
                            return (
                              <div key={k} style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '0.75rem' }}>
                                <span style={{ color: 'var(--text-secondary)', flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={label}>{label}</span>
                                <span style={{ fontWeight: 700, color: c, minWidth: '16px', textAlign: 'right' }}>{score}</span>
                              </div>
                            );
                          })}
                        </div>
                      )}

                      {/* Channels & Tools used */}
                      {((fb.channels && fb.channels.length > 0) || (fb.tools && fb.tools.length > 0)) && (
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px' }}>
                          {(fb.channels || []).map(ch => (
                            <span key={ch} style={{ padding: '2px 8px', borderRadius: '12px', fontSize: '0.72rem', background: 'rgba(59,130,246,0.12)', border: '1px solid rgba(59,130,246,0.3)', color: '#93c5fd' }}>📡 {ch}</span>
                          ))}
                          {(fb.tools || []).map(t => (
                            <span key={t} style={{ padding: '2px 8px', borderRadius: '12px', fontSize: '0.72rem', background: 'rgba(34,197,94,0.1)', border: '1px solid rgba(34,197,94,0.25)', color: '#86efac' }}>🔬 {t}</span>
                          ))}
                        </div>
                      )}

                      {/* Suggestion / Comment */}
                      {(fb.suggestion || fb.comment) && (
                        <div style={{ fontSize: '0.85rem', color: '#cbdce0', fontStyle: 'italic', paddingTop: '4px', borderTop: '1px solid rgba(255,255,255,0.05)' }}>
                          💬 "{fb.suggestion || fb.comment}"
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>

        ) : currentAppView === 'manual_entry' ? (
          /* MANUAL DATA ENTRY VIEW */
          <div className="manual-entry-view">
            <header className="dashboard-header">
              <h2>Manual Data Entry</h2>
            </header>

            {!currentJob ? (
              <div className="glass-panel empty-state">
                <FolderOpen size={48} opacity={0.5} style={{ marginBottom: '1rem', color: 'var(--text-secondary)' }} />
                <h2>กรุณาเลือกหรือสร้างรอบรัน (Session) ก่อนทำการบันทึกข้อมูล</h2>
              </div>
            ) : (
              <div className={`glass-panel form-container ${currentJob.status === 'stopped' ? 'machine-stopped-mode' : ''}`}>
                <h2 className="form-title">Manual Data Entry (Record Data {showStandby ? 'Standby' : (currentMachine?.name || 'เครื่องมือ')})</h2>
                {currentJob.status === 'stopped' && (
                  <div className="stopped-warning-banner">
                    <span className="warning-icon">⚠️</span>
                    <span>ขณะนี้เครื่องหยุดทำงานอยู่ (Machine is currently STOPPED)</span>
                  </div>
                )}
                {showStandby && (
                  <div className="stopped-warning-banner" style={{ background: 'rgba(59, 130, 246, 0.1)', borderColor: 'rgba(59, 130, 246, 0.25)', color: '#93c5fd' }}>
                    <span className="warning-icon">🏁</span>
                    <span>งานเสร็จสิ้นแล้ว เครื่องอยู่ในสถานะ Standby (Job has finished. Machine is on STANDBY)</span>
                  </div>
                )}
                {isViewingHistory && (
                  <div className="stopped-warning-banner" style={{ background: 'rgba(148, 163, 184, 0.1)', borderColor: 'rgba(148, 163, 184, 0.25)', color: '#94a3b8' }}>
                    <span className="warning-icon">📁</span>
                    <span>ขณะนี้คุณกำลังดูข้อมูลย้อนหลังของรอบรันนี้ (Viewing historical data for this session)</span>
                  </div>
                )}
                <form onSubmit={handleManualSubmit} className="data-form">
                  <div className="form-group-container">
                    <div className="form-group">
                      <label>TEMP (°C)</label>
                      <div className="form-inputs-row">
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">ตั้งค่า (SV)</span>
                          <input type="number" step="0.01" name="temp_set" value={formData.temp_set} onChange={handleInputChange} onFocus={handleInputFocus} onBlur={handleInputBlur} disabled={currentJob?.status === 'finished'} />
                        </div>
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">อ่านค่า (PV)</span>
                          <input type="number" step="0.01" name="temp_read" value={formData.temp_read} onChange={handleInputChange} onFocus={handleInputFocus} onBlur={handleInputBlur} disabled={currentJob?.status === 'finished'} />
                        </div>
                      </div>
                    </div>
                    <div className="form-group">
                      <label>pH</label>
                      <div className="form-inputs-row">
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">ตั้งค่า (SV)</span>
                          <input type="number" step="0.01" name="ph_set" value={formData.ph_set} onChange={handleInputChange} onFocus={handleInputFocus} onBlur={handleInputBlur} disabled={currentJob?.status === 'finished'} />
                        </div>
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">อ่านค่า (PV)</span>
                          <input type="number" step="0.01" name="ph_read" value={formData.ph_read} onChange={handleInputChange} onFocus={handleInputFocus} onBlur={handleInputBlur} disabled={currentJob?.status === 'finished'} />
                        </div>
                      </div>
                    </div>
                    <div className="form-group">
                      <label>DO (%)</label>
                      <div className="form-inputs-row">
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">ตั้งค่า (SV)</span>
                          <input type="number" step="1" name="do_set" value={formData.do_set} onChange={handleInputChange} onFocus={handleInputFocus} onBlur={handleInputBlur} disabled={currentJob?.status === 'finished'} />
                        </div>
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">อ่านค่า (PV)</span>
                          <input type="number" step="1" name="do_read" value={formData.do_read} onChange={handleInputChange} onFocus={handleInputFocus} onBlur={handleInputBlur} disabled={currentJob?.status === 'finished'} />
                        </div>
                      </div>
                    </div>
                    <div className="form-group">
                      <label>AGIT (RPM)</label>
                      <div className="form-inputs-row">
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">ตั้งค่า (SV)</span>
                          <input type="number" step="1" name="agit_set" value={formData.agit_set} onChange={handleInputChange} onFocus={handleInputFocus} onBlur={handleInputBlur} disabled={currentJob?.status === 'finished'} />
                        </div>
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">อ่านค่า (PV)</span>
                          <input type="number" step="1" name="agit_read" value={formData.agit_read} onChange={handleInputChange} onFocus={handleInputFocus} onBlur={handleInputBlur} disabled={currentJob?.status === 'finished'} />
                        </div>
                      </div>
                    </div>
                    <div className="form-group">
                      <label>AIR (L/M)</label>
                      <div className="form-inputs-row">
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">ตั้งค่า (SV)</span>
                          <input type="number" step="0.1" name="air_set" value={formData.air_set} onChange={handleInputChange} onFocus={handleInputFocus} onBlur={handleInputBlur} disabled={currentJob?.status === 'finished'} />
                        </div>
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">อ่านค่า (PV)</span>
                          <input type="number" step="0.1" name="air_read" value={formData.air_read} onChange={handleInputChange} onFocus={handleInputFocus} onBlur={handleInputBlur} disabled={currentJob?.status === 'finished'} />
                        </div>
                      </div>
                    </div>
                    <div className="form-group">
                      <label>VOLUME (L)
                        <span style={{ fontSize: '0.72rem', color: 'var(--text-secondary)', fontWeight: 'normal', marginLeft: '6px' }}>
                          (ถัง {maxVol} L)
                        </span>
                      </label>
                      <div className="form-inputs-row">
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">ตั้งค่า (SV)</span>
                          <input
                            type="number"
                            step="0.1"
                            min="0"
                            max={maxVol}
                            name="level_set"
                            value={typeof formData.level_set === 'number' && !isNaN(formData.level_set) ? parseFloat(((formData.level_set / 100) * maxVol).toFixed(1)) : (formData.level_set === 0 ? 0 : '')}
                            onChange={(e) => {
                              const lit = e.target.value === '' ? '' : parseFloat(e.target.value);
                              setFormData(prev => ({ ...prev, level_set: lit === '' ? '' : (lit / maxVol) * 100 }));
                            }}
                            onFocus={(e) => { e.target.select(); if (!e.target.value || parseFloat(e.target.value) === 0) setFormData(prev => ({ ...prev, level_set: '' })); }}
                            onBlur={(e) => {
                              let v = parseFloat(e.target.value);
                              if (isNaN(v) || e.target.value === '') v = 0;
                              v = Math.min(maxVol, Math.max(0, v));
                              setFormData(prev => ({ ...prev, level_set: (v / maxVol) * 100 }));
                            }}
                            disabled={currentJob?.status === 'finished'}
                          />
                        </div>
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">อ่านค่า (PV)</span>
                          <input
                            type="number"
                            step="0.1"
                            min="0"
                            max={maxVol}
                            name="level_read"
                            value={typeof formData.level_read === 'number' && !isNaN(formData.level_read) ? parseFloat(((formData.level_read / 100) * maxVol).toFixed(1)) : (formData.level_read === 0 ? 0 : '')}
                            onChange={(e) => {
                              const lit = e.target.value === '' ? '' : parseFloat(e.target.value);
                              setFormData(prev => ({ ...prev, level_read: lit === '' ? '' : (lit / maxVol) * 100 }));
                            }}
                            onFocus={(e) => { e.target.select(); if (!e.target.value || parseFloat(e.target.value) === 0) setFormData(prev => ({ ...prev, level_read: '' })); }}
                            onBlur={(e) => {
                              let v = parseFloat(e.target.value);
                              if (isNaN(v) || e.target.value === '') v = 0;
                              v = Math.min(maxVol, Math.max(0, v));
                              setFormData(prev => ({ ...prev, level_read: (v / maxVol) * 100 }));
                            }}
                            disabled={currentJob?.status === 'finished'}
                          />
                        </div>
                      </div>
                    </div>
                    <div className="form-group">
                      <label>AIR OUT (PMa)</label>
                      <div className="form-inputs-row">
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">ตั้งค่า (SV)</span>
                          <input type="number" step="0.1" name="air_out_set" value={formData.air_out_set} onChange={handleInputChange} onFocus={handleInputFocus} onBlur={handleInputBlur} disabled={currentJob?.status === 'finished'} />
                        </div>
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">อ่านค่า (PV)</span>
                          <input type="number" step="0.1" name="air_out_read" value={formData.air_out_read} onChange={handleInputChange} onFocus={handleInputFocus} onBlur={handleInputBlur} disabled={currentJob?.status === 'finished'} />
                        </div>
                      </div>
                    </div>
                    <div className="form-group">
                      <label>HEAT (%)</label>
                      <div className="form-inputs-row">
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">ตั้งค่า (SV)</span>
                          <input type="number" step="1" name="heat_set" value={formData.heat_set} onChange={handleInputChange} onFocus={handleInputFocus} onBlur={handleInputBlur} disabled={currentJob?.status === 'finished'} />
                        </div>
                        <div className="form-input-subgroup">
                          <span className="input-sublabel">อ่านค่า (PV)</span>
                          <input type="number" step="1" name="heat_read" value={formData.heat_read} onChange={handleInputChange} onFocus={handleInputFocus} onBlur={handleInputBlur} disabled={currentJob?.status === 'finished'} />
                        </div>
                      </div>
                    </div>
                    {/* Date and Time manual inputs */}
                    <div className="form-group" style={{ minWidth: '180px' }}>
                      <label>📅 DATE (วันที่บันทึก)</label>
                      <input
                        type="date"
                        name="date"
                        value={formData.date}
                        onChange={handleInputChange}
                        disabled={currentJob?.status === 'finished'}
                        style={{
                          width: '100%',
                          padding: '10px 14px',
                          borderRadius: '8px',
                          border: '1px solid var(--border-color)',
                          background: 'rgba(15, 23, 42, 0.5)',
                          color: 'white',
                          fontFamily: 'inherit',
                          fontSize: '0.95rem'
                        }}
                      />
                    </div>
                    <div className="form-group" style={{ minWidth: '180px' }}>
                      <label>⏰ TIME (เวลาที่บันทึก)</label>
                      <input
                        type="time"
                        name="time"
                        value={formData.time}
                        onChange={handleInputChange}
                        disabled={currentJob?.status === 'finished'}
                        style={{
                          width: '100%',
                          padding: '10px 14px',
                          borderRadius: '8px',
                          border: '1px solid var(--border-color)',
                          background: 'rgba(15, 23, 42, 0.5)',
                          color: 'white',
                          fontFamily: 'inherit',
                          fontSize: '0.95rem'
                        }}
                      />
                    </div>

                    {/* Remarks/Notes full-width input */}
                    <div className="form-group" style={{ flex: '1 1 100%' }}>
                      <label>REMARKS / NOTES (บันทึกข้อความ)</label>
                      <input
                        type="text"
                        name="remark"
                        placeholder="ระบุหมายเหตุหรือข้อความบันทึกที่นี่ (เช่น ปรับค่าอัตราไหล, ตรวจสภาพโพรบ)..."
                        value={formData.remark}
                        onChange={handleInputChange}
                        disabled={currentJob?.status === 'finished'}
                        style={{
                          width: '100%',
                          padding: '10px 14px',
                          borderRadius: '8px',
                          border: '1px solid var(--border-color)',
                          background: 'rgba(15, 23, 42, 0.5)',
                          color: 'white',
                          textAlign: 'left',
                          fontFamily: 'inherit',
                          fontSize: '0.95rem'
                        }}
                      />
                    </div>
                  </div>
                  <button type="submit" className="btn btn-primary" style={{ minWidth: '160px', margin: 0 }} disabled={currentJob?.status === 'finished'}>
                    <PlusCircle size={18} style={{ marginRight: '8px' }} /> Add Record
                  </button>
                </form>
              </div>
            )}
          </div>
        ) : (
          /* MONITORING VIEW */
          <>
            {isViewingHistory && (
              <div className="stopped-warning-banner" style={{ background: 'rgba(148, 163, 184, 0.1)', borderColor: 'rgba(148, 163, 184, 0.25)', color: '#94a3b8', marginBottom: '1rem', width: '100%' }}>
                <span className="warning-icon">📁</span>
                <span>ขณะนี้คุณกำลังดูข้อมูลย้อนหลังของรอบรันนี้ (Viewing historical data for this session)</span>
              </div>
            )}
            <header className="dashboard-header" style={{ flexWrap: 'wrap', gap: '1.5rem' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '1.5rem', flexWrap: 'wrap' }}>
                {/* Instrument and Session static labels (no dropdowns as requested) */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '1.5rem', flexWrap: 'wrap' }}>
                  <h2 style={{ fontSize: '1.1rem', fontWeight: 700, margin: 0, display: 'flex', alignItems: 'center', gap: '6px' }}>
                    🖥️ <span style={{ color: 'var(--text-secondary)', fontSize: '0.85rem', fontWeight: 500 }}>เครื่องมือ:</span>{' '}
                    <span style={{ color: showStandby ? 'var(--text-secondary)' : 'white' }}>
                      {showStandby ? 'Standby' : (currentMachine?.name || '—')}
                    </span>
                  </h2>

                  <h2 style={{ fontSize: '1.1rem', fontWeight: 700, margin: 0, display: 'flex', alignItems: 'center', gap: '6px' }}>
                    📁 <span style={{ color: 'var(--text-secondary)', fontSize: '0.85rem', fontWeight: 500 }}>รอบรัน:</span>{' '}
                    <span style={{ color: showStandby ? 'var(--text-secondary)' : 'var(--accent-blue)' }}>
                      {showStandby ? 'Standby' : (currentJob?.name || '—')}
                    </span>
                  </h2>
                </div>

                {currentJob && (
                  <div className="nav-tabs" style={{ margin: 0 }}>
                    <button
                      className={`nav-tab ${activeTab === 'diagram' ? 'active' : ''}`}
                      onClick={() => setActiveTab('diagram')}
                    >
                      <Cpu size={16} /> Diagram
                    </button>
                    <button
                      className={`nav-tab ${activeTab === 'dashboard' ? 'active' : ''}`}
                      onClick={() => setActiveTab('dashboard')}
                    >
                      <LayoutDashboard size={16} /> Dashboard
                    </button>
                    <button
                      className={`nav-tab ${activeTab === 'combined' ? 'active' : ''}`}
                      onClick={() => setActiveTab('combined')}
                    >
                      <ChartIcon size={16} /> Graph
                    </button>
                    <button
                      className={`nav-tab ${activeTab === 'table' ? 'active' : ''}`}
                      onClick={() => setActiveTab('table')}
                    >
                      <TableIcon size={16} /> Table
                    </button>
                    <button
                      className={`nav-tab ${activeTab === 'ai' ? 'active' : ''}`}
                      onClick={() => setActiveTab('ai')}
                    >
                      <MessageSquare size={16} /> AI Assistant
                    </button>
                  </div>
                )}
              </div>

              <div className="header-controls" style={{ display: 'flex', gap: '8px', alignItems: 'center', flexWrap: 'wrap' }}>
                <button
                  className={`theme-toggle-btn`}
                  onClick={() => setTheme(prev => {
                    if (prev === 'dark') return 'light';
                    if (prev === 'light') return 'hmi';
                    return 'dark';
                  })}
                  title={`ธีม: ${theme === 'dark' ? 'Dark Mode' : theme === 'light' ? 'Light Mode' : 'Industrial HMI Mode'}`}
                  style={{ margin: 0 }}
                >
                  {theme === 'dark' && '🌙'}
                  {theme === 'light' && '☀️'}
                  {theme === 'hmi' && '⚙️'}
                </button>
                {currentJob && userRole === 'admin' && currentJob.status !== 'finished' && (
                  <button
                    className="status-control-btn btn-stop-machine"
                    onClick={() => handleToggleJobStatus(currentJob.id, currentJob.status || 'running')}
                    style={{ margin: 0, height: '38px', padding: '0 16px', display: 'inline-flex', alignItems: 'center', gap: '6px', fontWeight: 600, background: 'linear-gradient(135deg, #dc2626, #b91c1c)', borderColor: '#dc2626' }}
                  >
                    ✅ เสร็จสิ้นงาน
                  </button>
                )}
                {currentJob && currentJob.status === 'finished' && (
                  <span style={{
                    height: '38px', padding: '0 14px', display: 'inline-flex', alignItems: 'center', gap: '6px',
                    fontSize: '0.82rem', fontWeight: 700, color: '#86efac',
                    background: 'rgba(34,197,94,0.12)', border: '1px solid rgba(34,197,94,0.35)', borderRadius: '8px'
                  }}>
                    🏁 งานเสร็จสิ้นแล้ว
                  </span>
                )}
                {currentJob && currentJob.expiresAt && new Date() > new Date(currentJob.expiresAt) && (
                  <span style={{
                    height: '38px', padding: '0 14px', display: 'inline-flex', alignItems: 'center', gap: '6px',
                    fontSize: '0.82rem', fontWeight: 700, color: '#fca5a5',
                    background: 'rgba(239,68,68,0.12)', border: '1px solid rgba(239,68,68,0.35)', borderRadius: '8px'
                  }}>
                    ⏰ หมดอายุการแชร์แล้ว
                  </span>
                )}
                <button
                  className={`toggle-btn ${isReplay ? 'active' : ''}`}
                  onClick={() => {
                    if (isReplay) {
                      setIsReplay(false);
                      setIsReplayPlaying(false);
                      setReplayIndex(1);
                    } else {
                      // start replay from stored data
                      setReplayIndex(1);
                      setIsReplay(true);
                      setIsReplayPlaying(true);
                    }
                  }}
                  disabled={!currentJob || currentJobData.length === 0}
                  style={{ margin: 0 }}
                >
                  {isReplay ? 'Replay ON' : 'Replay'}
                </button>
                <button className="export-btn" onClick={exportToExcel} disabled={!currentJob} style={{ margin: 0 }}>
                  <Download size={18} style={{ marginRight: '8px' }} /> Export Excel
                </button>
                <button className="export-btn" onClick={exportToPDF} disabled={!currentJob} style={{ margin: 0 }}>
                  <Download size={18} style={{ marginRight: '8px' }} /> Export PDF
                </button>
                <button className="export-btn" onClick={exportToCSV} disabled={!currentJob} style={{ margin: 0 }}>
                  <Download size={18} style={{ marginRight: '8px' }} /> Export CSV
                </button>
                {currentJob && userRole === 'admin' && (
                  <button
                    className="export-btn"
                    onClick={() => {
                      setShareModalJobId(currentJob.id);
                      setShowCustomerShareModal(true);
                    }}
                    style={{ margin: 0, background: 'linear-gradient(135deg, var(--accent-green), #059669)', border: 'none', color: '#fff' }}
                    title="ตั้งค่าการแชร์และคัดลอกลิงก์ให้ลูกค้า"
                  >
                    <Users size={18} style={{ marginRight: '8px' }} /> แชร์ลูกค้า
                  </button>
                )}
                {/* Auto simulation removed */}
              </div>
            </header>

            {/* Global Parameter Visibility Selector Bar */}
            {currentJob && (
              <div className="glass-panel" style={{ padding: '0.75rem 1.25rem', marginBottom: '1rem', display: 'flex', alignItems: 'center', gap: '1rem', flexWrap: 'wrap' }}>
                <span style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '6px' }}>
                  ⚙️ แสดงค่าพารามิเตอร์:
                </span>
                <div style={{ display: 'flex', gap: '0.5rem', flexWrap: 'wrap' }}>
                  {[
                    { key: 'temp', label: 'TEMP', color: 'var(--accent-red)' },
                    { key: 'ph', label: 'pH', color: 'var(--accent-blue)' },
                    { key: 'do', label: 'DO', color: 'var(--accent-green)' },
                    { key: 'agit', label: 'AGIT', color: 'var(--accent-yellow)' },
                    { key: 'air', label: 'AIR FLOW', color: 'var(--accent-purple)' },
                    { key: 'level', label: 'VOLUME', color: 'var(--accent-green)' },
                    { key: 'air_out', label: 'AIR OUT', color: 'var(--accent-blue)' },
                    { key: 'heat', label: 'HEAT', color: 'var(--accent-yellow)' }
                  ].map(p => {
                    const active = visibleParameters[p.key];
                    return (
                      <button
                        key={p.key}
                        onClick={() => setVisibleParameters(prev => ({ ...prev, [p.key]: !prev[p.key] }))}
                        style={{
                          background: active
                            ? (theme === 'light' ? 'rgba(15, 23, 42, 0.06)' : 'rgba(255,255,255,0.08)')
                            : 'transparent',
                          borderColor: active ? p.color : 'var(--border-color)',
                          color: active ? 'var(--text-primary)' : 'var(--text-secondary)',
                          borderWidth: '1px',
                          borderStyle: 'solid',
                          borderRadius: '20px',
                          padding: '4px 12px',
                          fontSize: '0.8rem',
                          fontWeight: active ? 600 : 400,
                          cursor: 'pointer',
                          display: 'inline-flex',
                          alignItems: 'center',
                          gap: '6px',
                          transition: 'all 0.2s ease',
                          margin: 0
                        }}
                      >
                        <span style={{ width: '8px', height: '8px', borderRadius: '50%', backgroundColor: active ? p.color : 'transparent', border: active ? 'none' : '1px solid var(--text-secondary)' }}></span>
                        {p.label}
                      </button>
                    );
                  })}
                </div>
              </div>
            )}



            {!currentJob ? (
              <div className="glass-panel empty-state">
                <h2>Select or create a session to start.</h2>
              </div>
            ) : currentJobData.length === 0 ? (
              <div className="glass-panel empty-state">
                <FolderOpen size={48} opacity={0.5} />
                <h2>No data recorded yet</h2>
                <p>Use the manual entry form above to add your first record.</p>
              </div>
            ) : (
              activeTab === 'diagram' ? (
                <BSTRDiagram
                  dataPoint={lastDataPointForDisplay}
                  chartData={chartData}
                  isReplaying={isReplay}
                  isReplayingPlaying={isReplayPlaying}
                  jobStatus={currentJob?.status || 'running'}
                  onToggleStatus={() => handleToggleJobStatus(currentJob?.id, currentJob?.status || 'running')}
                  userRole={userRole}
                  isViewingHistory={isViewingHistory}
                  theme={theme}
                  aboutSystem={aboutSystem}
                  machineName={currentMachine?.name || ''}
                />
              ) : activeTab === 'dashboard' ? (
                <>
                  {/* Real-time Metrics Grid */}
                  <div className="metrics-grid">
                    {visibleParameters.temp && (
                      <div className="glass-panel metric-card">
                        <Thermometer color="var(--accent-red)" size={24} style={{ marginBottom: '10px' }} />
                        <div className="metric-title">Temperature</div>
                        <div className="metric-value-container" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                          <div className="metric-value" style={{ color: 'var(--accent-red)', fontSize: '2.2rem', lineHeight: 1.1 }}>
                            {typeof (lastDataPointForDisplay.temp_read !== undefined ? lastDataPointForDisplay.temp_read : lastDataPointForDisplay?.temp) === 'number' ? ((lastDataPointForDisplay.temp_read !== undefined ? lastDataPointForDisplay.temp_read : lastDataPointForDisplay?.temp) || 0).toFixed(2) : '-'}<span className="metric-unit">°C</span>
                          </div>
                          <div className="metric-value-sv" style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '4px' }}>
                            SV (Set): {typeof (lastDataPointForDisplay.temp_set !== undefined ? lastDataPointForDisplay.temp_set : lastDataPointForDisplay?.temp) === 'number' ? ((lastDataPointForDisplay.temp_set !== undefined ? lastDataPointForDisplay.temp_set : lastDataPointForDisplay?.temp) || 0).toFixed(2) : '-'}°C
                          </div>
                        </div>
                      </div>
                    )}
                    {visibleParameters.ph && (
                      <div className="glass-panel metric-card">
                        <Droplets color="var(--accent-blue)" size={24} style={{ marginBottom: '10px' }} />
                        <div className="metric-title">pH Level</div>
                        <div className="metric-value-container" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                          <div className="metric-value" style={{ color: 'var(--accent-blue)', fontSize: '2.2rem', lineHeight: 1.1 }}>
                            {typeof (lastDataPointForDisplay.ph_read !== undefined ? lastDataPointForDisplay.ph_read : lastDataPointForDisplay?.ph) === 'number' ? ((lastDataPointForDisplay.ph_read !== undefined ? lastDataPointForDisplay.ph_read : lastDataPointForDisplay?.ph) || 0).toFixed(2) : '-'}
                          </div>
                          <div className="metric-value-sv" style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '4px' }}>
                            SV (Set): {typeof (lastDataPointForDisplay.ph_set !== undefined ? lastDataPointForDisplay.ph_set : lastDataPointForDisplay?.ph) === 'number' ? ((lastDataPointForDisplay.ph_set !== undefined ? lastDataPointForDisplay.ph_set : lastDataPointForDisplay?.ph) || 0).toFixed(2) : '-'}
                          </div>
                        </div>
                      </div>
                    )}
                    {visibleParameters.do && (
                      <div className="glass-panel metric-card">
                        <Activity color="var(--accent-green)" size={24} style={{ marginBottom: '10px' }} />
                        <div className="metric-title">Dissolved Oxygen</div>
                        <div className="metric-value-container" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                          <div className="metric-value" style={{ color: 'var(--accent-green)', fontSize: '2.2rem', lineHeight: 1.1 }}>
                            {lastDataPointForDisplay.do_read !== undefined ? lastDataPointForDisplay.do_read : lastDataPointForDisplay.do}<span className="metric-unit">%</span>
                          </div>
                          <div className="metric-value-sv" style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '4px' }}>
                            SV (Set): {lastDataPointForDisplay.do_set !== undefined ? lastDataPointForDisplay.do_set : lastDataPointForDisplay.do}%
                          </div>
                        </div>
                      </div>
                    )}
                    {visibleParameters.agit && (
                      <div className="glass-panel metric-card">
                        <RotateCw color="var(--accent-yellow)" size={24} style={{ marginBottom: '10px' }} />
                        <div className="metric-title">Agitation</div>
                        <div className="metric-value-container" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                          <div className="metric-value" style={{ color: 'var(--accent-yellow)', fontSize: '2.2rem', lineHeight: 1.1 }}>
                            {lastDataPointForDisplay.agit_read !== undefined ? lastDataPointForDisplay.agit_read : lastDataPointForDisplay.agit}<span className="metric-unit">RPM</span>
                          </div>
                          <div className="metric-value-sv" style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '4px' }}>
                            SV (Set): {lastDataPointForDisplay.agit_set !== undefined ? lastDataPointForDisplay.agit_set : lastDataPointForDisplay.agit} RPM
                          </div>
                        </div>
                      </div>
                    )}
                    {visibleParameters.air && (
                      <div className="glass-panel metric-card">
                        <Wind color="var(--accent-purple)" size={24} style={{ marginBottom: '10px' }} />
                        <div className="metric-title">Air Flow</div>
                        <div className="metric-value-container" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                          <div className="metric-value" style={{ color: 'var(--accent-purple)', fontSize: '2.2rem', lineHeight: 1.1 }}>
                            {typeof (lastDataPointForDisplay.air_read !== undefined ? lastDataPointForDisplay.air_read : lastDataPointForDisplay?.air) === 'number' ? ((lastDataPointForDisplay.air_read !== undefined ? lastDataPointForDisplay.air_read : lastDataPointForDisplay?.air) || 0).toFixed(1) : '-'}<span className="metric-unit">L/M</span>
                          </div>
                          <div className="metric-value-sv" style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '4px' }}>
                            SV (Set): {typeof (lastDataPointForDisplay.air_set !== undefined ? lastDataPointForDisplay.air_set : lastDataPointForDisplay?.air) === 'number' ? ((lastDataPointForDisplay.air_set !== undefined ? lastDataPointForDisplay.air_set : lastDataPointForDisplay?.air) || 0).toFixed(1) : '-'} L/M
                          </div>
                        </div>
                      </div>
                    )}
                  </div>

                  {/* Charts Grid */}
                  <div className="charts-grid">
                    <div className="glass-panel chart-container">
                      <div className="chart-header">
                        <h2 className="chart-title" style={{ color: 'var(--accent-red)' }}>Temperature Trend</h2>
                      </div>
                      <ResponsiveContainer width="100%" height="85%">
                        <LineChart data={chartData}>
                          <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.1)" />
                          <XAxis dataKey="cultureHour" stroke="var(--text-secondary)" tick={{ fontSize: 12 }} />
                          <YAxis domain={['auto', 'auto']} stroke="var(--text-secondary)" tick={{ fontSize: 12 }} tickFormatter={(val) => typeof val === 'number' ? val.toFixed(2) : val} />
                          <Tooltip content={<CustomTooltip />} />
                          <Legend wrapperStyle={{ fontSize: '12px' }} />
                          <Line type="monotone" dataKey="temp_read" name="TEMP PV (Read)" stroke="var(--accent-red)" strokeWidth={3} dot={true} activeDot={{ r: 8 }} />
                          <Line type="monotone" dataKey="temp_set" name="TEMP SV (Set)" stroke="var(--accent-red)" strokeWidth={2} strokeDasharray="5 5" dot={false} />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>

                    <div className="glass-panel chart-container">
                      <div className="chart-header">
                        <h2 className="chart-title" style={{ color: 'var(--accent-blue)' }}>pH Trend</h2>
                      </div>
                      <ResponsiveContainer width="100%" height="85%">
                        <LineChart data={chartData}>
                          <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.1)" />
                          <XAxis dataKey="cultureHour" stroke="var(--text-secondary)" tick={{ fontSize: 12 }} />
                          <YAxis domain={['auto', 'auto']} stroke="var(--text-secondary)" tick={{ fontSize: 12 }} />
                          <Tooltip content={<CustomTooltip />} />
                          <Legend wrapperStyle={{ fontSize: '12px' }} />
                          <Line type="monotone" dataKey="ph_read" name="pH PV (Read)" stroke="var(--accent-blue)" strokeWidth={3} dot={true} activeDot={{ r: 8 }} />
                          <Line type="monotone" dataKey="ph_set" name="pH SV (Set)" stroke="var(--accent-blue)" strokeWidth={2} strokeDasharray="5 5" dot={false} />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>

                    <div className="glass-panel chart-container">
                      <div className="chart-header">
                        <h2 className="chart-title" style={{ color: 'var(--accent-green)' }}>DO & Agitation Trend</h2>
                      </div>
                      <ResponsiveContainer width="100%" height="85%">
                        <LineChart data={chartData}>
                          <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.1)" />
                          <XAxis dataKey="cultureHour" stroke="var(--text-secondary)" tick={{ fontSize: 12 }} />
                          <YAxis yAxisId="left" stroke="var(--text-secondary)" tick={{ fontSize: 12 }} />
                          <YAxis yAxisId="right" orientation="right" domain={['auto', 'auto']} stroke="var(--text-secondary)" tick={{ fontSize: 12 }} />
                          <Tooltip content={<CustomTooltip />} />
                          <Legend wrapperStyle={{ fontSize: '12px' }} />
                          <Line yAxisId="left" type="monotone" dataKey="do_read" name="DO PV (%)" stroke="var(--accent-green)" strokeWidth={3} dot={true} />
                          <Line yAxisId="left" type="monotone" dataKey="do_set" name="DO SV (%)" stroke="var(--accent-green)" strokeWidth={2} strokeDasharray="5 5" dot={false} />
                          <Line yAxisId="right" type="monotone" dataKey="agit_read" name="AGIT PV (RPM)" stroke="var(--accent-yellow)" strokeWidth={3} dot={true} />
                          <Line yAxisId="right" type="monotone" dataKey="agit_set" name="AGIT SV (RPM)" stroke="var(--accent-yellow)" strokeWidth={2} strokeDasharray="5 5" dot={false} />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                </>
              ) : activeTab === 'combined' ? (
                <div className="glass-panel" style={{ height: '650px', padding: '2rem' }}>
                  <div className="chart-header" style={{ marginBottom: '1.5rem', flexDirection: 'column', alignItems: 'flex-start', gap: '1rem' }}>
                    <h2 className="chart-title">Combined Bioprocess Overview</h2>

                    {/* Toggle Selector Buttons */}
                    <div style={{ display: 'flex', gap: '0.75rem', flexWrap: 'wrap', width: '100%', padding: '0.5rem', background: 'rgba(15, 23, 42, 0.3)', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                      <button
                        className="toggle-btn"
                        style={{
                          borderColor: 'var(--accent-red)',
                          color: visibleParameters.temp ? 'white' : 'var(--accent-red)',
                          background: visibleParameters.temp ? 'var(--accent-red)' : 'transparent',
                          fontWeight: 600,
                          fontSize: '0.85rem',
                          padding: '6px 12px'
                        }}
                        onClick={() => setVisibleParameters(prev => ({ ...prev, temp: !prev.temp }))}
                      >
                        TEMP ({visibleParameters.temp ? 'ON' : 'OFF'})
                      </button>
                      <button
                        className="toggle-btn"
                        style={{
                          borderColor: 'var(--accent-blue)',
                          color: visibleParameters.ph ? 'white' : 'var(--accent-blue)',
                          background: visibleParameters.ph ? 'var(--accent-blue)' : 'transparent',
                          fontWeight: 600,
                          fontSize: '0.85rem',
                          padding: '6px 12px'
                        }}
                        onClick={() => setVisibleParameters(prev => ({ ...prev, ph: !prev.ph }))}
                      >
                        pH ({visibleParameters.ph ? 'ON' : 'OFF'})
                      </button>
                      <button
                        className="toggle-btn"
                        style={{
                          borderColor: 'var(--accent-green)',
                          color: visibleParameters.do ? 'white' : 'var(--accent-green)',
                          background: visibleParameters.do ? 'var(--accent-green)' : 'transparent',
                          fontWeight: 600,
                          fontSize: '0.85rem',
                          padding: '6px 12px'
                        }}
                        onClick={() => setVisibleParameters(prev => ({ ...prev, do: !prev.do }))}
                      >
                        DO ({visibleParameters.do ? 'ON' : 'OFF'})
                      </button>
                      <button
                        className="toggle-btn"
                        style={{
                          borderColor: 'var(--accent-yellow)',
                          color: visibleParameters.agit ? 'white' : 'var(--accent-yellow)',
                          background: visibleParameters.agit ? 'var(--accent-yellow)' : 'transparent',
                          fontWeight: 600,
                          fontSize: '0.85rem',
                          padding: '6px 12px'
                        }}
                        onClick={() => setVisibleParameters(prev => ({ ...prev, agit: !prev.agit }))}
                      >
                        AGIT ({visibleParameters.agit ? 'ON' : 'OFF'})
                      </button>
                      <button
                        className="toggle-btn"
                        style={{
                          borderColor: 'var(--accent-purple)',
                          color: visibleParameters.air ? 'white' : 'var(--accent-purple)',
                          background: visibleParameters.air ? 'var(--accent-purple)' : 'transparent',
                          fontWeight: 600,
                          fontSize: '0.85rem',
                          padding: '6px 12px'
                        }}
                        onClick={() => setVisibleParameters(prev => ({ ...prev, air: !prev.air }))}
                      >
                        AIR ({visibleParameters.air ? 'ON' : 'OFF'})
                      </button>
                    </div>
                  </div>
                  <ResponsiveContainer width="100%" height="80%">
                    <LineChart data={chartData}>
                      <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.1)" />
                      <XAxis dataKey="cultureHour" stroke="var(--text-secondary)" tick={{ fontSize: 12 }} />
                      <YAxis yAxisId="left" stroke="var(--text-secondary)" tick={{ fontSize: 12 }} label={{ value: 'TEMP / pH / DO / AIR', angle: -90, position: 'insideLeft', fill: 'var(--text-secondary)' }} />
                      <YAxis yAxisId="right" orientation="right" domain={['auto', 'auto']} stroke="var(--text-secondary)" tick={{ fontSize: 12 }} label={{ value: 'AGIT (RPM)', angle: 90, position: 'insideRight', fill: 'var(--text-secondary)' }} />
                      <Tooltip content={<CustomTooltip />} />
                      <Legend wrapperStyle={{ fontSize: '14px', paddingTop: '10px' }} />
                      {visibleParameters.temp && <Line yAxisId="left" type="monotone" dataKey="temp_read" name="TEMP PV (°C)" stroke="var(--accent-red)" strokeWidth={3} dot={true} />}
                      {visibleParameters.temp && <Line yAxisId="left" type="monotone" dataKey="temp_set" name="TEMP SV (°C)" stroke="var(--accent-red)" strokeWidth={2} strokeDasharray="5 5" dot={false} />}
                      {visibleParameters.ph && <Line yAxisId="left" type="monotone" dataKey="ph_read" name="pH PV" stroke="var(--accent-blue)" strokeWidth={3} dot={true} />}
                      {visibleParameters.ph && <Line yAxisId="left" type="monotone" dataKey="ph_set" name="pH SV" stroke="var(--accent-blue)" strokeWidth={2} strokeDasharray="5 5" dot={false} />}
                      {visibleParameters.do && <Line yAxisId="left" type="monotone" dataKey="do_read" name="DO PV (%)" stroke="var(--accent-green)" strokeWidth={3} dot={true} />}
                      {visibleParameters.do && <Line yAxisId="left" type="monotone" dataKey="do_set" name="DO SV (%)" stroke="var(--accent-green)" strokeWidth={2} strokeDasharray="5 5" dot={false} />}
                      {visibleParameters.agit && <Line yAxisId="right" type="monotone" dataKey="agit_read" name="AGIT PV (RPM)" stroke="var(--accent-yellow)" strokeWidth={3} dot={true} />}
                      {visibleParameters.agit && <Line yAxisId="right" type="monotone" dataKey="agit_set" name="AGIT SV (RPM)" stroke="var(--accent-yellow)" strokeWidth={2} strokeDasharray="5 5" dot={false} />}
                      {visibleParameters.air && <Line yAxisId="left" type="monotone" dataKey="air_read" name="AIR PV (L/M)" stroke="var(--accent-purple)" strokeWidth={3} dot={true} />}
                      {visibleParameters.air && <Line yAxisId="left" type="monotone" dataKey="air_set" name="AIR SV (L/M)" stroke="var(--accent-purple)" strokeWidth={2} strokeDasharray="5 5" dot={false} />}
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              ) : activeTab === 'ai' ? (
                <div className="ai-assistant-container" style={{ display: 'flex', gap: '1.5rem', height: '650px', flexDirection: window.innerWidth < 1024 ? 'column' : 'row' }}>
                  {/* Left Column: AI Analyzer Report */}
                  <div className="glass-panel ai-report-panel" style={{ flex: 1.2, display: 'flex', flexDirection: 'column', padding: '1.5rem', overflow: 'hidden' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem', borderBottom: '1px solid var(--border-color)', paddingBottom: '0.75rem' }}>
                      <h2 style={{ fontSize: '1.2rem', fontWeight: 700, margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}>
                        🤖 AI Batch Insights & Report
                      </h2>
                      {aiReport && (
                        <div style={{ display: 'flex', gap: '8px' }}>
                          <button
                            className="btn btn-secondary btn-small"
                            onClick={() => {
                              navigator.clipboard.writeText(aiReport);
                              alert('คัดลอกรายงานลงคลิปบอร์ดแล้ว!');
                            }}
                          >
                            Copy
                          </button>
                          <button
                            className="btn btn-secondary btn-small"
                            onClick={() => {
                              const printWindow = window.open('', '_blank');
                              printWindow.document.write(`<html><head><title>AI Report - ${currentJob?.name}</title><style>body { font-family: sans-serif; padding: 2rem; line-height: 1.6; color: #333; } h1, h2, h3 { color: #1e3a8a; } pre { background: #f4f4f4; padding: 1rem; border-radius: 4px; }</style></head><body><h1>AI Analysis Report</h1><p><strong>Session:</strong> ${currentJob?.name} | <strong>Machine:</strong> ${currentMachine?.name}</p><hr/><div>${aiReport.replace(/\n/g, '<br/>')}</div></body></html>`);
                              printWindow.document.close();
                              printWindow.print();
                            }}
                          >
                            Print
                          </button>
                        </div>
                      )}
                    </div>

                    <div style={{ flex: 1, overflowY: 'auto', paddingRight: '0.5rem', display: 'flex', flexDirection: 'column', justifyContent: aiReport ? 'flex-start' : 'center', alignItems: aiReport ? 'stretch' : 'center' }}>
                      {isAiReportLoading ? (
                        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '1rem', margin: 'auto' }}>
                          <div className="ai-loading-pulse">🤖</div>
                          <p style={{ color: 'var(--text-secondary)', fontWeight: 500, animation: 'pulse 1.5s infinite' }}>กำลังวิเคราะห์พารามิเตอร์และจัดทำรายงาน...</p>
                        </div>
                      ) : aiReport ? (
                        <div className="ai-report-content" style={{ fontSize: '0.95rem', lineHeight: '1.6', whiteSpace: 'pre-wrap' }}>
                          {aiReport.split('\n').map((line, i) => {
                            if (line.startsWith('###')) {
                              return <h3 key={i} style={{ color: 'white', marginTop: '1.25rem', marginBottom: '0.5rem', fontWeight: 700, fontSize: '1.1rem', borderBottom: '1px solid rgba(255,255,255,0.05)', paddingBottom: '4px' }}>{line.replace('###', '').trim()}</h3>;
                            }
                            if (line.startsWith('####')) {
                              return <h4 key={i} style={{ color: 'var(--accent-blue)', marginTop: '1rem', marginBottom: '0.4rem', fontWeight: 600, fontSize: '1rem' }}>{line.replace('####', '').trim()}</h4>;
                            }
                            if (line.startsWith('*   **') || line.startsWith('-   **') || line.startsWith('* **') || line.startsWith('- **')) {
                              const match = line.match(/^[\*\-]\s+\*\*(.*?)\*\*(.*)$/);
                              if (match) {
                                return (
                                  <div key={i} style={{ display: 'flex', alignItems: 'flex-start', margin: '4px 0', paddingLeft: '8px' }}>
                                    <span style={{ color: 'var(--accent-purple)', marginRight: '8px', fontSize: '1.1rem', lineHeight: '1' }}>•</span>
                                    <span><strong>{match[1]}</strong>{match[2]}</span>
                                  </div>
                                );
                              }
                            }
                            if (line.startsWith('*') || line.startsWith('-')) {
                              return (
                                <div key={i} style={{ display: 'flex', alignItems: 'flex-start', margin: '4px 0', paddingLeft: '8px' }}>
                                  <span style={{ color: 'var(--accent-purple)', marginRight: '8px', fontSize: '1.1rem', lineHeight: '1' }}>•</span>
                                  <span>{line.substring(1).trim()}</span>
                                </div>
                              );
                            }
                            if (line.match(/^\d+\./)) {
                              return (
                                <div key={i} style={{ display: 'flex', alignItems: 'flex-start', margin: '6px 0', paddingLeft: '8px' }}>
                                  <span style={{ color: 'var(--accent-blue)', marginRight: '8px', fontWeight: 600 }}>{line.match(/^\d+\./)[0]}</span>
                                  <span>{line.replace(/^\d+\.\s*/, '')}</span>
                                </div>
                              );
                            }
                            return <p key={i} style={{ margin: '0.4rem 0', color: 'var(--text-primary)' }}>{line}</p>;
                          })}
                        </div>
                      ) : (
                        <div style={{ textAlign: 'center', padding: '2rem', maxWidth: '400px', margin: 'auto' }}>
                          <div style={{ fontSize: '3rem', marginBottom: '1rem', filter: 'drop-shadow(0 0 10px rgba(59, 130, 246, 0.4))' }}>🔬</div>
                          <h3 style={{ marginBottom: '0.5rem', color: 'white' }}>พร้อมรายงานผลวิเคราะห์รอบรัน</h3>
                          <p style={{ color: 'var(--text-secondary)', fontSize: '0.85rem', marginBottom: '1.5rem' }}>AI จะสแกนข้อมูลทั้งหมดในเซสชันนี้และวิเคราะห์หาแนวโน้ม ความผิดปกติ พร้อมคำแนะนำการผลิต</p>
                          <button
                            className="btn btn-blue"
                            onClick={async () => {
                              setIsAiReportLoading(true);
                              try {
                                const res = await fetch('/api/ai/analyze', {
                                  method: 'POST',
                                  headers: { 'Content-Type': 'application/json' },
                                  body: JSON.stringify({ jobId: currentJobId })
                                });
                                if (res.ok) {
                                  const data = await res.json();
                                  setAiReport(data.report);
                                  setIsAiReportMock(data.isMock);
                                } else {
                                  const errData = await res.json().catch(() => ({}));
                                  alert(`ไม่สามารถวิเคราะห์ข้อมูลได้: ${errData.error || 'เกิดข้อผิดพลาดไม่ทราบสาเหตุจากเซิร์ฟเวอร์'}`);
                                }
                              } catch (e) {
                                console.error(e);
                                alert('เกิดข้อผิดพลาดในการเชื่อมต่อเซิร์ฟเวอร์');
                              } finally {
                                setIsAiReportLoading(false);
                              }
                            }}
                          >
                            🚀 เริ่มการวิเคราะห์ข้อมูล
                          </button>
                        </div>
                      )}
                    </div>
                  </div>

                  {/* Right Column: AI Co-pilot Chatbot */}
                  <div className="glass-panel ai-chat-panel" style={{ flex: 1, display: 'flex', flexDirection: 'column', padding: '1.5rem', overflow: 'hidden' }}>
                    <h2 style={{ fontSize: '1.2rem', fontWeight: 700, margin: 0, paddingBottom: '0.75rem', borderBottom: '1px solid var(--border-color)', display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '1rem' }}>
                      💬 AI Bioprocess Chatbot
                    </h2>

                    {/* Chat Messages Area */}
                    <div style={{ flex: 1, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '0.75rem', paddingRight: '0.5rem', marginBottom: '1rem' }}>
                      {aiChatMessages.map((msg, i) => (
                        <div
                          key={i}
                          style={{
                            alignSelf: msg.role === 'user' ? 'flex-end' : 'flex-start',
                            maxWidth: '85%',
                            padding: '10px 14px',
                            borderRadius: msg.role === 'user' ? '12px 12px 2px 12px' : '12px 12px 12px 2px',
                            background: msg.role === 'user' ? 'linear-gradient(135deg, var(--accent-blue), #1e40af)' : 'rgba(255,255,255,0.06)',
                            border: msg.role === 'user' ? 'none' : '1px solid var(--border-color)',
                            color: 'white',
                            fontSize: '0.9rem',
                            lineHeight: '1.4',
                            boxShadow: '0 4px 6px -1px rgba(0,0,0,0.1)',
                            whiteSpace: 'pre-wrap'
                          }}
                        >
                          {msg.content}
                        </div>
                      ))}
                      {isAiChatLoading && (
                        <div style={{ alignSelf: 'flex-start', background: 'rgba(255,255,255,0.04)', padding: '10px 14px', borderRadius: '12px 12px 12px 2px', border: '1px solid var(--border-color)', color: 'var(--text-secondary)', fontSize: '0.85rem' }}>
                          <span className="typing-dots"><span>.</span><span>.</span><span>.</span></span> กำลังตอบกลับ
                        </div>
                      )}
                    </div>

                    {/* Prompt suggestions / Chips */}
                    <div style={{ display: 'flex', gap: '6px', overflowX: 'auto', paddingBottom: '8px', marginBottom: '8px' }}>
                      <button
                        className="btn btn-secondary btn-small"
                        style={{ whiteSpace: 'nowrap', borderRadius: '20px' }}
                        disabled={isAiChatLoading}
                        onClick={() => handleSendChipMessage('ตรวจสอบจุดผิดปกติในการรันนี้ให้หน่อย')}
                      >
                        🔍 เช็กจุดผิดปกติ
                      </button>
                      <button
                        className="btn btn-secondary btn-small"
                        style={{ whiteSpace: 'nowrap', borderRadius: '20px' }}
                        disabled={isAiChatLoading}
                        onClick={() => handleSendChipMessage('ค่า pH และ DO ในรันนี้เป็นอย่างไรบ้าง')}
                      >
                        🧪 ดูค่า pH / DO
                      </button>
                      <button
                        className="btn btn-secondary btn-small"
                        style={{ whiteSpace: 'nowrap', borderRadius: '20px' }}
                        disabled={isAiChatLoading}
                        onClick={() => handleSendChipMessage('ขอคำแนะนำในการคุมการกวนและการจ่ายอากาศ')}
                      >
                        ⚙️ แนะนำการกวน/จ่ายลม
                      </button>
                    </div>

                    {/* Chat Input Field */}
                    <form
                      onSubmit={async (e) => {
                        e.preventDefault();
                        if (!aiNewMessage.trim() || isAiChatLoading) return;
                        const userText = aiNewMessage;
                        setAiNewMessage('');
                        await sendChatMessage(userText);
                      }}
                      style={{ display: 'flex', gap: '8px' }}
                    >
                      <input
                        type="text"
                        value={aiNewMessage}
                        onChange={(e) => setAiNewMessage(e.target.value)}
                        placeholder="พิมพ์ถามผู้ช่วย AI ของคุณ..."
                        disabled={isAiChatLoading}
                        style={{ flex: 1, padding: '10px 14px', borderRadius: '8px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.9rem' }}
                      />
                      <button type="submit" className="btn btn-blue" disabled={isAiChatLoading || !aiNewMessage.trim()} style={{ margin: 0 }}>
                        ส่ง
                      </button>
                    </form>
                  </div>
                </div>
              ) : (
                <div className="glass-panel" style={{ padding: '2rem' }}>
                  <div className="table-header-controls">
                    <h2 className="chart-title">Recorded Data</h2>
                    {userRole === 'admin' && (
                      <button
                        className="delete-job-btn"
                        onClick={clearAllData}
                        style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', background: 'rgba(239, 68, 68, 0.1)', padding: '8px 16px', borderRadius: '8px' }}
                      >
                        <Trash2 size={16} /> Clear All Records
                      </button>
                    )}
                  </div>
                  <div className="data-table-container">
                    <table className="data-table">
                      <thead>
                        <tr>
                          <th rowSpan="2" style={{ verticalAlign: 'middle', textAlign: 'center', cursor: 'pointer' }} onClick={() => toggleSort('date')}>
                            Date {sortField === 'date' ? (sortAsc ? '▲' : '▼') : ''}
                          </th>
                          <th rowSpan="2" style={{ verticalAlign: 'middle', textAlign: 'center', cursor: 'pointer' }} onClick={() => toggleSort('time')}>
                            Time {sortField === 'time' ? (sortAsc ? '▲' : '▼') : ''}
                          </th>
                          <th rowSpan="2" style={{ verticalAlign: 'middle', textAlign: 'center', cursor: 'pointer' }} onClick={() => toggleSort('cultureHour')}>
                            ชั่วโมงที่ {sortField === 'cultureHour' ? (sortAsc ? '▲' : '▼') : ''}
                          </th>
                          {visibleParameters.temp && <th colSpan="2" style={{ textAlign: 'center', color: 'var(--accent-red)' }}>Temp (°C)</th>}
                          {visibleParameters.ph && <th colSpan="2" style={{ textAlign: 'center', color: 'var(--accent-blue)' }}>pH</th>}
                          {visibleParameters.do && <th colSpan="2" style={{ textAlign: 'center', color: 'var(--accent-green)' }}>DO (%)</th>}
                          {visibleParameters.agit && <th colSpan="2" style={{ textAlign: 'center', color: 'var(--accent-yellow)' }}>AGIT (RPM)</th>}
                          {visibleParameters.air && <th colSpan="2" style={{ textAlign: 'center', color: 'var(--accent-purple)' }}>AIR (L/M)</th>}
                          {visibleParameters.level && <th colSpan="2" style={{ textAlign: 'center', color: 'var(--accent-green)' }}>VOLUME (L)</th>}
                          {visibleParameters.air_out && <th colSpan="2" style={{ textAlign: 'center', color: 'var(--accent-purple)' }}>AIR OUT (PMa)</th>}
                          {visibleParameters.heat && <th colSpan="2" style={{ textAlign: 'center', color: 'var(--accent-yellow)' }}>HEAT (%)</th>}
                          <th rowSpan="2" style={{ verticalAlign: 'middle', textAlign: 'left', padding: '12px', minWidth: '150px' }}>Remarks / บันทึก</th>
                          {userRole === 'admin' && <th rowSpan="2" style={{ width: '80px', textAlign: 'center', verticalAlign: 'middle' }}>Action</th>}
                        </tr>
                        <tr>
                          {visibleParameters.temp && (
                            <>
                              <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>SV</th>
                              <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>PV</th>
                            </>
                          )}
                          {visibleParameters.ph && (
                            <>
                              <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>SV</th>
                              <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>PV</th>
                            </>
                          )}
                          {visibleParameters.do && (
                            <>
                              <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>SV</th>
                              <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>PV</th>
                            </>
                          )}
                          {visibleParameters.agit && (
                            <>
                              <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>SV</th>
                              <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>PV</th>
                            </>
                          )}
                          {visibleParameters.air && (
                            <>
                              <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>SV</th>
                              <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>PV</th>
                            </>
                          )}
                          {visibleParameters.level && (
                            <>
                              <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>SV</th>
                              <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>PV</th>
                            </>
                          )}
                          {visibleParameters.air_out && (
                            <>
                              <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>SV</th>
                              <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>PV</th>
                            </>
                          )}
                          {visibleParameters.heat && (
                            <>
                              <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>SV</th>
                              <th style={{ textAlign: 'center', fontSize: '0.75rem', padding: '6px' }}>PV</th>
                            </>
                          )}
                        </tr>
                      </thead>
                      <tbody>
                        {getSortedRows().map((row, index) => (
                          editingRowIndex === row.originalIndex ? (
                            <tr key={index} style={{ background: 'rgba(59, 130, 246, 0.05)' }}>
                              <td style={{ textAlign: 'center', padding: '6px' }}>
                                <input
                                  type="date"
                                  value={editingRowData.date}
                                  onChange={(e) => handleEditChange('date', e.target.value)}
                                  style={{ padding: '6px 4px', borderRadius: '4px', border: '1px solid var(--accent-blue)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.85rem', width: '115px' }}
                                  required
                                />
                              </td>
                              <td style={{ textAlign: 'center', padding: '6px' }}>
                                <input
                                  type="time"
                                  value={editingRowData.time}
                                  onChange={(e) => handleEditChange('time', e.target.value)}
                                  style={{ padding: '6px 4px', borderRadius: '4px', border: '1px solid var(--accent-blue)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.85rem', width: '85px' }}
                                  required
                                />
                              </td>
                              <td style={{ textAlign: 'center', padding: '6px', color: 'var(--text-secondary)' }}>
                                {getEditingRowCultureHour()} ชม.
                              </td>
                              {visibleParameters.temp && (
                                <>
                                  <td style={{ textAlign: 'center', padding: '6px' }}>
                                    <input
                                      type="number"
                                      step="0.01"
                                      value={editingRowData.temp_set}
                                      onChange={(e) => handleEditChange('temp_set', parseFloat(e.target.value) || 0)}
                                      style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.85rem', width: '45px', textAlign: 'center' }}
                                    />
                                  </td>
                                  <td style={{ textAlign: 'center', padding: '6px' }}>
                                    <input
                                      type="number"
                                      step="0.01"
                                      value={editingRowData.temp_read}
                                      onChange={(e) => handleEditChange('temp_read', parseFloat(e.target.value) || 0)}
                                      style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--accent-red)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.85rem', width: '45px', textAlign: 'center', fontWeight: 600 }}
                                    />
                                  </td>
                                </>
                              )}
                              {visibleParameters.ph && (
                                <>
                                  <td style={{ textAlign: 'center', padding: '6px' }}>
                                    <input
                                      type="number"
                                      step="0.01"
                                      value={editingRowData.ph_set}
                                      onChange={(e) => handleEditChange('ph_set', parseFloat(e.target.value) || 0)}
                                      style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.85rem', width: '50px', textAlign: 'center' }}
                                    />
                                  </td>
                                  <td style={{ textAlign: 'center', padding: '6px' }}>
                                    <input
                                      type="number"
                                      step="0.01"
                                      value={editingRowData.ph_read}
                                      onChange={(e) => handleEditChange('ph_read', parseFloat(e.target.value) || 0)}
                                      style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--accent-blue)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.85rem', width: '50px', textAlign: 'center', fontWeight: 600 }}
                                    />
                                  </td>
                                </>
                              )}
                              {visibleParameters.do && (
                                <>
                                  <td style={{ textAlign: 'center', padding: '6px' }}>
                                    <input
                                      type="number"
                                      step="1"
                                      value={editingRowData.do_set}
                                      onChange={(e) => handleEditChange('do_set', parseFloat(e.target.value) || 0)}
                                      style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.85rem', width: '40px', textAlign: 'center' }}
                                    />
                                  </td>
                                  <td style={{ textAlign: 'center', padding: '6px' }}>
                                    <input
                                      type="number"
                                      step="1"
                                      value={editingRowData.do_read}
                                      onChange={(e) => handleEditChange('do_read', parseFloat(e.target.value) || 0)}
                                      style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--accent-green)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.85rem', width: '40px', textAlign: 'center', fontWeight: 600 }}
                                    />
                                  </td>
                                </>
                              )}
                              {visibleParameters.agit && (
                                <>
                                  <td style={{ textAlign: 'center', padding: '6px' }}>
                                    <input
                                      type="number"
                                      step="1"
                                      value={editingRowData.agit_set}
                                      onChange={(e) => handleEditChange('agit_set', parseFloat(e.target.value) || 0)}
                                      style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.85rem', width: '50px', textAlign: 'center' }}
                                    />
                                  </td>
                                  <td style={{ textAlign: 'center', padding: '6px' }}>
                                    <input
                                      type="number"
                                      step="1"
                                      value={editingRowData.agit_read}
                                      onChange={(e) => handleEditChange('agit_read', parseFloat(e.target.value) || 0)}
                                      style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--accent-yellow)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.85rem', width: '50px', textAlign: 'center', fontWeight: 600 }}
                                    />
                                  </td>
                                </>
                              )}
                              {visibleParameters.air && (
                                <>
                                  <td style={{ textAlign: 'center', padding: '6px' }}>
                                    <input
                                      type="number"
                                      step="0.1"
                                      value={editingRowData.air_set}
                                      onChange={(e) => handleEditChange('air_set', parseFloat(e.target.value) || 0)}
                                      style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.85rem', width: '45px', textAlign: 'center' }}
                                    />
                                  </td>
                                  <td style={{ textAlign: 'center', padding: '6px' }}>
                                    <input
                                      type="number"
                                      step="0.1"
                                      value={editingRowData.air_read}
                                      onChange={(e) => handleEditChange('air_read', parseFloat(e.target.value) || 0)}
                                      style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--accent-purple)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.85rem', width: '45px', textAlign: 'center', fontWeight: 600 }}
                                    />
                                  </td>
                                </>
                              )}
                              {visibleParameters.level && (
                                <>
                                  <td style={{ textAlign: 'center', padding: '6px' }}>
                                    <input
                                      type="number"
                                      step="0.1"
                                      value={editingRowData.level_set !== undefined && editingRowData.level_set !== null ? parseFloat(((editingRowData.level_set / 100) * maxVol).toFixed(1)) : ''}
                                      onChange={(e) => handleEditChange('level_set', (parseFloat(e.target.value) || 0) / maxVol * 100)}
                                      style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.85rem', width: '55px', textAlign: 'center' }}
                                    />
                                  </td>
                                  <td style={{ textAlign: 'center', padding: '6px' }}>
                                    <input
                                      type="number"
                                      step="0.1"
                                      value={editingRowData.level_read !== undefined && editingRowData.level_read !== null ? parseFloat(((editingRowData.level_read / 100) * maxVol).toFixed(1)) : ''}
                                      onChange={(e) => handleEditChange('level_read', (parseFloat(e.target.value) || 0) / maxVol * 100)}
                                      style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--accent-green)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.85rem', width: '55px', textAlign: 'center', fontWeight: 600 }}
                                    />
                                  </td>
                                </>
                              )}
                              {visibleParameters.air_out && (
                                <>
                                  <td style={{ textAlign: 'center', padding: '6px' }}>
                                    <input
                                      type="number"
                                      step="0.1"
                                      value={editingRowData.air_out_set}
                                      onChange={(e) => handleEditChange('air_out_set', parseFloat(e.target.value) || 0)}
                                      style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.85rem', width: '45px', textAlign: 'center' }}
                                    />
                                  </td>
                                  <td style={{ textAlign: 'center', padding: '6px' }}>
                                    <input
                                      type="number"
                                      step="0.1"
                                      value={editingRowData.air_out_read}
                                      onChange={(e) => handleEditChange('air_out_read', parseFloat(e.target.value) || 0)}
                                      style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--accent-purple)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.85rem', width: '45px', textAlign: 'center', fontWeight: 600 }}
                                    />
                                  </td>
                                </>
                              )}
                              {visibleParameters.heat && (
                                <>
                                  <td style={{ textAlign: 'center', padding: '6px' }}>
                                    <input
                                      type="number"
                                      step="1"
                                      value={editingRowData.heat_set}
                                      onChange={(e) => handleEditChange('heat_set', parseFloat(e.target.value) || 0)}
                                      style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.85rem', width: '45px', textAlign: 'center' }}
                                    />
                                  </td>
                                  <td style={{ textAlign: 'center', padding: '6px' }}>
                                    <input
                                      type="number"
                                      step="1"
                                      value={editingRowData.heat_read}
                                      onChange={(e) => handleEditChange('heat_read', parseFloat(e.target.value) || 0)}
                                      style={{ padding: '6px 2px', borderRadius: '4px', border: '1px solid var(--accent-yellow)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.85rem', width: '45px', textAlign: 'center', fontWeight: 600 }}
                                    />
                                  </td>
                                </>
                              )}
                              <td style={{ textAlign: 'left', padding: '6px' }}>
                                <input
                                  type="text"
                                  value={editingRowData.remark}
                                  onChange={(e) => handleEditChange('remark', e.target.value)}
                                  style={{ padding: '6px 8px', borderRadius: '4px', border: '1px solid var(--border-color)', background: 'var(--bg-color)', color: 'var(--text-primary)', fontSize: '0.85rem', width: '100%', minWidth: '100px' }}
                                />
                              </td>
                              {userRole === 'admin' && (
                                <td style={{ textAlign: 'center', padding: '6px' }}>
                                  <div style={{ display: 'flex', gap: '8px', justifyContent: 'center', alignItems: 'center' }}>
                                    <button
                                      className="save-row-btn"
                                      onClick={saveEditRow}
                                      title="Save changes"
                                      style={{ background: 'transparent', border: 'none', color: 'var(--accent-green)', cursor: 'pointer', display: 'flex', alignItems: 'center' }}
                                    >
                                      <Check size={16} />
                                    </button>
                                    <button
                                      className="cancel-row-btn"
                                      onClick={cancelEditRow}
                                      title="Cancel editing"
                                      style={{ background: 'transparent', border: 'none', color: 'var(--accent-red)', cursor: 'pointer', display: 'flex', alignItems: 'center' }}
                                    >
                                      <X size={16} />
                                    </button>
                                  </div>
                                </td>
                              )}
                            </tr>
                          ) : (
                            <tr key={index}>
                              <td style={{ textAlign: 'center' }}>{row.timestamp ? toYYYYMMDD(row.timestamp) : (row.date || '')}</td>
                              <td style={{ textAlign: 'center' }}>{row.timestamp ? toHHMM(row.timestamp) : (row.time || '')}</td>
                              <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{row.cultureHour !== undefined ? `${row.cultureHour} ชม.` : '-'}</td>
                              {visibleParameters.temp && (
                                <>
                                  <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{typeof row.temp_set === 'number' ? row.temp_set.toFixed(1) : '-'}</td>
                                  <td style={{ textAlign: 'center', color: 'var(--accent-red)', fontWeight: 600 }}>{typeof row.temp_read === 'number' ? row.temp_read.toFixed(1) : '-'}</td>
                                </>
                              )}
                              {visibleParameters.ph && (
                                <>
                                  <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{typeof row.ph_set === 'number' ? row.ph_set.toFixed(2) : '-'}</td>
                                  <td style={{ textAlign: 'center', color: 'var(--accent-blue)', fontWeight: 600 }}>{typeof row.ph_read === 'number' ? row.ph_read.toFixed(2) : '-'}</td>
                                </>
                              )}
                              {visibleParameters.do && (
                                <>
                                  <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{row.do_set || '-'}</td>
                                  <td style={{ textAlign: 'center', color: 'var(--accent-green)', fontWeight: 600 }}>{row.do_read || '-'}</td>
                                </>
                              )}
                              {visibleParameters.agit && (
                                <>
                                  <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{row.agit_set || '-'}</td>
                                  <td style={{ textAlign: 'center', color: 'var(--accent-yellow)', fontWeight: 600 }}>{row.agit_read || '-'}</td>
                                </>
                              )}
                              {visibleParameters.air && (
                                <>
                                  <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{typeof row.air_set === 'number' ? row.air_set.toFixed(1) : '-'}</td>
                                  <td style={{ textAlign: 'center', color: 'var(--accent-purple)', fontWeight: 600 }}>{typeof row.air_read === 'number' ? row.air_read.toFixed(1) : '-'}</td>
                                </>
                              )}
                              {visibleParameters.level && (
                                <>
                                  <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{typeof row.level_set === 'number' ? row.level_set.toFixed(1) : '-'}</td>
                                  <td style={{ textAlign: 'center', color: 'var(--accent-green)', fontWeight: 600 }}>{typeof row.level_read === 'number' ? row.level_read.toFixed(1) : '-'}</td>
                                </>
                              )}
                              {visibleParameters.air_out && (
                                <>
                                  <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{typeof row.air_out_set === 'number' ? row.air_out_set.toFixed(1) : '-'}</td>
                                  <td style={{ textAlign: 'center', color: 'var(--accent-purple)', fontWeight: 600 }}>{typeof row.air_out_read === 'number' ? row.air_out_read.toFixed(1) : '-'}</td>
                                </>
                              )}
                              {visibleParameters.heat && (
                                <>
                                  <td style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>{typeof row.heat_set === 'number' ? row.heat_set.toFixed(1) : '-'}</td>
                                  <td style={{ textAlign: 'center', color: 'var(--accent-yellow)', fontWeight: 600 }}>{typeof row.heat_read === 'number' ? row.heat_read.toFixed(1) : '-'}</td>
                                </>
                              )}
                              <td className="remarks-cell" style={{ textAlign: 'left', padding: '12px', fontSize: '0.85rem', color: 'var(--text-secondary)' }}>{row.remark || '-'}</td>
                              {userRole === 'admin' && (
                                <td style={{ textAlign: 'center' }}>
                                  <div style={{ display: 'flex', gap: '8px', justifyContent: 'center', alignItems: 'center' }}>
                                    <button
                                      className="edit-row-btn"
                                      onClick={() => startEditRow(row.originalIndex, row)}
                                      title="Edit this record"
                                      style={{ background: 'transparent', border: 'none', color: 'var(--accent-blue)', cursor: 'pointer', display: 'flex', alignItems: 'center', opacity: 0.7 }}
                                      onMouseEnter={e => e.currentTarget.style.opacity = 1}
                                      onMouseLeave={e => e.currentTarget.style.opacity = 0.7}
                                    >
                                      <Edit3 size={16} />
                                    </button>
                                    <button
                                      className="delete-row-btn"
                                      onClick={() => deleteDataPoint(row.originalIndex)}
                                      title="Delete this record"
                                      style={{ margin: '0' }}
                                    >
                                      <Trash2 size={16} />
                                    </button>
                                  </div>
                                </td>
                              )}
                            </tr>
                          )
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )
            )}
          </>
        )}

        {/* Replay Control Panel */}
        {isReplay && sortedFullData.length > 0 && (
          <div className="replay-control-panel">
            <div className="replay-controls-row">
              {/* Buttons Group */}
              <div className="replay-btn-group">
                <button
                  className="replay-control-btn"
                  onClick={() => {
                    setReplayIndex(1);
                    setIsReplayPlaying(true);
                  }}
                  title="Restart Replay"
                >
                  <RotateCcw size={18} />
                </button>
                <button
                  className="replay-control-btn active"
                  style={{ background: isReplayPlaying ? 'var(--accent-blue)' : 'rgba(255, 255, 255, 0.05)' }}
                  onClick={() => setIsReplayPlaying(!isReplayPlaying)}
                  title={isReplayPlaying ? "Pause" : "Play"}
                >
                  {isReplayPlaying ? <Pause size={18} /> : <Play size={18} />}
                </button>
              </div>


              {/* Scrubber Slider */}
              <div className="replay-slider-container">
                <input
                  type="range"
                  min="1"
                  max={sortedFullData.length}
                  value={replayIndex}
                  onChange={(e) => {
                    setReplayIndex(Number(e.target.value));
                    setIsReplayPlaying(false);
                  }}
                  className="replay-slider"
                />
                <span className="replay-status-info">
                  จุดที่ {replayIndex} / {sortedFullData.length} (ชั่วโมงเลี้ยงเชื้อ: {chartData[replayIndex - 1]?.cultureHour?.toFixed(1) || '0.0'} ชม.)
                </span>
              </div>


              {/* Speed & Close Group */}
              <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', flexWrap: 'wrap' }}>
                <div className="replay-speed-selector">
                  <button
                    className={`replay-speed-btn ${replaySpeed === 2500 ? 'active' : ''}`}
                    onClick={() => setReplaySpeed(2500)}
                    title="Slow Speed"
                  >
                    0.5x
                  </button>
                  <button
                    className={`replay-speed-btn ${replaySpeed === 1500 ? 'active' : ''}`}
                    onClick={() => setReplaySpeed(1500)}
                    title="Normal Speed"
                  >
                    1x
                  </button>
                  <button
                    className={`replay-speed-btn ${replaySpeed === 600 ? 'active' : ''}`}
                    onClick={() => setReplaySpeed(600)}
                    title="Fast Speed"
                  >
                    2.5x
                  </button>
                  <button
                    className={`replay-speed-btn ${replaySpeed === 250 ? 'active' : ''}`}
                    onClick={() => setReplaySpeed(250)}
                    title="Hyper Speed"
                  >
                    5x
                  </button>
                </div>

                <button
                  className="replay-close-btn"
                  onClick={() => {
                    setIsReplay(false);
                    setIsReplayPlaying(false);
                    setReplayIndex(1);
                  }}
                >
                  Exit Replay
                </button>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Glassmorphic Customer Share Modal */}
      {showCustomerShareModal && shareModalJobId && (
        <div className="modal-backdrop" onClick={() => { setShowCustomerShareModal(false); setShareModalJobId(null); }}>
          <div className="modal-container" onClick={(e) => e.stopPropagation()}>

            {/* Modal Header */}
            <div className="modal-header">
              <h3>
                <Users size={22} color="var(--accent-green)" />
                ตั้งค่าลิงก์ลูกค้า & ความปลอดภัย (Customer Share)
              </h3>
              <button
                className="modal-close-btn"
                onClick={() => {
                  setShowCustomerShareModal(false);
                  setShareModalJobId(null);
                }}
                title="ปิด"
              >
                <X size={20} />
              </button>
            </div>

            {/* Modal Body */}
            {(() => {
              const job = jobs.find(j => j.id === shareModalJobId);
              if (!job) return <div style={{ color: 'var(--accent-red)' }}>ไม่พบข้อมูลรอบบันทึก (Session Not Found)</div>;

              const machine = machines.find(m => m.id === job.machineId);

              // Expiry calculations
              let expiryStatusText = "ไม่มีวันหมดอายุ (Unlimited)";
              let isExpired = false;
              if (job.expiresAt) {
                const expDate = new Date(job.expiresAt);
                isExpired = expDate < new Date();
                expiryStatusText = isExpired
                  ? `หมดอายุแล้วเมื่อ ${formatDateTime(expDate)}`
                  : `หมดอายุวันที่ ${formatDateTime(expDate)}`;
              }

              // Look up customer assigned to this machine
              const assignedCustomer = customers.find(c => c.machineId === job.machineId);
              const loginUrl = `${window.location.origin}/?job=${job.id}`;

              // Email invitation body template
              const customerName = assignedCustomer ? assignedCustomer.companyName : "[ชื่อลูกค้า]";
              const expiryInfoText = job.expiresAt ? formatDateTime(job.expiresAt) : "ไม่มีวันหมดอายุ";
              const invitationText = `เรียนคุณ ${customerName},\n\nทางแล็บขอส่งลิงก์สำหรับเข้าดูข้อมูลไบโอโพรเซสรอบรัน "${job.name}" (${machine?.name || 'เครื่องมือ'}) แบบเรียลไทม์\n\nลิงก์เข้าสู่ระบบ: ${loginUrl}\nวันหมดอายุ: ${expiryInfoText}\n\nขอบคุณค่ะ/ครับ\nDBMS System`;

              // Auto-populate editable text when first opening / when job changes
              const currentEditableText = shareInvitationText || invitationText;
              if (!shareInvitationText) {
                setTimeout(() => setShareInvitationText(invitationText), 0);
              }

              return (
                <div className="modal-body">

                  {/* Info Row */}
                  <div className="modal-info-box">
                    <div>
                      <span style={{ color: 'var(--text-secondary)' }}>รอบรัน (Session): </span>
                      <strong style={{ color: 'var(--accent-blue)' }}>{job.name}</strong>
                    </div>
                    <div>
                      <span style={{ color: 'var(--text-secondary)' }}>เครื่องมือ: </span>
                      <strong>{machine?.name || '-'}</strong>
                    </div>
                  </div>

                  {/* Direct Link */}
                  <div>
                    <label className="modal-label">🔗 ลิงก์เข้าใช้งานตรงสำหรับลูกค้า (Direct Share Link)</label>
                    <div className="modal-input-row">
                      <input
                        type="text"
                        readOnly
                        value={loginUrl}
                        className="modal-input"
                        onClick={(e) => e.target.select()}
                      />
                      <button
                        className="btn btn-blue"
                        style={{ margin: 0 }}
                        onClick={async () => {
                          await navigator.clipboard.writeText(loginUrl);
                          alert("คัดลอกลิงก์เรียบร้อยแล้ว!");
                        }}
                      >
                        คัดลอกลิงก์
                      </button>
                    </div>
                  </div>

                  {/* Access Expiration */}
                  <div className="modal-section">
                    <label className="modal-label">⏱️ กำหนดเวลาหมดอายุ (Access Expiration)</label>
                    <div className="modal-preset-row">
                      {[
                        { label: '24 ชม. (1 วัน)', value: 24 },
                        { label: '72 ชม. (3 วัน)', value: 72 },
                        { label: '168 ชม. (7 วัน)', value: 168 },
                        { label: 'ไม่จำกัด', value: 0 }
                      ].map((preset) => (
                        <button
                          key={preset.value}
                          type="button"
                          className="modal-preset-btn"
                          onClick={async () => {
                            let expiresAt = null;
                            if (preset.value > 0) {
                              expiresAt = new Date(Date.now() + preset.value * 60 * 60 * 1000).toISOString();
                            }
                            try {
                              const res = await fetch(`/api/jobs/${job.id}/expiry`, {
                                method: 'PUT',
                                headers: { 'Content-Type': 'application/json' },
                                body: JSON.stringify({ expiresAt })
                              });
                              if (res.ok) {
                                const data = await res.json();
                                applyDBUpdate(data);
                              }
                            } catch (err) {
                              console.error(err);
                            }
                          }}
                        >
                          {preset.label}
                        </button>
                      ))}

                      {/* Manual Input */}
                      <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginLeft: 'auto' }}>
                        <input
                          type="number"
                          placeholder="ชม. เช่น 48"
                          className="modal-input"
                          style={{ width: '90px', padding: '6px 8px', textAlign: 'center', fontSize: '0.85rem' }}
                          onKeyDown={async (e) => {
                            if (e.key === 'Enter') {
                              const hrs = parseFloat(e.target.value);
                              if (!isNaN(hrs) && hrs >= 0) {
                                let expiresAt = null;
                                if (hrs > 0) {
                                  expiresAt = new Date(Date.now() + hrs * 60 * 60 * 1000).toISOString();
                                }
                                try {
                                  const res = await fetch(`/api/jobs/${job.id}/expiry`, {
                                    method: 'PUT',
                                    headers: { 'Content-Type': 'application/json' },
                                    body: JSON.stringify({ expiresAt })
                                  });
                                  if (res.ok) {
                                    const data = await res.json();
                                    applyDBUpdate(data);
                                    e.target.value = '';
                                    alert(`ตั้งค่าหมดอายุในอีก ${hrs} ชั่วโมงเรียบร้อย!`);
                                  }
                                } catch (err) {
                                  console.error(err);
                                }
                              }
                            }
                          }}
                        />
                        <span style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>ชม. (กด Enter)</span>
                      </div>
                    </div>

                    <div style={{ fontSize: '0.85rem', marginTop: '10px' }}>
                      สถานะอายุลิงก์: <span style={{ color: isExpired ? 'var(--accent-red)' : 'var(--accent-green)', fontWeight: 700 }}>{expiryStatusText}</span>
                    </div>
                  </div>

                  {/* Allow Historical Viewing Toggle */}
                  <div className="modal-section" style={{ borderTop: '1px solid rgba(255, 255, 255, 0.05)', paddingTop: '12px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                      <div>
                        <label className="modal-label" style={{ marginBottom: '2px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                          📂 อนุญาตให้เข้าดูย้อนหลัง (Historical View)
                        </label>
                        <div style={{ fontSize: '0.78rem', color: 'var(--text-secondary)' }}>
                          ให้ลูกค้าเข้าดูบอร์ดข้อมูลได้แบบอ่านอย่างเดียว หลังจากการรันเสร็จสิ้นหรือหมดอายุแชร์
                        </div>
                      </div>
                      <label className="switch-container" style={{ display: 'inline-block', position: 'relative', width: '48px', height: '24px', cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={!!job.allowHistoryView}
                          onChange={async (e) => {
                            const val = e.target.checked;
                            try {
                              const res = await fetch(`/api/jobs/${job.id}/history-permission`, {
                                method: 'PUT',
                                headers: { 'Content-Type': 'application/json' },
                                body: JSON.stringify({ allowHistoryView: val })
                              });
                              if (res.ok) {
                                const data = await res.json();
                                applyDBUpdate(data);
                              }
                            } catch (err) {
                              console.error(err);
                            }
                          }}
                          style={{ opacity: 0, width: 0, height: 0 }}
                        />
                        <span className="switch-slider" style={{
                          position: 'absolute', cursor: 'pointer', top: 0, left: 0, right: 0, bottom: 0,
                          backgroundColor: job.allowHistoryView ? 'var(--accent-green)' : '#475569',
                          transition: '0.3s', borderRadius: '24px',
                          boxShadow: job.allowHistoryView ? '0 0 10px rgba(16, 185, 129, 0.4)' : 'none'
                        }}>
                          <span style={{
                            position: 'absolute', content: '""', height: '18px', width: '18px', left: job.allowHistoryView ? '26px' : '4px', bottom: '3px',
                            backgroundColor: 'white', transition: '0.3s', borderRadius: '50%'
                          }} />
                        </span>
                      </label>
                    </div>
                  </div>

                  {/* Customer Assignment */}
                  <div className="modal-section">
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.5rem' }}>
                      <label className="modal-label" style={{ margin: 0 }}>👥 ลูกค้าผู้ดูแลเครื่องมือ (Assigned Customer)</label>
                      {assignedCustomer && (
                        <button
                          type="button"
                          className="btn btn-primary btn-small"
                          style={{ margin: 0 }}
                          onClick={async () => {
                            await navigator.clipboard.writeText(currentEditableText);
                            alert("คัดลอกข้อความคำเชิญสำหรับส่งให้ลูกค้าเรียบร้อยแล้ว!");
                          }}
                        >
                          📋 คัดลอกข้อความ
                        </button>
                      )}
                    </div>

                    {assignedCustomer ? (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>
                        <div style={{ fontSize: '0.9rem', color: 'var(--text-primary)', background: 'rgba(255, 255, 255, 0.02)', padding: '10px 14px', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                          <div>บริษัท/ชื่อ: <strong>{assignedCustomer.companyName}</strong></div>
                          <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', marginTop: '2px' }}>อีเมล: {assignedCustomer.email || 'ไม่ได้ระบุ'}</div>
                        </div>

                        {/* Editable invitation textarea */}
                        <div style={{ position: 'relative' }}>
                          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '6px' }}>
                            <span style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>✏️ แก้ไขข้อความได้โดยตรง</span>
                            <button
                              type="button"
                              onClick={() => setShareInvitationText(invitationText)}
                              style={{
                                fontSize: '0.75rem', padding: '3px 8px', borderRadius: '6px',
                                border: '1px solid rgba(255,255,255,0.15)',
                                background: 'rgba(255,255,255,0.06)', color: 'var(--text-secondary)',
                                cursor: 'pointer'
                              }}
                            >
                              🔄 รีเซ็ตข้อความ
                            </button>
                          </div>
                          <textarea
                            value={currentEditableText}
                            onChange={(e) => setShareInvitationText(e.target.value)}
                            rows={9}
                            style={{
                              width: '100%',
                              padding: '12px 14px',
                              borderRadius: '10px',
                              border: '1px solid rgba(0, 240, 255, 0.2)',
                              background: 'rgba(0, 15, 30, 0.5)',
                              color: 'var(--text-primary)',
                              fontSize: '0.85rem',
                              lineHeight: 1.65,
                              resize: 'vertical',
                              fontFamily: 'inherit',
                              outline: 'none',
                              transition: 'border-color 0.2s, box-shadow 0.2s',
                              boxSizing: 'border-box'
                            }}
                            onFocus={e => {
                              e.target.style.borderColor = 'rgba(0,240,255,0.5)';
                              e.target.style.boxShadow = '0 0 0 3px rgba(0,240,255,0.08)';
                            }}
                            onBlur={e => {
                              e.target.style.borderColor = 'rgba(0,240,255,0.2)';
                              e.target.style.boxShadow = 'none';
                            }}
                          />
                        </div>
                      </div>

                    ) : (
                      <div style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', background: 'rgba(239, 68, 68, 0.05)', padding: '12px', borderRadius: '8px', border: '1px solid rgba(239, 68, 68, 0.15)', textAlign: 'center', lineHeight: 1.5 }}>
                        ⚠️ ยังไม่ได้ระบุข้อมูลลูกค้ารับผิดชอบเครื่องมือ "{machine?.name || ''}" ในฐานข้อมูล<br />
                        <span
                          style={{ color: 'var(--accent-blue)', textDecoration: 'underline', cursor: 'pointer', fontWeight: 600, display: 'inline-block', marginTop: '6px' }}
                          onClick={() => {
                            setShowCustomerShareModal(false);
                            setShareModalJobId(null);
                            setCurrentAppView('customers');
                            setCustomerFormData(prev => ({ ...prev, machineId: job.machineId }));
                          }}
                        >
                          คลิกเพื่อไปหน้าฐานข้อมูลลูกค้าเพื่อเพิ่มข้อมูลและจับคู่เครื่องมือ
                        </span>
                      </div>
                    )}
                  </div>

                </div>
              );
            })()}

          </div>
        </div>
      )}

      {/* Glassmorphic Feedback Modal — แบบประเมินความพึงพอใจ */}
      {showFeedbackModal && (() => {
        const closeFeedbackModal = () => {
          setShowFeedbackModal(false);
          if (window._pendingLogout) {
            window._pendingLogout = false;
            setUserRole(null);
            setActiveCustomerJobId(null);
            setIsMobileMenuOpen(false);
          }
        };

        // Group questions by section
        const sections = [];
        const sectionMap = {};
        SURVEY_QUESTIONS.forEach(q => {
          if (!sectionMap[q.section]) {
            sectionMap[q.section] = [];
            sections.push({ name: q.section, questions: sectionMap[q.section] });
          }
          sectionMap[q.section].push(q);
        });

        const scoreLabels = { 5: 'มากที่สุด', 4: 'มาก', 3: 'ปานกลาง', 2: 'น้อย', 1: 'น้อยที่สุด' };
        const shortScoreLabels = { 5: 'มากสุด', 4: 'มาก', 3: 'กลาง', 2: 'น้อย', 1: 'น้อยสุด' };
        const scoreColors = { 5: '#22c55e', 4: '#84cc16', 3: '#eab308', 2: '#f97316', 1: '#ef4444' };

        const toggleChannel = (ch) => {
          setFeedbackChannels(prev => prev.includes(ch) ? prev.filter(x => x !== ch) : [...prev, ch]);
        };
        const toggleTool = (t) => {
          setFeedbackTools(prev => prev.includes(t) ? prev.filter(x => x !== t) : [...prev, t]);
        };

        const inputStyle = {
          padding: '10px 12px', borderRadius: '8px',
          border: '1px solid var(--border-color)',
          background: 'var(--bg-color)', color: 'var(--text-primary)',
          fontSize: '0.85rem', outline: 'none', fontFamily: 'inherit', width: '100%', boxSizing: 'border-box'
        };

        return (
          <div className="modal-backdrop" onClick={closeFeedbackModal}>
            <div
              className="modal-container"
              style={{ maxWidth: '600px', padding: '0', overflow: 'hidden', maxHeight: '90vh', display: 'flex', flexDirection: 'column' }}
              onClick={(e) => e.stopPropagation()}
            >
              {/* Modal Header */}
              <div style={{
                padding: '1.25rem 1.75rem',
                borderBottom: '1px solid var(--border-color)',
                display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                background: 'linear-gradient(135deg, rgba(234,179,8,0.12), rgba(234,179,8,0.04))',
                flexShrink: 0
              }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                  <Star size={22} color="var(--accent-yellow)" />
                  <div>
                    <h3 style={{ margin: 0, fontSize: '1.05rem', fontWeight: 700, color: 'var(--accent-yellow)' }}>
                      แบบประเมินความพึงพอใจ
                    </h3>
                    <p style={{ margin: 0, fontSize: '0.75rem', color: 'var(--text-secondary)', marginTop: '2px' }}>
                      มาตราส่วน 5 ระดับ: 5 = มากที่สุด &nbsp;···&nbsp; 1 = น้อยที่สุด
                    </p>
                  </div>
                </div>
                <button className="modal-close-btn" onClick={closeFeedbackModal}>✕</button>
              </div>

              {/* Modal Scrollable Body */}
              <div style={{ overflowY: 'auto', flex: 1, padding: '1.5rem 1.75rem' }}>
                {feedbackSuccess ? (
                  <div style={{ textAlign: 'center', padding: '2.5rem 0', color: 'var(--accent-green)' }}>
                    <div style={{ fontSize: '4rem', marginBottom: '12px' }}>✓</div>
                    <h4 style={{ fontWeight: 700, fontSize: '1.1rem', margin: '0 0 8px' }}>ขอบคุณสำหรับความคิดเห็นของคุณ!</h4>
                    <p style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', margin: 0 }}>ระบบบันทึกแบบประเมินเรียบร้อยแล้ว</p>
                  </div>
                ) : (
                  <form onSubmit={submitFeedback} style={{ display: 'flex', flexDirection: 'column', gap: '1.75rem' }}>

                    {/* Score Table Header */}
                    <div style={{
                      display: 'grid',
                      gridTemplateColumns: '1fr 48px 48px 48px 48px 48px',
                      gap: '4px',
                      alignItems: 'center',
                      padding: '8px 10px',
                      background: 'rgba(234,179,8,0.06)',
                      borderRadius: '8px',
                      border: '1px solid rgba(234,179,8,0.15)',
                      fontSize: '0.7rem',
                      fontWeight: 700,
                      color: 'var(--text-secondary)',
                      letterSpacing: '0.03em'
                    }}>
                      <span>หัวข้อการประเมิน</span>
                      {[5, 4, 3, 2, 1].map(s => (
                        <span key={s} style={{ textAlign: 'center', color: scoreColors[s] }}>{s}<br /><span style={{ fontSize: '0.65rem', fontWeight: 400 }}>{shortScoreLabels[s]}</span></span>
                      ))}
                    </div>

                    {/* Sections + Questions */}
                    {sections.map((sec, si) => (
                      <div key={si} style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                        {/* Section Header */}
                        <div style={{
                          fontSize: '0.78rem', fontWeight: 700,
                          color: 'var(--accent-blue)',
                          padding: '4px 0 6px',
                          borderBottom: '1px solid rgba(59,130,246,0.2)',
                          display: 'flex', alignItems: 'center', gap: '6px'
                        }}>
                          <span style={{
                            display: 'inline-block', width: '20px', height: '20px',
                            borderRadius: '50%', background: 'rgba(59,130,246,0.2)',
                            textAlign: 'center', lineHeight: '20px', fontSize: '0.65rem'
                          }}>{si + 1}</span>
                          ด้าน{sec.name}
                        </div>

                        {/* Questions in this section */}
                        {sec.questions.map((q, qi) => (
                          <div key={q.id} style={{
                            display: 'grid',
                            gridTemplateColumns: '1fr 48px 48px 48px 48px 48px',
                            gap: '4px',
                            alignItems: 'center',
                            padding: '8px 10px',
                            borderRadius: '8px',
                            background: qi % 2 === 0 ? (theme === 'light' ? 'rgba(15,23,42,0.02)' : 'rgba(255,255,255,0.02)') : 'transparent',
                            transition: 'background 0.15s'
                          }}>
                            <span style={{ fontSize: '0.82rem', color: 'var(--text-primary)', lineHeight: 1.4, paddingRight: '8px' }}>
                              {SURVEY_QUESTIONS.indexOf(q) + 1}. {q.text}
                            </span>
                            {[5, 4, 3, 2, 1].map(score => (
                              <label key={score} style={{ display: 'flex', justifyContent: 'center', cursor: 'pointer' }}>
                                <input
                                  type="radio"
                                  name={q.id}
                                  value={score}
                                  checked={feedbackScores[q.id] === score}
                                  onChange={() => setFeedbackScores(prev => ({ ...prev, [q.id]: score }))}
                                  style={{ display: 'none' }}
                                />
                                <span style={{
                                  display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                                  width: '28px', height: '28px', borderRadius: '50%',
                                  border: `2px solid ${feedbackScores[q.id] === score ? scoreColors[score] : (theme === 'light' ? 'rgba(15,23,42,0.15)' : 'rgba(255,255,255,0.15)')}`,
                                  background: feedbackScores[q.id] === score ? scoreColors[score] + '33' : 'transparent',
                                  fontSize: '0.75rem', fontWeight: 700,
                                  color: feedbackScores[q.id] === score ? scoreColors[score] : (theme === 'light' ? 'rgba(15,23,42,0.35)' : 'rgba(255,255,255,0.35)'),
                                  transition: 'all 0.15s',
                                  userSelect: 'none'
                                }}>
                                  {score}
                                </span>
                              </label>
                            ))}
                          </div>
                        ))}
                      </div>
                    ))}

                    {/* Channel checkboxes */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                      <span style={{ fontSize: '0.82rem', fontWeight: 700, color: 'var(--accent-blue)' }}>
                        📡 ช่องทางที่รู้จัก FTC <span style={{ fontWeight: 400, color: 'var(--text-secondary)' }}>(เลือกได้มากกว่า 1 ข้อ)</span>
                      </span>
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
                        {CHANNEL_OPTIONS.map(ch => (
                          <label key={ch} style={{
                            display: 'flex', alignItems: 'center', gap: '6px',
                            padding: '6px 12px', borderRadius: '20px', cursor: 'pointer',
                            border: `1px solid ${feedbackChannels.includes(ch) ? 'rgba(59,130,246,0.6)' : (theme === 'light' ? 'rgba(15,23,42,0.12)' : 'rgba(255,255,255,0.12)')}`,
                            background: feedbackChannels.includes(ch) ? 'rgba(59,130,246,0.15)' : (theme === 'light' ? 'rgba(15,23,42,0.03)' : 'rgba(255,255,255,0.03)'),
                            fontSize: '0.8rem', color: feedbackChannels.includes(ch) ? 'var(--accent-blue)' : 'var(--text-secondary)',
                            transition: 'all 0.15s', userSelect: 'none'
                          }}>
                            <input
                              type="checkbox"
                              checked={feedbackChannels.includes(ch)}
                              onChange={() => toggleChannel(ch)}
                              style={{ display: 'none' }}
                            />
                            {feedbackChannels.includes(ch) ? '✓ ' : ''}{ch}
                          </label>
                        ))}
                      </div>
                    </div>

                    {/* Tool checkboxes */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                      <span style={{ fontSize: '0.82rem', fontWeight: 700, color: 'var(--accent-green)' }}>
                        🔬 เครื่องมือที่มาใช้บริการ <span style={{ fontWeight: 400, color: 'var(--text-secondary)' }}>(เลือกได้มากกว่า 1 ข้อ)</span>
                      </span>
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
                        {TOOL_OPTIONS.map(t => (
                          <label key={t} style={{
                            display: 'flex', alignItems: 'center', gap: '6px',
                            padding: '6px 12px', borderRadius: '20px', cursor: 'pointer',
                            border: `1px solid ${feedbackTools.includes(t) ? 'rgba(34,197,94,0.6)' : (theme === 'light' ? 'rgba(15,23,42,0.12)' : 'rgba(255,255,255,0.12)')}`,
                            background: feedbackTools.includes(t) ? 'rgba(34,197,94,0.12)' : (theme === 'light' ? 'rgba(15,23,42,0.03)' : 'rgba(255,255,255,0.03)'),
                            fontSize: '0.8rem', color: feedbackTools.includes(t) ? 'var(--accent-green)' : 'var(--text-secondary)',
                            transition: 'all 0.15s', userSelect: 'none'
                          }}>
                            <input
                              type="checkbox"
                              checked={feedbackTools.includes(t)}
                              onChange={() => toggleTool(t)}
                              style={{ display: 'none' }}
                            />
                            {feedbackTools.includes(t) ? '✓ ' : ''}{t}
                          </label>
                        ))}
                      </div>
                    </div>

                    {/* Suggestion textarea */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                      <label style={{ fontSize: '0.82rem', fontWeight: 700, color: 'var(--text-secondary)' }}>
                        💬 ความคาดหวัง / ข้อเสนอแนะเพิ่มเติม
                      </label>
                      <textarea
                        value={feedbackSuggestion}
                        onChange={(e) => setFeedbackSuggestion(e.target.value)}
                        placeholder="พิมพ์ความคิดเห็น ข้อเสนอแนะ หรือความคาดหวังของคุณที่นี่..."
                        style={{ ...inputStyle, height: '80px', resize: 'none' }}
                      />
                    </div>

                    {/* Submit Button */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', paddingTop: '4px' }}>
                      <button type="submit" className="btn btn-primary" style={{ width: '100%', height: '44px', margin: 0 }}>
                        ✅ ส่งแบบประเมิน
                      </button>
                      {window._pendingLogout && (
                        <button
                          type="button"
                          onClick={closeFeedbackModal}
                          style={{
                            width: '100%', height: '36px',
                            background: 'transparent',
                            border: '1px solid var(--border-color)',
                            borderRadius: '8px',
                            color: 'var(--text-secondary)',
                            fontSize: '0.8rem', cursor: 'pointer',
                          }}
                        >
                          ข้ามการประเมินและออกจากระบบ
                        </button>
                      )}
                    </div>

                  </form>
                )}
              </div>
            </div>
          </div>
        );
      })()}



      {/* Glassmorphic Add New Session Modal / Unified Run Setup Wizard */}
      {showAddSessionModal && (
        <div className="modal-backdrop" onClick={() => { if (!showWizardSuccess) setShowAddSessionModal(false); }}>
          <div className="modal-container" onClick={(e) => e.stopPropagation()}>

            {/* Modal Header */}
            <div className="modal-header">
              <h3>
                <FolderPlus size={22} color="var(--accent-blue)" />
                {showWizardSuccess ? 'ตั้งค่าการบันทึกข้อมูลและลูกค้าสำเร็จ' : 'ตั้งค่าเปิดงานใหม่ & จับคู่ลูกค้า (New Run Setup)'}
              </h3>
              {!showWizardSuccess && (
                <button
                  className="modal-close-btn"
                  onClick={() => setShowAddSessionModal(false)}
                  title="ปิด"
                >
                  <X size={20} />
                </button>
              )}
            </div>

            {/* Modal Body */}
            {showWizardSuccess && wizardSuccessJobId ? (
              /* Success Screen */
              (() => {
                const job = jobs.find(j => j.id === wizardSuccessJobId);
                const machine = machines.find(m => m.id === job?.machineId);
                const assignedCustomer = customers.find(c => c.machineId === job?.machineId);
                const loginUrl = `${window.location.origin}/?job=${job?.id}`;
                const expiryInfoText = job?.expiresAt ? new Date(job.expiresAt).toLocaleString('th-TH') : "ไม่มีวันหมดอายุ";
                const customerName = assignedCustomer ? assignedCustomer.companyName : "[ชื่อลูกค้า]";
                const invitationText = `เรียนคุณ ${customerName},\n\nทางแล็บขอส่งลิงก์สำหรับเข้าดูข้อมูลไบโอโพรเซสรอบรัน "${job?.name || ''}" (${machine?.name || ''}) แบบเรียลไทม์\n\nลิงก์เข้าสู่ระบบ: ${loginUrl}\nวันหมดอายุ: ${expiryInfoText}\n\nขอบคุณค่ะ/ครับ\nDBMS System`;

                return (
                  <div className="modal-body">
                    <div style={{ textAlign: 'center', margin: '1rem 0' }}>
                      <span style={{ fontSize: '3rem' }}>🚀</span>
                      <h4 style={{ fontSize: '1.25rem', color: 'var(--accent-green)', fontWeight: 700, marginTop: '8px' }}>ตั้งค่าเปิดรอบรันข้อมูลสำเร็จ!</h4>
                    </div>

                    <div className="modal-info-box">
                      <div>
                        <span style={{ color: 'var(--text-secondary)' }}>รอบรัน (Session): </span>
                        <strong style={{ color: 'var(--accent-blue)' }}>{job?.name}</strong>
                      </div>
                      <div>
                        <span style={{ color: 'var(--text-secondary)' }}>เครื่องมือ: </span>
                        <strong>{machine?.name || '-'}</strong>
                      </div>
                    </div>

                    {/* Expiry Label */}
                    <div style={{ fontSize: '0.9rem', color: 'var(--text-secondary)' }}>
                      ⏱️ สิทธิ์การเข้าใช้งานหมดอายุวันที่: <strong style={{ color: 'var(--accent-yellow)' }}>{expiryInfoText}</strong>
                    </div>

                    {/* Share Link */}
                    <div>
                      <label className="modal-label">🔗 ลิงก์เข้าใช้งานตรงสำหรับลูกค้า (Direct Share Link)</label>
                      <div className="modal-input-row">
                        <input
                          type="text"
                          readOnly
                          value={loginUrl}
                          className="modal-input"
                          onClick={(e) => e.target.select()}
                        />
                        <button
                          className="btn btn-blue"
                          style={{ margin: 0 }}
                          onClick={async () => {
                            await navigator.clipboard.writeText(loginUrl);
                            alert("คัดลอกลิงก์สำเร็จ!");
                          }}
                        >
                          คัดลอกลิงก์
                        </button>
                      </div>
                    </div>

                    {/* Email Invitation text */}
                    <div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.5rem' }}>
                        <label className="modal-label" style={{ margin: 0 }}>📋 ข้อความคำเชิญสำหรับลูกค้า</label>
                        <button
                          type="button"
                          className="btn btn-primary btn-small"
                          style={{ margin: 0 }}
                          onClick={async () => {
                            await navigator.clipboard.writeText(invitationText);
                            alert("คัดลอกข้อความคำเชิญสำเร็จ!");
                          }}
                        >
                          📋 คัดลอกคำเชิญ
                        </button>
                      </div>
                      <div className="modal-preview-box">
                        {invitationText}
                      </div>
                    </div>

                    {/* Actions */}
                    <div style={{ display: 'flex', justifyContent: 'center', marginTop: '1rem', borderTop: '1px solid var(--border-color)', paddingTop: '1.25rem' }}>
                      <button
                        className="btn btn-primary"
                        style={{ margin: 0 }}
                        onClick={() => {
                          setCurrentMachineId(job.machineId);
                          setCurrentJobId(job.id);
                          setCurrentAppView('monitoring');
                          setActiveTab('dashboard');
                          setShowAddSessionModal(false);
                          setShowWizardSuccess(false);
                          setWizardSuccessJobId(null);
                        }}
                      >
                        เริ่มเข้าดูแดชบอร์ดข้อมูล (Open Dashboard)
                      </button>
                    </div>

                  </div>
                );
              })()
            ) : (
              /* Setup Form */
              <form onSubmit={submitNewJob} className="modal-body" style={{ margin: 0 }}>

                {/* Step 1: Run Config */}
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem' }}>
                  <div>
                    <label className="modal-label">🖥️ เครื่องมือ / อุปกรณ์ *</label>
                    <select
                      value={newSessionMachineId}
                      onChange={(e) => setNewSessionMachineId(e.target.value)}
                      className="modal-input"
                      style={{ width: '100%', padding: '10px', height: '42px', backgroundImage: 'none' }}
                      required
                    >
                      {machines.map(m => (
                        <option key={m.id} value={m.id}>{m.name}</option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="modal-label">📝 ชื่อรอบบันทึกข้อมูล *</label>
                    <input
                      type="text"
                      value={newSessionName}
                      onChange={(e) => setNewSessionName(e.target.value)}
                      placeholder="เช่น Session 1 หรือ Batch-01"
                      className="modal-input"
                      required
                      style={{ width: '100%', padding: '10px' }}
                    />
                  </div>
                </div>

                {/* Step 2: Customer Assignment */}
                <div className="modal-section">
                  <label className="modal-label">👥 ระบุลูกค้าผู้เข้าใช้งาน (Customer Assignment)</label>
                  <select
                    value={newSessionCustomerId}
                    onChange={(e) => {
                      setNewSessionCustomerId(e.target.value);
                      if (e.target.value !== 'ADD_NEW') {
                        setNewSessionCustomerName('');
                        setNewSessionCustomerEmail('');
                      }
                    }}
                    className="modal-input"
                    style={{ width: '100%', padding: '10px', height: '42px', backgroundImage: 'none', marginBottom: newSessionCustomerId === 'ADD_NEW' ? '12px' : '0' }}
                  >
                    <option value="">-- ไม่ระบุลูกค้า (No Customer) --</option>
                    <option value="ADD_NEW">➕ เพิ่มลูกค้าใหม่ (Create New Customer)</option>
                    {customers.map(c => (
                      <option key={c.id} value={c.id}>{c.companyName} ({getMachineName(c.machineId)})</option>
                    ))}
                  </select>

                  {/* Add New Customer Fields */}
                  {newSessionCustomerId === 'ADD_NEW' && (
                    <div style={{ display: 'flex', gap: '1rem', background: 'rgba(255, 255, 255, 0.02)', padding: '12px', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                      <div style={{ flex: 1 }}>
                        <span style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>ชื่อบริษัท/ลูกค้า *</span>
                        <input
                          type="text"
                          value={newSessionCustomerName}
                          onChange={(e) => setNewSessionCustomerName(e.target.value)}
                          placeholder="ชื่อบริษัท/ชื่อลูกค้า"
                          className="modal-input"
                          required
                          style={{ width: '100%', marginTop: '4px', height: '36px', padding: '6px 10px' }}
                        />
                      </div>
                      <div style={{ flex: 1 }}>
                        <span style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>อีเมล (Email)</span>
                        <input
                          type="email"
                          value={newSessionCustomerEmail}
                          onChange={(e) => setNewSessionCustomerEmail(e.target.value)}
                          placeholder="customer@example.com"
                          className="modal-input"
                          style={{ width: '100%', marginTop: '4px', height: '36px', padding: '6px 10px' }}
                        />
                      </div>
                    </div>
                  )}
                </div>

                {/* Step 3: Link Expiration */}
                <div className="modal-section">
                  <label className="modal-label">⏱️ กำหนดวันหมดอายุลิงก์แชร์ลูกค้า (Link Expiration)</label>
                  <div className="modal-preset-row" style={{ marginBottom: '4px' }}>
                    {[
                      { label: '24 ชม. (1 วัน)', value: 24 },
                      { label: '72 ชม. (3 วัน)', value: 72 },
                      { label: '168 ชม. (7 วัน)', value: 168 },
                      { label: 'ไม่จำกัด (Unlimited)', value: 0 }
                    ].map(preset => (
                      <button
                        key={preset.value}
                        type="button"
                        className="modal-preset-btn"
                        style={{
                          border: newSessionExpiryHours === preset.value ? '1px solid var(--accent-blue)' : '1px solid var(--border-color)',
                          background: newSessionExpiryHours === preset.value ? 'rgba(59, 130, 246, 0.15)' : 'transparent',
                          fontWeight: newSessionExpiryHours === preset.value ? 700 : 'normal',
                          color: newSessionExpiryHours === preset.value ? '#ffffff' : 'var(--text-primary)'
                        }}
                        onClick={() => setNewSessionExpiryHours(preset.value)}
                      >
                        {preset.label}
                      </button>
                    ))}

                    {/* Expiry info text */}
                    <span style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', marginLeft: '10px' }}>
                      {newSessionExpiryHours > 0
                        ? `(หมดอายุอีกใน ${newSessionExpiryHours} ชม.)`
                        : '(สามารถเข้าดูได้ตลอดเวลา)'}
                    </span>
                  </div>
                </div>

                {/* Actions */}
                <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '0.75rem', marginTop: '1rem', borderTop: '1px solid var(--border-color)', paddingTop: '1.25rem' }}>
                  <button
                    type="button"
                    className="btn btn-secondary"
                    onClick={() => setShowAddSessionModal(false)}
                  >
                    ยกเลิก (Cancel)
                  </button>
                  <button
                    type="submit"
                    className="btn btn-blue"
                  >
                    สร้างและบันทึกข้อมูล (Create & Setup)
                  </button>
                </div>

              </form>
            )}

          </div>
        </div>
      )}

      {/* Glassmorphic Edit Instrument Modal */}
      {showEditMachineModal && (
        <div className="modal-backdrop" onClick={() => setShowEditMachineModal(false)}>
          <div className="modal-container" onClick={(e) => e.stopPropagation()}>

            {/* Modal Header */}
            <div className="modal-header">
              <h3>
                <Cpu size={22} color="var(--accent-blue)" />
                แก้ไขชื่อเครื่องมือ (Edit Instrument Name)
              </h3>
              <button
                className="modal-close-btn"
                onClick={() => setShowEditMachineModal(false)}
                title="ปิด"
              >
                <X size={20} />
              </button>
            </div>

            {/* Modal Body */}
            <form onSubmit={submitEditMachine} className="modal-body" style={{ margin: 0 }}>

              <div>
                <label className="modal-label">🖥️ ชื่อเครื่องมือ / bioreactor *</label>
                <input
                  type="text"
                  value={editingMachineName}
                  onChange={(e) => setEditingMachineName(e.target.value)}
                  placeholder="กรอกชื่อเครื่องมือ"
                  className="modal-input"
                  required
                  style={{ width: '100%', padding: '10px' }}
                />
              </div>

              {/* Actions */}
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '0.75rem', marginTop: '1rem', borderTop: '1px solid var(--border-color)', paddingTop: '1.25rem' }}>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={() => setShowEditMachineModal(false)}
                >
                  ยกเลิก
                </button>
                <button
                  type="submit"
                  className="btn btn-blue"
                >
                  บันทึก
                </button>
              </div>

            </form>

          </div>
        </div>
      )}

      {/* Glassmorphic Add New Instrument Modal */}
      {showAddMachineModal && (
        <div className="modal-backdrop" onClick={() => setShowAddMachineModal(false)}>
          <div className="modal-container" onClick={(e) => e.stopPropagation()}>

            {/* Modal Header */}
            <div className="modal-header">
              <h3>
                <Cpu size={22} color="var(--accent-blue)" />
                เพิ่มเครื่องมือ / อุปกรณ์ใหม่ (Add Instrument)
              </h3>
              <button
                className="modal-close-btn"
                onClick={() => setShowAddMachineModal(false)}
                title="ปิด"
              >
                <X size={20} />
              </button>
            </div>

            {/* Modal Body */}
            <form onSubmit={submitNewMachine} className="modal-body" style={{ margin: 0 }}>

              <div>
                <label className="modal-label">🖥️ ชื่อเครื่องมือ / bioreactor เครื่องใหม่ *</label>
                <input
                  type="text"
                  value={newMachineName}
                  onChange={(e) => setNewMachineName(e.target.value)}
                  placeholder="เช่น Bioreactor 2 หรือ Fermenter 20L"
                  className="modal-input"
                  required
                  style={{ width: '100%', padding: '10px' }}
                />
              </div>

              {/* Actions */}
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '0.75rem', marginTop: '1rem', borderTop: '1px solid var(--border-color)', paddingTop: '1.25rem' }}>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={() => setShowAddMachineModal(false)}
                >
                  ยกเลิก
                </button>
                <button
                  type="submit"
                  className="btn btn-blue"
                >
                  สร้างเครื่องมือใหม่
                </button>
              </div>

            </form>

          </div>
        </div>
      )}
    </div>
  );
}

export default App;
