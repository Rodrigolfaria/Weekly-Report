#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import html
import json
import re
from io import BytesIO
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import ZipFile


SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOCUMENT_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"a": SPREADSHEET_NS, "r": DOCUMENT_REL_NS, "pr": PACKAGE_REL_NS}

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>__TITLE__</title>
  <style>
    :root {
      --bg: #f4f7fb;
      --paper: rgba(255, 255, 255, 0.88);
      --panel: #ffffff;
      --panel-alt: #eef6ff;
      --ink: #102033;
      --muted: #546579;
      --line: #d8e2ef;
      --accent: #1264d6;
      --accent-2: #0f766e;
      --accent-3: #c06a0a;
      --accent-4: #c81e5a;
      --shadow: 0 24px 54px rgba(15, 23, 42, 0.10);
      --radius: 24px;
    }

    * {
      box-sizing: border-box;
    }

    body {
      margin: 0;
      font-family: "Avenir Next", "Segoe UI", sans-serif;
      color: var(--ink);
      background:
        radial-gradient(circle at 10% 0%, rgba(18, 100, 214, 0.18), transparent 28%),
        radial-gradient(circle at 95% 15%, rgba(15, 118, 110, 0.14), transparent 24%),
        linear-gradient(180deg, #eef4fb 0%, var(--bg) 25%, var(--bg) 100%);
    }

    .page {
      width: min(1400px, calc(100vw - 28px));
      margin: 0 auto;
      padding: 20px 0 44px;
    }

    .hero {
      margin-top: 10px;
      padding: 28px;
      border-radius: 30px;
      background:
        linear-gradient(135deg, rgba(16, 32, 51, 0.96), rgba(18, 100, 214, 0.94) 56%, rgba(15, 118, 110, 0.94));
      color: white;
      box-shadow: var(--shadow);
      backdrop-filter: blur(12px);
    }

    .hero-grid {
      display: grid;
      grid-template-columns: minmax(0, 1.35fr) minmax(320px, 0.85fr);
      gap: 24px;
      align-items: start;
    }

    .eyebrow {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 7px 12px;
      border-radius: 999px;
      background: rgba(255, 255, 255, 0.14);
      font-size: 12px;
      letter-spacing: 0.08em;
      text-transform: uppercase;
    }

    h1 {
      margin: 14px 0 10px;
      font-size: clamp(30px, 4vw, 48px);
      line-height: 1.03;
      letter-spacing: -0.03em;
    }

    .hero p {
      margin: 8px 0;
      max-width: 760px;
      color: rgba(255, 255, 255, 0.86);
      line-height: 1.55;
      font-size: 15px;
    }

    .hero-meta {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-top: 16px;
    }

    .hero-chip {
      padding: 9px 12px;
      border-radius: 999px;
      background: rgba(255, 255, 255, 0.12);
      border: 1px solid rgba(255, 255, 255, 0.14);
      font-size: 13px;
      color: rgba(255, 255, 255, 0.9);
    }

    .hero-side {
      background: rgba(255, 255, 255, 0.12);
      border: 1px solid rgba(255, 255, 255, 0.12);
      border-radius: 24px;
      padding: 18px 18px 16px;
    }

    .hero-side h2 {
      margin: 0 0 10px;
      font-size: 16px;
      color: rgba(255, 255, 255, 0.95);
    }

    .theme-toggle-wrap {
      display: grid;
      gap: 10px;
      margin-top: 14px;
    }

    .theme-toggle-row {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
    }

    .theme-toggle-label {
      font-size: 12px;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      color: rgba(255, 255, 255, 0.82);
      font-weight: 700;
    }

    .theme-toggle-state {
      font-size: 13px;
      color: rgba(255, 255, 255, 0.88);
    }

    .theme-toggle {
      appearance: none;
      border: 1px solid rgba(255, 255, 255, 0.20);
      background: rgba(255, 255, 255, 0.14);
      color: white;
      border-radius: 999px;
      padding: 5px;
      width: 100%;
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: flex-start;
      transition: background 180ms ease, border-color 180ms ease;
    }

    .theme-toggle-thumb {
      width: 50%;
      min-width: 138px;
      padding: 10px 12px;
      border-radius: 999px;
      background: rgba(255, 255, 255, 0.92);
      color: #102033;
      font-size: 13px;
      font-weight: 800;
      text-align: center;
      transition: transform 180ms ease, background 180ms ease, color 180ms ease;
    }

    .preset-grid,
    .toggle-grid,
    .filter-grid {
      display: grid;
      gap: 12px;
    }

    .preset-grid {
      grid-template-columns: repeat(2, minmax(0, 1fr));
      margin-top: 14px;
    }

    .preset-btn,
    .ghost-btn {
      appearance: none;
      border: 1px solid rgba(255, 255, 255, 0.2);
      background: rgba(255, 255, 255, 0.12);
      color: white;
      border-radius: 14px;
      padding: 10px 12px;
      font: inherit;
      cursor: pointer;
      transition: transform 140ms ease, background 140ms ease, border-color 140ms ease;
    }

    .preset-btn:hover,
    .ghost-btn:hover,
    .preset-btn.is-active {
      background: rgba(255, 255, 255, 0.22);
      border-color: rgba(255, 255, 255, 0.28);
      transform: translateY(-1px);
    }

    .shell {
      display: grid;
      grid-template-columns: 320px minmax(0, 1fr);
      gap: 24px;
      margin-top: 26px;
      align-items: start;
    }

    .sidebar {
      position: sticky;
      top: 24px;
      display: grid;
      gap: 18px;
    }

    .panel {
      background: var(--paper);
      border: 1px solid rgba(216, 226, 239, 0.92);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
      backdrop-filter: blur(12px);
    }

    .panel-inner {
      padding: 20px;
    }

    .panel h2,
    .panel h3 {
      margin: 0 0 14px;
    }

    .panel-note {
      margin: 0 0 14px;
      color: var(--muted);
      line-height: 1.45;
      font-size: 14px;
    }

    .filter-grid {
      grid-template-columns: 1fr;
    }

    .field {
      display: grid;
      gap: 6px;
    }

    .field label {
      font-size: 12px;
      font-weight: 700;
      letter-spacing: 0.04em;
      text-transform: uppercase;
      color: var(--muted);
    }

    .field input,
    .field select {
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 11px 12px;
      font: inherit;
      color: var(--ink);
      background: white;
    }

    .field input:focus,
    .field select:focus {
      outline: 2px solid rgba(18, 100, 214, 0.18);
      border-color: rgba(18, 100, 214, 0.44);
    }

    .toggle-grid {
      grid-template-columns: 1fr 1fr;
    }

    .toggle {
      display: flex;
      align-items: center;
      gap: 10px;
      padding: 10px 12px;
      border-radius: 16px;
      background: white;
      border: 1px solid var(--line);
      font-size: 14px;
    }

    .toggle input {
      width: 18px;
      height: 18px;
      accent-color: var(--accent);
    }

    .main {
      display: grid;
      gap: 22px;
    }

    .section {
      padding: 22px;
    }

    .section-header {
      display: flex;
      justify-content: space-between;
      gap: 18px;
      align-items: flex-start;
      flex-wrap: wrap;
      margin-bottom: 18px;
    }

    .section-title {
      margin: 0;
      font-size: 24px;
      letter-spacing: -0.02em;
    }

    .section-subtitle {
      margin: 6px 0 0;
      font-size: 14px;
      color: var(--muted);
      max-width: 760px;
      line-height: 1.55;
    }

    .status-box {
      min-width: 260px;
      padding: 14px 16px;
      border-radius: 18px;
      background: linear-gradient(180deg, #ffffff, #f7fbff);
      border: 1px solid var(--line);
    }

    .status-box strong {
      display: block;
      margin-bottom: 6px;
      font-size: 14px;
    }

    .status-box span {
      display: block;
      color: var(--muted);
      font-size: 14px;
      line-height: 1.45;
    }

    .chips {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
    }

    .chip {
      padding: 7px 11px;
      border-radius: 999px;
      background: var(--panel-alt);
      border: 1px solid #bfdbfe;
      color: #1d4ed8;
      font-size: 13px;
    }

    .kpis {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
      gap: 16px;
    }

    .card {
      padding: 18px;
      border-radius: 20px;
      background: linear-gradient(180deg, #ffffff, #f9fbff);
      border: 1px solid var(--line);
      overflow: hidden;
      position: relative;
    }

    .card::after {
      content: "";
      position: absolute;
      inset: auto -22px -32px auto;
      width: 90px;
      height: 90px;
      border-radius: 50%;
      background: radial-gradient(circle, rgba(18, 100, 214, 0.14), transparent 72%);
    }

    .card-label {
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: var(--muted);
      margin-bottom: 8px;
    }

    .card-value {
      font-size: 34px;
      line-height: 1;
      font-weight: 800;
      margin-bottom: 10px;
      letter-spacing: -0.03em;
    }

    .card-meta {
      font-size: 14px;
      color: var(--muted);
      line-height: 1.45;
    }

    .grid-2,
    .grid-3 {
      display: grid;
      gap: 18px;
    }

    .grid-2 {
      grid-template-columns: repeat(auto-fit, minmax(330px, 1fr));
    }

    .grid-3 {
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
    }

    .chart {
      padding: 18px;
      border-radius: 22px;
      background: linear-gradient(180deg, #ffffff, #f8fbff);
      border: 1px solid var(--line);
    }

    .chart h3 {
      margin: 0 0 14px;
      font-size: 16px;
    }

    .bar-chart,
    .line-chart,
    .table-wrap {
      width: 100%;
    }

    .bar-list {
      display: grid;
      gap: 12px;
    }

    .bar-row {
      display: grid;
      grid-template-columns: minmax(0, 190px) minmax(0, 1fr) auto;
      gap: 12px;
      align-items: center;
    }

    .bar-label {
      font-size: 13px;
      color: var(--ink);
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    .bar-track {
      width: 100%;
      height: 14px;
      border-radius: 999px;
      background: #e6edf7;
      overflow: hidden;
    }

    .bar-fill {
      height: 100%;
      border-radius: inherit;
      min-width: 2px;
    }

    .bar-value {
      min-width: 52px;
      text-align: right;
      font-size: 13px;
      color: var(--muted);
    }

    .line-svg {
      display: block;
      width: 100%;
      height: auto;
      overflow: visible;
    }

    .column-chart-wrap {
      width: 100%;
      overflow-x: auto;
      padding-bottom: 6px;
      min-width: 0;
      max-width: 100%;
    }

    .column-chart-svg {
      display: block;
      width: 100%;
      height: auto;
      min-width: 560px;
    }

    .table-wrap {
      overflow: auto;
      border: 1px solid var(--line);
      border-radius: 18px;
      background: white;
      min-width: 0;
      max-width: 100%;
    }

    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 14px;
    }

    th,
    td {
      padding: 12px 14px;
      border-bottom: 1px solid var(--line);
      text-align: left;
      vertical-align: top;
    }

    th {
      position: sticky;
      top: 0;
      background: #edf5ff;
      color: #143b79;
      z-index: 1;
    }

    tr:nth-child(even) td {
      background: #fbfdff;
    }

    .stats-table thead tr:first-child th {
      background: #ffffff;
      color: #111827;
      font-size: 13px;
      text-align: center;
      border-bottom: 0;
    }

    .stats-table thead tr:nth-child(2) th {
      background: #ffffff;
      color: #111827;
      font-size: 13px;
      text-align: center;
      border-bottom: 0;
    }

    .stats-table thead tr:nth-child(3) th {
      background: #f3f7fc;
      color: #163a78;
      text-align: center;
    }

    .stats-table tbody td {
      text-align: center;
    }

    .stats-table tbody td:nth-child(1),
    .stats-table tbody td:nth-child(2) {
      text-align: left;
      font-weight: 600;
    }

    .empty {
      padding: 28px 14px;
      text-align: center;
      color: var(--muted);
      border: 1px dashed var(--line);
      border-radius: 18px;
      background: #fbfdff;
    }

    .two-col {
      display: grid;
      grid-template-columns: 1.15fr 0.85fr;
      gap: 18px;
      align-items: start;
    }

    .footer-note {
      font-size: 13px;
      color: var(--muted);
      line-height: 1.55;
    }

    .view-tabs {
      display: flex;
      gap: 12px;
      flex-wrap: wrap;
      margin-top: 22px;
    }

    .view-tab {
      appearance: none;
      border: 1px solid #c8d8ec;
      background: rgba(255, 255, 255, 0.84);
      color: var(--ink);
      border-radius: 999px;
      padding: 12px 16px;
      font: inherit;
      font-weight: 700;
      cursor: pointer;
      box-shadow: var(--shadow);
    }

    .view-tab.is-active {
      background: linear-gradient(135deg, #102033, #1264d6);
      color: white;
      border-color: transparent;
    }

    .view-panel {
      margin-top: 24px;
    }

    .weekly-root {
      display: grid;
      gap: 22px;
    }

    .flat-time-root {
      display: grid;
      gap: 22px;
    }

    .flat-time-toolbar {
      display: flex;
      gap: 16px;
      align-items: end;
      justify-content: space-between;
      flex-wrap: wrap;
    }

    .flat-time-upload-panel {
      display: flex;
      gap: 12px;
      align-items: center;
      flex-wrap: wrap;
    }

    .file-input {
      width: 100%;
      border: 1px dashed var(--line);
      border-radius: 14px;
      padding: 10px 12px;
      font: inherit;
      color: var(--ink);
      background: white;
    }

    .tag-list {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-top: 10px;
    }

    .tag {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 7px 11px;
      border-radius: 999px;
      background: var(--panel-alt);
      border: 1px solid var(--line);
      color: var(--ink);
      font-size: 13px;
    }

    .tag-muted {
      color: var(--muted);
    }

    .tag button {
      appearance: none;
      border: 0;
      background: transparent;
      color: inherit;
      cursor: pointer;
      font: inherit;
      padding: 0;
    }

    .weekly-toolbar {
      display: flex;
      gap: 16px;
      align-items: end;
      justify-content: space-between;
      flex-wrap: wrap;
    }

    .weekly-toolbar-actions {
      display: flex;
      gap: 12px;
      align-items: center;
      flex-wrap: wrap;
    }

    .action-btn {
      appearance: none;
      border: 1px solid #0f4fb4;
      background: linear-gradient(135deg, #102033, #1264d6);
      color: white;
      border-radius: 14px;
      padding: 12px 16px;
      font: inherit;
      font-weight: 700;
      cursor: pointer;
      box-shadow: 0 12px 28px rgba(18, 100, 214, 0.22);
      transition: transform 140ms ease, box-shadow 140ms ease, border-color 140ms ease;
    }

    .action-btn:hover {
      transform: translateY(-1px);
      border-color: #0b46a3;
      box-shadow: 0 16px 30px rgba(18, 100, 214, 0.28);
    }

    .weekly-banner {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      flex-wrap: wrap;
      padding: 18px 20px;
      border-radius: 22px;
      background: linear-gradient(135deg, #102033, #1264d6 56%, #0f766e);
      color: white;
    }

    .weekly-banner h2 {
      margin: 0;
      font-size: clamp(24px, 3vw, 34px);
      letter-spacing: -0.03em;
    }

    .weekly-banner p {
      margin: 6px 0 0;
      color: rgba(255, 255, 255, 0.84);
      font-size: 14px;
      line-height: 1.5;
    }

    .weekly-banner .chip {
      background: rgba(255, 255, 255, 0.14);
      border-color: rgba(255, 255, 255, 0.14);
      color: white;
    }

    .report-grid-2,
    .report-grid-3,
    .history-grid {
      display: grid;
      gap: 18px;
    }

    .report-grid-2 {
      grid-template-columns: repeat(auto-fit, minmax(360px, 1fr));
    }

    .report-grid-2 > *,
    .report-grid-3 > *,
    .flat-time-chart-stack > *,
    .drill-grid > * {
      min-width: 0;
    }

    .flat-time-chart-stack {
      display: grid;
      gap: 18px;
    }

    .report-grid-3 {
      grid-template-columns: repeat(2, minmax(0, 1fr));
    }

    .history-grid {
      grid-template-columns: repeat(auto-fit, minmax(420px, 1fr));
    }

    .report-card {
      padding: 20px;
      border-radius: 22px;
      background: linear-gradient(180deg, #ffffff, #f8fbff);
      border: 1px solid var(--line);
      min-width: 0;
      max-width: 100%;
      overflow: hidden;
    }

    .report-card h3 {
      margin: 0 0 8px;
      font-size: 17px;
    }

    .report-note {
      margin: 0 0 14px;
      font-size: 13px;
      color: var(--muted);
      line-height: 1.5;
    }

    .confidence-badge,
    .drill-chip,
    .table-action {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 6px;
      border-radius: 999px;
      padding: 5px 10px;
      font-size: 12px;
      font-weight: 800;
      letter-spacing: 0.02em;
      border: 1px solid transparent;
    }

    .confidence-badge.high {
      color: #166534;
      background: rgba(34, 197, 94, 0.16);
      border-color: rgba(34, 197, 94, 0.28);
    }

    .confidence-badge.medium {
      color: #92400e;
      background: rgba(245, 158, 11, 0.16);
      border-color: rgba(245, 158, 11, 0.28);
    }

    .confidence-badge.low {
      color: #991b1b;
      background: rgba(239, 68, 68, 0.16);
      border-color: rgba(239, 68, 68, 0.28);
    }

    .trend-indicator {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      font-weight: 800;
      white-space: nowrap;
    }

    .trend-indicator.fast {
      color: #166534;
    }

    .trend-indicator.slow {
      color: #991b1b;
    }

    .trend-indicator .arrow {
      font-size: 13px;
      line-height: 1;
    }

    .table-action {
      appearance: none;
      background: rgba(18, 100, 214, 0.08);
      border-color: rgba(18, 100, 214, 0.18);
      color: var(--ink);
      cursor: pointer;
      font: inherit;
      text-align: left;
      padding: 6px 10px;
      max-width: 100%;
      white-space: normal;
      word-break: break-word;
    }

    .table-action:hover {
      background: rgba(18, 100, 214, 0.16);
    }

    .drill-grid {
      display: grid;
      gap: 18px;
      grid-template-columns: repeat(auto-fit, minmax(340px, 1fr));
    }

    .drill-list {
      display: grid;
      gap: 10px;
    }

    .drill-item {
      padding: 12px 14px;
      border-radius: 16px;
      background: var(--panel-alt);
      border: 1px solid var(--line);
    }

    .drill-item strong {
      display: block;
      margin-bottom: 4px;
      color: var(--ink);
    }

    .drill-item span {
      display: block;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.45;
    }

    .metric-strip {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 14px;
    }

    .metric-pill {
      padding: 16px 18px;
      border-radius: 20px;
      background: linear-gradient(180deg, #ffffff, #f8fbff);
      border: 1px solid var(--line);
    }

    .metric-pill .label {
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: var(--muted);
      margin-bottom: 8px;
    }

    .metric-pill .value {
      display: flex;
      align-items: baseline;
      gap: 8px;
      flex-wrap: wrap;
      line-height: 1.05;
      letter-spacing: -0.03em;
      margin-bottom: 8px;
    }

    .metric-pill .value-main {
      font-size: 30px;
      font-weight: 800;
      color: var(--ink);
    }

    .metric-pill .value-suffix {
      font-size: 14px;
      font-weight: 600;
      color: var(--muted);
      letter-spacing: 0;
      text-transform: none;
    }

    .metric-pill .meta {
      font-size: 14px;
      color: var(--muted);
      line-height: 1.45;
    }

    .legend {
      display: flex;
      gap: 12px;
      flex-wrap: wrap;
      margin: 10px 0 16px;
    }

    .legend-item {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      font-size: 15px;
      color: var(--muted);
    }

    .legend-dot {
      width: 14px;
      height: 14px;
      border-radius: 50%;
      flex: 0 0 auto;
    }

    body.theme-corona {
      --bg: #0f1015;
      --paper: #191c24;
      --panel: #191c24;
      --panel-alt: #0d1118;
      --ink: #f5f5f5;
      --muted: #a1aab8;
      --line: #2c2e33;
      --accent: #0090e7;
      --accent-2: #00d25b;
      --accent-3: #ffab00;
      --accent-4: #fc424a;
      --shadow: 0 18px 34px rgba(0, 0, 0, 0.28);
      --radius: 10px;
      background:
        radial-gradient(circle at 0% 0%, rgba(0, 144, 231, 0.16), transparent 22%),
        radial-gradient(circle at 100% 0%, rgba(0, 210, 91, 0.12), transparent 20%),
        linear-gradient(180deg, #0a0b0f 0%, #0f1015 100%);
    }

    body.theme-corona .hero {
      border-radius: 14px;
      background: linear-gradient(135deg, #191c24, #111318 58%, #151922);
      border: 1px solid #2c2e33;
      box-shadow: 0 18px 30px rgba(0, 0, 0, 0.28);
    }

    body.theme-corona .hero-side,
    body.theme-corona .panel,
    body.theme-corona .card,
    body.theme-corona .chart,
    body.theme-corona .report-card,
    body.theme-corona .metric-pill,
    body.theme-corona .status-box,
    body.theme-corona .table-wrap {
      background: #191c24;
      border-color: #2c2e33;
      box-shadow: none;
      backdrop-filter: none;
    }

    body.theme-corona .card::after {
      background: radial-gradient(circle, rgba(0, 144, 231, 0.18), transparent 72%);
    }

    body.theme-corona .view-tab {
      border-color: #2c2e33;
      background: #191c24;
      color: #d5d9e0;
      box-shadow: none;
      border-radius: 8px;
    }

    body.theme-corona .view-tab.is-active,
    body.theme-corona .action-btn {
      background: linear-gradient(135deg, #0090e7, #0069aa);
      color: white;
      border-color: transparent;
      box-shadow: none;
    }

    body.theme-corona .chip,
    body.theme-corona .weekly-banner .chip {
      background: rgba(0, 144, 231, 0.12);
      border-color: rgba(0, 144, 231, 0.24);
      color: #8fd8ff;
    }

    body.theme-corona .weekly-banner {
      border-radius: 12px;
      border: 1px solid #2c2e33;
      background: linear-gradient(135deg, #191c24, #12151b);
    }

    body.theme-corona .field input,
    body.theme-corona .field select,
    body.theme-corona .toggle,
    body.theme-corona .theme-toggle,
    body.theme-corona .file-input {
      background: #0f1015;
      color: #f5f5f5;
      border-color: #2c2e33;
    }

    body.theme-corona .field input::placeholder {
      color: #7f8896;
    }

    body.theme-corona .field input:focus,
    body.theme-corona .field select:focus {
      outline-color: rgba(0, 144, 231, 0.22);
      border-color: rgba(0, 144, 231, 0.56);
    }

    body.theme-corona .theme-toggle-thumb {
      transform: translateX(100%);
      background: linear-gradient(135deg, #0090e7, #0069aa);
      color: white;
    }

    body.theme-corona th {
      background: #0f1015;
      color: #8fd8ff;
      border-bottom-color: #2c2e33;
    }

    body.theme-corona td {
      border-bottom-color: #2c2e33;
    }

    body.theme-corona tr:nth-child(even) td,
    body.theme-corona .empty {
      background: #111318;
    }

    body.theme-corona .stats-table thead tr:first-child th,
    body.theme-corona .stats-table thead tr:nth-child(2) th,
    body.theme-corona .stats-table thead tr:nth-child(3) th {
      background: #0f1015;
      color: #d5d9e0;
    }

    body.theme-corona .bar-track {
      background: #0f1015;
    }

    body.theme-corona .tag {
      background: rgba(0, 144, 231, 0.12);
      border-color: rgba(0, 144, 231, 0.24);
      color: #d5d9e0;
    }

    body.theme-corona .eyebrow,
    body.theme-corona .hero-chip {
      background: rgba(255, 255, 255, 0.07);
      border-color: rgba(255, 255, 255, 0.08);
    }

    [hidden] {
      display: none !important;
    }

    @media (max-width: 1120px) {
      .shell,
      .hero-grid,
      .two-col,
      .history-grid {
        grid-template-columns: 1fr;
      }

      .report-grid-3 {
        grid-template-columns: 1fr;
      }

      .sidebar {
        position: static;
      }

      .hero {
        position: static;
      }
    }

    @media (max-width: 720px) {
      .page {
        width: min(100vw - 18px, 1400px);
      }

      .hero,
      .section,
      .panel-inner {
        padding: 18px;
      }

      .toggle-grid,
      .preset-grid,
      .view-tabs {
        grid-template-columns: 1fr;
      }

      .bar-row {
        grid-template-columns: 1fr;
        gap: 8px;
      }

      .bar-value {
        text-align: left;
      }

    }

    @page {
      size: A4 landscape;
      margin: 10mm;
    }

    @media print {
      body {
        background: white;
      }

      .page {
        width: 100%;
        margin: 0;
        padding: 0;
      }

      .hero,
      .view-tabs,
      #dashboard-view,
      .weekly-toolbar .field,
      .weekly-toolbar .action-btn {
        display: none !important;
      }

      #weekly-report-view,
      #weekly-report-view[hidden] {
        display: block !important;
        margin-top: 0;
      }

      .panel,
      .report-card,
      .metric-pill,
      .chart,
      .table-wrap,
      .weekly-banner,
      .section,
      .report-grid-2 > *,
      .report-grid-3 > * {
        break-inside: avoid;
        page-break-inside: avoid;
        box-shadow: none !important;
      }

      .panel,
      .report-card,
      .chart,
      .metric-pill,
      .table-wrap {
        background: white !important;
        backdrop-filter: none !important;
      }

      .section {
        padding: 12px 0;
      }

      .section-header {
        margin-bottom: 10px;
      }

      .section-subtitle,
      .weekly-banner p,
      .panel-note,
      .footer-note {
        color: #374151 !important;
      }

      .status-box {
        min-width: 0;
        padding: 10px 12px;
      }

      .weekly-root {
        gap: 14px;
      }

      .weekly-toolbar {
        align-items: start;
      }

      .report-grid-2,
      .report-grid-3,
      .grid-2,
      .grid-3,
      .two-col {
        grid-template-columns: 1fr !important;
      }

      .metric-strip {
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }

      .column-chart-wrap,
      .table-wrap {
        overflow: visible !important;
      }

      .column-chart-svg {
        min-width: 0;
      }

      th {
        position: static;
      }

      thead {
        display: table-header-group;
      }

      tr,
      svg,
      .column-chart-wrap,
      .table-wrap {
        break-inside: avoid;
        page-break-inside: avoid;
      }
    }
  </style>
</head>
<body>
  <script id="dashboard-data" type="application/json">__DATA_JSON__</script>

  <div class="page">
    <section class="hero">
      <div class="hero-grid">
        <div>
          <div class="eyebrow">Interactive Operational Dashboard</div>
          <h1>Filter by day, week, rig, field, category, rep, and more.</h1>
          <p>The report is still generated from the spreadsheet automatically, but now the HTML behaves like a small dashboard. Once opened in the browser, it recalculates KPIs, charts, rankings, and tables instantly as you apply filters.</p>
          <p><strong>Source file:</strong> __SOURCE_FILE__<br><strong>Generated at:</strong> __GENERATED_AT__</p>
        </div>

        <div class="hero-side">
          <h2>Quick Period</h2>
          <div class="preset-grid">
            <button class="preset-btn" data-preset="all">All time</button>
            <button class="preset-btn" data-preset="last7">Last 7 days</button>
            <button class="preset-btn" data-preset="last30">Last 30 days</button>
            <button class="preset-btn" data-preset="last90">Last 90 days</button>
          </div>
          <p class="panel-note" style="color: rgba(255,255,255,0.8); margin-top: 14px;">The quick range buttons use the latest date found in the spreadsheet as the reference point, which makes them consistent even for historical datasets.</p>
          <button id="reset-filters" class="ghost-btn" style="width: 100%;">Reset all filters</button>
          <div class="theme-toggle-wrap">
            <div class="theme-toggle-row">
              <span class="theme-toggle-label">Theme</span>
              <span id="theme-toggle-state" class="theme-toggle-state">Classic</span>
            </div>
            <button id="theme-toggle" class="theme-toggle" type="button" aria-label="Toggle dashboard theme">
              <span id="theme-toggle-thumb" class="theme-toggle-thumb">Classic</span>
            </button>
          </div>
        </div>
      </div>
    </section>

    <div class="view-tabs" role="tablist" aria-label="Report Views">
      <button class="view-tab is-active" data-view="dashboard-view">Interactive Dashboard</button>
      <button class="view-tab" data-view="weekly-report-view">Weekly Report</button>
      <button class="view-tab" data-view="flat-time-view">Flat Time</button>
    </div>

    <div id="dashboard-view" class="view-panel">
      <div class="shell">
      <aside class="sidebar">
        <section class="panel">
          <div class="panel-inner">
            <h2>Filters</h2>
            <p class="panel-note">Use the controls below to focus the report. All cards, charts, rankings, and detailed tables update together.</p>
            <div class="filter-grid">
              <div class="field">
                <label for="start-date">Start Date</label>
                <input id="start-date" type="date">
              </div>
              <div class="field">
                <label for="end-date">End Date</label>
                <input id="end-date" type="date">
              </div>
              <div class="field">
                <label for="week-filter">Week</label>
                <select id="week-filter"></select>
              </div>
              <div class="field">
                <label for="month-filter">Month</label>
                <select id="month-filter"></select>
              </div>
              <div class="field">
                <label for="rig-filter">Rig</label>
                <select id="rig-filter"></select>
              </div>
              <div class="field">
                <label for="field-filter">Field</label>
                <select id="field-filter"></select>
              </div>
              <div class="field">
                <label for="well-filter">Well</label>
                <select id="well-filter"></select>
              </div>
              <div class="field">
                <label for="category-filter">Category</label>
                <select id="category-filter"></select>
              </div>
              <div class="field">
                <label for="type-filter">Type</label>
                <select id="type-filter"></select>
              </div>
              <div class="field">
                <label for="app-filter">Corva App</label>
                <select id="app-filter"></select>
              </div>
              <div class="field">
                <label for="rep-filter">RTES Rep</label>
                <select id="rep-filter"></select>
              </div>
              <div class="field">
                <label for="validation-filter">Validation</label>
                <select id="validation-filter">
                  <option value="">All validations</option>
                  <option value="validated">Validated only</option>
                  <option value="not_validated">Not validated only</option>
                </select>
              </div>
              <div class="field">
                <label for="granularity-filter">Trend Granularity</label>
                <select id="granularity-filter">
                  <option value="day">Daily</option>
                  <option value="week">Weekly</option>
                  <option value="month">Monthly</option>
                </select>
              </div>
              <div class="field">
                <label for="search-filter">Search</label>
                <input id="search-filter" type="search" placeholder="Rig, well, description, comment...">
              </div>
            </div>
          </div>
        </section>

        <section class="panel">
          <div class="panel-inner">
            <h2>What To Show</h2>
            <p class="panel-note">You asked for control over what appears on screen, so each main block can be shown or hidden independently.</p>
            <div class="toggle-grid">
              <label class="toggle"><input type="checkbox" data-target="summary-section" checked> Summary</label>
              <label class="toggle"><input type="checkbox" data-target="charts-section" checked> Charts</label>
              <label class="toggle"><input type="checkbox" data-target="rankings-section" checked> Rankings</label>
              <label class="toggle"><input type="checkbox" data-target="details-section" checked> Details</label>
              <label class="toggle"><input type="checkbox" data-target="ca-section" checked> RTES CA</label>
              <label class="toggle"><input type="checkbox" data-target="notes-section" checked> Notes</label>
            </div>
          </div>
        </section>
      </aside>

      <main class="main">
        <section id="summary-section" class="panel section">
          <div class="section-header">
            <div>
              <h2 class="section-title">Filtered Executive Summary</h2>
              <p class="section-subtitle">The KPI cards below reflect only the records that match the current filter set. This makes the same HTML useful for daily drill-downs, weekly summaries, and focused rig reviews.</p>
            </div>
            <div class="status-box">
              <strong id="results-title">Current result set</strong>
              <span id="results-subtitle">Preparing dashboard...</span>
            </div>
          </div>
          <div id="active-filters" class="chips"></div>
          <div id="kpi-grid" class="kpis" style="margin-top: 16px;"></div>
        </section>

        <section id="charts-section" class="panel section">
          <div class="section-header">
            <div>
              <h2 class="section-title">Interactive Charts</h2>
              <p class="section-subtitle">Trend and ranking charts update immediately with every filter change. Use the trend granularity control on the left to switch between daily, weekly, and monthly views.</p>
            </div>
          </div>
          <div class="two-col">
            <div class="chart">
              <h3>Intervention Trend</h3>
              <div id="trend-chart" class="line-chart"></div>
            </div>
            <div class="chart">
              <h3>Interventions by Category</h3>
              <div id="category-chart" class="bar-chart"></div>
            </div>
          </div>
          <div class="grid-3" style="margin-top: 18px;">
            <div class="chart">
              <h3>Top Rigs</h3>
              <div id="rig-chart" class="bar-chart"></div>
            </div>
            <div class="chart">
              <h3>Interventions by Type</h3>
              <div id="type-chart" class="bar-chart"></div>
            </div>
            <div class="chart">
              <h3>Corva Apps</h3>
              <div id="app-chart" class="bar-chart"></div>
            </div>
          </div>
        </section>

        <section id="rankings-section" class="panel section">
          <div class="section-header">
            <div>
              <h2 class="section-title">Rankings And Breakdowns</h2>
              <p class="section-subtitle">These summary tables help compare the currently selected slice of the operation without going back to Excel.</p>
            </div>
          </div>
          <div class="grid-2">
            <div>
              <h3>Top Categories</h3>
              <div id="category-table"></div>
            </div>
            <div>
              <h3>Top RTES Reps</h3>
              <div id="rep-table"></div>
            </div>
            <div>
              <h3>Top Fields</h3>
              <div id="field-table"></div>
            </div>
            <div>
              <h3>Top Wells</h3>
              <div id="well-table"></div>
            </div>
          </div>
        </section>

        <section id="details-section" class="panel section">
          <div class="section-header">
            <div>
              <h2 class="section-title">Filtered Intervention Details</h2>
              <p class="section-subtitle">This table shows the records behind the cards and charts so the report stays auditable. It is intentionally capped for browser performance, while still reflecting the filtered result set.</p>
            </div>
          </div>
          <div id="intervention-table"></div>
        </section>

        <section id="ca-section" class="panel section">
          <div class="section-header">
            <div>
              <h2 class="section-title">RTES Cost Avoidance</h2>
              <p class="section-subtitle">Rows from the RTES CA sheet are also filtered where possible, especially by period, rig, and well. This lets the financial view travel with the operational filter context.</p>
            </div>
          </div>
          <div class="grid-2">
            <div class="chart">
              <h3>Cost Avoidance by Rig</h3>
              <div id="ca-chart" class="bar-chart"></div>
            </div>
            <div id="ca-table"></div>
          </div>
        </section>

        <section id="notes-section" class="panel section">
          <div class="section-header">
            <div>
              <h2 class="section-title">Automation Notes</h2>
              <p class="section-subtitle">This interactive HTML is still static as a file, which means it remains easy to automate. The generator bakes the data and the JavaScript into one output file.</p>
            </div>
          </div>
          <p class="footer-note">Next automation step: keep the spreadsheet in a known folder, run <code>python3 generate_report.py</code> on a schedule, and optionally publish the resulting HTML to a shared drive or export it to PDF for distribution.</p>
        </section>
      </main>
      </div>
    </div>

    <div id="weekly-report-view" class="view-panel" hidden>
      <div class="weekly-root">
        <section class="panel section">
          <div class="section-header">
            <div>
              <h2 class="section-title">Weekly Report</h2>
              <p class="section-subtitle">This second tab keeps the current dashboard intact and reproduces the structure of the attached weekly Excel report inside the same HTML. It focuses on the selected date range while also showing cumulative weekly context.</p>
            </div>
            <div class="status-box">
              <strong id="weekly-report-title">Preparing weekly report...</strong>
              <span id="weekly-report-subtitle">Calculating weekly aggregates.</span>
            </div>
          </div>

          <div class="weekly-toolbar">
            <div style="display:flex; gap:16px; flex-wrap:wrap;">
              <div class="field" style="min-width: 180px;">
                <label for="weekly-report-start-date">Start Date</label>
                <input id="weekly-report-start-date" type="date">
              </div>
              <div class="field" style="min-width: 180px;">
                <label for="weekly-report-end-date">End Date</label>
                <input id="weekly-report-end-date" type="date">
              </div>
              <div class="field" style="min-width: 220px;">
                <label>&nbsp;</label>
                <button id="weekly-export-pdf" class="action-btn" type="button">Export Weekly Report PDF</button>
              </div>
            </div>
            <div class="weekly-toolbar-actions">
              <div class="chips">
                <span id="weekly-report-range" class="chip">Range pending</span>
                <span class="chip">Provider column uses available source fields when contractor data is missing.</span>
              </div>
            </div>
          </div>
        </section>

        <section class="panel section">
          <div class="weekly-banner">
            <div>
              <h2>RTES Weekly Report</h2>
              <p id="weekly-banner-copy">Styled after the attached Excel workbook, while remaining interactive and browser-friendly.</p>
            </div>
            <div class="chips">
              <span class="chip" id="weekly-banner-chip-1">Week summary</span>
              <span class="chip" id="weekly-banner-chip-2">Highlights</span>
              <span class="chip" id="weekly-banner-chip-3">Historical context</span>
            </div>
          </div>
        </section>

        <section class="panel section">
          <div class="report-grid-2">
            <div class="report-card">
              <h3>Weekly Category of Interventions and Validity</h3>
              <p class="report-note">Equivalent to the Excel weekly category summary for the selected period.</p>
              <div class="legend">
                <span class="legend-item"><span class="legend-dot" style="background:#1264d6;"></span>Interventions</span>
                <span class="legend-item"><span class="legend-dot" style="background:#0f766e;"></span>Rig Action</span>
                <span class="legend-item"><span class="legend-dot" style="background:#c06a0a;"></span>Validation %</span>
              </div>
              <div id="weekly-category-chart"></div>
              <div id="weekly-category-table"></div>
            </div>
            <div class="report-card">
              <h3>Cumulative Category of Interventions and Validity</h3>
              <p class="report-note">Cumulative values are calculated up to the end of the selected week, so you can move backwards historically and still get period-correct totals.</p>
              <div class="legend">
                <span class="legend-item"><span class="legend-dot" style="background:#1264d6;"></span>Interventions</span>
                <span class="legend-item"><span class="legend-dot" style="background:#0f766e;"></span>Rig Action</span>
                <span class="legend-item"><span class="legend-dot" style="background:#be123c;"></span>Validated</span>
                <span class="legend-item"><span class="legend-dot" style="background:#c06a0a;"></span>Validation %</span>
              </div>
              <div id="cumulative-category-chart"></div>
              <div id="cumulative-category-table"></div>
            </div>
          </div>
        </section>

        <section class="panel section">
          <div class="section-header">
            <div>
              <h2 class="section-title">Intervention Summary Blocks</h2>
              <p class="section-subtitle">This area mirrors the three Excel summary blocks: Wiper Trip, ROP, and KPI. The logic is derived from the intervention log using text and field matching, then grouped by rig and well for the selected week.</p>
            </div>
          </div>
          <div class="report-grid-3">
            <div class="report-card">
              <h3>Wiper Trip Interventions Summary</h3>
              <div class="legend">
                <span class="legend-item"><span class="legend-dot" style="background:#1264d6;"></span>Saved Time</span>
                <span class="legend-item"><span class="legend-dot" style="background:#c81e5a;"></span>Loss Time</span>
              </div>
              <div id="wiper-summary-chart"></div>
              <div id="wiper-summary-table"></div>
            </div>
            <div class="report-card">
              <h3>ROP Interventions Summary</h3>
              <div class="legend">
                <span class="legend-item"><span class="legend-dot" style="background:#1264d6;"></span>Saved Time</span>
                <span class="legend-item"><span class="legend-dot" style="background:#c81e5a;"></span>Loss Time</span>
              </div>
              <div id="rop-summary-chart"></div>
              <div id="rop-summary-table"></div>
            </div>
            <div class="report-card">
              <h3>KPI Interventions Summary</h3>
              <div class="legend">
                <span class="legend-item"><span class="legend-dot" style="background:#1264d6;"></span>Saved Time</span>
                <span class="legend-item"><span class="legend-dot" style="background:#c81e5a;"></span>Loss Time</span>
              </div>
              <div id="kpi-summary-chart"></div>
              <div id="kpi-summary-table"></div>
            </div>
          </div>
        </section>

        <section class="panel section">
          <div class="section-header">
            <div>
              <h2 class="section-title">Weekly Highlights</h2>
              <p class="section-subtitle">These tables follow the same split used in the workbook: realized savings on the left and potential savings or avoidance on the right.</p>
            </div>
          </div>
          <div id="weekly-highlight-metrics" class="metric-strip" style="margin-bottom: 18px;"></div>
          <div class="report-grid-2">
            <div class="report-card">
              <h3>Saved Time / Cost Saving</h3>
              <div id="actual-highlights-table"></div>
            </div>
            <div class="report-card">
              <h3>Potential Saved Time / Cost Avoidance</h3>
              <div id="potential-highlights-table"></div>
            </div>
          </div>
        </section>

        <section class="panel section">
          <div class="section-header">
            <div>
              <h2 class="section-title">Weekly Intervention Statistics and Analysis</h2>
              <p class="section-subtitle">Rig and well level intervention counts and validation rates by category for the selected weekly period.</p>
            </div>
          </div>
          <div id="weekly-stats-metrics" class="metric-strip" style="margin-bottom: 18px;"></div>
          <div id="weekly-stats-table"></div>
        </section>

      </div>
    </div>

    <div id="flat-time-view" class="view-panel" hidden>
      <div class="flat-time-root">
        <section class="panel section">
          <div class="section-header">
            <div>
              <h2 class="section-title">Flat Time</h2>
              <p class="section-subtitle">Compare flat time benchmark files to identify which activities and groups are consuming more time and where procedural improvements can reduce the total well duration.</p>
            </div>
            <div class="status-box">
              <strong id="flat-time-title">Preparing flat time comparison...</strong>
              <span id="flat-time-subtitle">Loading benchmark datasets.</span>
            </div>
          </div>

          <div class="flat-time-toolbar">
            <div style="display:flex; gap:16px; flex-wrap:wrap;">
              <div class="field" style="min-width: 190px;">
                <label for="flat-time-rig">Rig</label>
                <select id="flat-time-rig">
                  <option value="">All rigs</option>
                </select>
              </div>
              <div class="field" style="min-width: 190px;">
                <label for="flat-time-section">Section Size</label>
                <select id="flat-time-section">
                  <option value="">All section sizes</option>
                </select>
              </div>
              <div class="field" style="min-width: 190px;">
                <label for="flat-time-metric">Comparison Metric</label>
                <select id="flat-time-metric">
                  <option value="subject">Subject Well Time</option>
                  <option value="mean">Mean Time</option>
                  <option value="median">Median Time</option>
                </select>
              </div>
              <div class="field" style="min-width: 190px;">
                <label for="flat-time-top-n">Top Activities</label>
                <select id="flat-time-top-n">
                  <option value="8">Top 8</option>
                  <option value="10" selected>Top 10</option>
                  <option value="15">Top 15</option>
                </select>
              </div>
              <div class="field" style="min-width: 190px;">
                <label for="flat-time-mode">Analysis Mode</label>
                <select id="flat-time-mode">
                  <option value="executive" selected>Executive</option>
                  <option value="engineering">Engineering</option>
                </select>
              </div>
              <div class="field" style="min-width: 220px;">
                <label for="flat-time-well">Selected Well</label>
                <select id="flat-time-well">
                  <option value="">Auto-select worst well</option>
                </select>
              </div>
            </div>
            <div class="flat-time-upload-panel">
              <div class="field" style="min-width: 320px;">
                <label for="flat-time-upload">Add More CSV Files</label>
                <input id="flat-time-upload" class="file-input" type="file" accept=".csv,text/csv" multiple>
              </div>
              <button id="flat-time-recalculate" class="action-btn" type="button">Recalculate</button>
              <button id="flat-time-clear-uploads" class="action-btn" type="button">Clear All CSVs</button>
            </div>
          </div>

          <div id="flat-time-dataset-tags" class="tag-list"></div>
        </section>

        <section class="panel section" data-flat-mode="executive engineering">
          <div id="flat-time-summary" class="metric-strip"></div>
        </section>

        <section class="panel section" data-flat-mode="executive engineering">
          <div class="flat-time-chart-stack">
            <div class="report-card">
              <h3>Well Ranking by Excess Time</h3>
              <p class="report-note">Ranks wells by time above the recommended ideal, highlighting the main activity and group driving the excess.</p>
              <div id="flat-time-well-ranking"></div>
            </div>
            <div class="report-card">
              <h3>Pareto of Recoverable Hours</h3>
              <p class="report-note">Shows which activities concentrate the largest recoverable hours so improvement effort can be focused where it matters most.</p>
              <div id="flat-time-pareto-chart"></div>
            </div>
          </div>
        </section>

        <section class="panel section" data-flat-mode="executive engineering">
          <div class="flat-time-chart-stack">
            <div class="report-card">
              <h3>Well vs Ideal Waterfall</h3>
              <p class="report-note">Starts from the selected well actual flat time, subtracts the biggest activity gaps and lands on the recommended ideal total.</p>
              <div id="flat-time-waterfall-chart"></div>
            </div>
            <div class="report-card">
              <h3>Section Benchmark Chart</h3>
              <p class="report-note">Compares each section average against the recommended ideal and shows the spread to recover.</p>
              <div id="flat-time-section-benchmark-chart"></div>
            </div>
          </div>
        </section>

        <section class="panel section" data-flat-mode="executive engineering">
          <div class="flat-time-chart-stack">
            <div class="report-card">
              <h3>Rig Benchmark Summary</h3>
              <p class="report-note">Summarizes rig-level average flat time, ideal target, excess time and the main repeating activity.</p>
              <div id="flat-time-rig-summary"></div>
            </div>
            <div class="report-card">
              <h3>Opportunity Pipeline</h3>
              <p class="report-note">Executive list of the activities with the highest recoverable hours and the strongest case for procedural action.</p>
              <div id="flat-time-opportunity-pipeline"></div>
            </div>
          </div>
        </section>

        <section class="panel section" data-flat-mode="engineering">
          <div class="flat-time-chart-stack">
            <div class="report-card">
              <h3>Group Comparison</h3>
              <p class="report-note">Compare the largest flat time group totals across all uploaded benchmark files.</p>
              <div id="flat-time-group-chart"></div>
            </div>
            <div class="report-card">
              <h3>Top Activities By Time</h3>
              <p class="report-note">Highlight the activities that are consuming the most time across the comparison set.</p>
              <div id="flat-time-activity-chart"></div>
            </div>
          </div>
        </section>

        <section class="panel section" data-flat-mode="executive engineering">
          <div class="report-card">
            <h3>Activity Benchmark Table</h3>
            <p class="report-note">Engineering view of each activity, including sample size, distribution, recommended ideal time, variability and recoverable hours.</p>
            <div id="flat-time-benchmark-table"></div>
          </div>
        </section>

        <section class="panel section" data-flat-mode="executive engineering">
          <div class="report-card">
            <h3>Drill-down Explorer</h3>
            <p class="report-note" id="flat-time-drilldown-note">Click a well or activity in the tables above to open the benchmark, peer comparison and ideal-time logic.</p>
            <div class="drill-grid">
              <div>
                <h3 style="font-size:16px; margin-bottom:10px;">Selected Well Breakdown</h3>
                <div id="flat-time-well-drilldown"></div>
              </div>
              <div>
                <h3 style="font-size:16px; margin-bottom:10px;">Selected Activity Benchmark</h3>
                <div id="flat-time-activity-drilldown"></div>
              </div>
            </div>
          </div>
        </section>

        <section class="panel section" data-flat-mode="engineering">
          <div class="flat-time-chart-stack">
            <div class="report-card">
              <h3>Reduction Opportunity Matrix</h3>
              <p class="report-note">Gap is calculated against the ideal achievable time for the activity. Example: if 4 wells ran near 30 hr and 1 well ran 45 hr, the gap shown is 15 hr.</p>
              <div id="flat-time-opportunity-table"></div>
            </div>
            <div class="report-card">
              <h3>Group Totals By Dataset</h3>
              <p class="report-note">Friendly comparison table showing how each benchmark dataset is distributed by flat time group.</p>
              <div id="flat-time-group-table"></div>
            </div>
          </div>
        </section>

        <section class="panel section" data-flat-mode="engineering">
          <div class="report-grid-2">
            <div class="report-card">
              <h3>Top Loss Drivers by Well</h3>
              <p class="report-note">Lists the top three activities that explain the excess time for each well.</p>
              <div id="flat-time-loss-drivers"></div>
            </div>
            <div class="report-card">
              <h3>Variability Box Plot</h3>
              <p class="report-note">Shows spread by activity so unstable work can be separated from predictable, repeatable work.</p>
              <div id="flat-time-variability-chart"></div>
            </div>
          </div>
        </section>

        <section class="panel section" data-flat-mode="engineering">
          <div class="report-card">
            <h3>Activity Heatmap</h3>
            <p class="report-note">Heatmap of hours above ideal by well and activity, making repeated loss patterns visible at a glance.</p>
            <div id="flat-time-heatmap"></div>
          </div>
        </section>

        <section class="panel section" data-flat-mode="executive engineering">
          <div class="report-card">
            <h3>Perfect Flat Time vs Wells</h3>
            <p class="report-note">Compares the ideal cumulative flat time curve against the actual wells. Days are cumulative on X and depth progression is represented by section sequence on Y.</p>
            <div id="flat-time-perfect-chart"></div>
          </div>
        </section>
      </div>
    </div>
  </div>

  <script>
    const dashboardData = JSON.parse(document.getElementById("dashboard-data").textContent);

    const ui = {
      startDate: document.getElementById("start-date"),
      endDate: document.getElementById("end-date"),
      week: document.getElementById("week-filter"),
      month: document.getElementById("month-filter"),
      rig: document.getElementById("rig-filter"),
      field: document.getElementById("field-filter"),
      well: document.getElementById("well-filter"),
      category: document.getElementById("category-filter"),
      type: document.getElementById("type-filter"),
      app: document.getElementById("app-filter"),
      rep: document.getElementById("rep-filter"),
      validation: document.getElementById("validation-filter"),
      granularity: document.getElementById("granularity-filter"),
      search: document.getElementById("search-filter"),
      reset: document.getElementById("reset-filters"),
      themeToggle: document.getElementById("theme-toggle"),
      themeToggleState: document.getElementById("theme-toggle-state"),
      themeToggleThumb: document.getElementById("theme-toggle-thumb"),
      resultsTitle: document.getElementById("results-title"),
      resultsSubtitle: document.getElementById("results-subtitle"),
      activeFilters: document.getElementById("active-filters"),
      kpiGrid: document.getElementById("kpi-grid"),
      trendChart: document.getElementById("trend-chart"),
      categoryChart: document.getElementById("category-chart"),
      rigChart: document.getElementById("rig-chart"),
      typeChart: document.getElementById("type-chart"),
      appChart: document.getElementById("app-chart"),
      categoryTable: document.getElementById("category-table"),
      repTable: document.getElementById("rep-table"),
      fieldTable: document.getElementById("field-table"),
      wellTable: document.getElementById("well-table"),
      interventionTable: document.getElementById("intervention-table"),
      caChart: document.getElementById("ca-chart"),
      caTable: document.getElementById("ca-table"),
      presetButtons: Array.from(document.querySelectorAll(".preset-btn")),
      toggles: Array.from(document.querySelectorAll("[data-target]")),
      viewTabs: Array.from(document.querySelectorAll(".view-tab")),
      viewPanels: Array.from(document.querySelectorAll(".view-panel")),
      weeklyReportStartDate: document.getElementById("weekly-report-start-date"),
      weeklyReportEndDate: document.getElementById("weekly-report-end-date"),
      weeklyExportPdf: document.getElementById("weekly-export-pdf"),
      weeklyReportTitle: document.getElementById("weekly-report-title"),
      weeklyReportSubtitle: document.getElementById("weekly-report-subtitle"),
      weeklyReportRange: document.getElementById("weekly-report-range"),
      weeklyBannerCopy: document.getElementById("weekly-banner-copy"),
      weeklyBannerChip1: document.getElementById("weekly-banner-chip-1"),
      weeklyBannerChip2: document.getElementById("weekly-banner-chip-2"),
      weeklyBannerChip3: document.getElementById("weekly-banner-chip-3"),
      weeklyCategoryTable: document.getElementById("weekly-category-table"),
      weeklyCategoryChart: document.getElementById("weekly-category-chart"),
      cumulativeCategoryTable: document.getElementById("cumulative-category-table"),
      cumulativeCategoryChart: document.getElementById("cumulative-category-chart"),
      wiperSummaryTable: document.getElementById("wiper-summary-table"),
      wiperSummaryChart: document.getElementById("wiper-summary-chart"),
      ropSummaryTable: document.getElementById("rop-summary-table"),
      ropSummaryChart: document.getElementById("rop-summary-chart"),
      kpiSummaryTable: document.getElementById("kpi-summary-table"),
      kpiSummaryChart: document.getElementById("kpi-summary-chart"),
      weeklyHighlightMetrics: document.getElementById("weekly-highlight-metrics"),
      actualHighlightsTable: document.getElementById("actual-highlights-table"),
      potentialHighlightsTable: document.getElementById("potential-highlights-table"),
      weeklyStatsMetrics: document.getElementById("weekly-stats-metrics"),
      weeklyStatsTable: document.getElementById("weekly-stats-table"),
      flatTimeTitle: document.getElementById("flat-time-title"),
      flatTimeSubtitle: document.getElementById("flat-time-subtitle"),
      flatTimeRig: document.getElementById("flat-time-rig"),
      flatTimeSection: document.getElementById("flat-time-section"),
      flatTimeMetric: document.getElementById("flat-time-metric"),
      flatTimeTopN: document.getElementById("flat-time-top-n"),
      flatTimeMode: document.getElementById("flat-time-mode"),
      flatTimeWell: document.getElementById("flat-time-well"),
      flatTimeUpload: document.getElementById("flat-time-upload"),
      flatTimeRecalculate: document.getElementById("flat-time-recalculate"),
      flatTimeClearUploads: document.getElementById("flat-time-clear-uploads"),
      flatTimeDatasetTags: document.getElementById("flat-time-dataset-tags"),
      flatTimeSummary: document.getElementById("flat-time-summary"),
      flatTimeWellRanking: document.getElementById("flat-time-well-ranking"),
      flatTimeParetoChart: document.getElementById("flat-time-pareto-chart"),
      flatTimeWaterfallChart: document.getElementById("flat-time-waterfall-chart"),
      flatTimeSectionBenchmarkChart: document.getElementById("flat-time-section-benchmark-chart"),
      flatTimeRigSummary: document.getElementById("flat-time-rig-summary"),
      flatTimeOpportunityPipeline: document.getElementById("flat-time-opportunity-pipeline"),
      flatTimeGroupChart: document.getElementById("flat-time-group-chart"),
      flatTimeActivityChart: document.getElementById("flat-time-activity-chart"),
      flatTimeBenchmarkTable: document.getElementById("flat-time-benchmark-table"),
      flatTimeDrilldownNote: document.getElementById("flat-time-drilldown-note"),
      flatTimeWellDrilldown: document.getElementById("flat-time-well-drilldown"),
      flatTimeActivityDrilldown: document.getElementById("flat-time-activity-drilldown"),
      flatTimeOpportunityTable: document.getElementById("flat-time-opportunity-table"),
      flatTimeGroupTable: document.getElementById("flat-time-group-table"),
      flatTimeLossDrivers: document.getElementById("flat-time-loss-drivers"),
      flatTimeVariabilityChart: document.getElementById("flat-time-variability-chart"),
      flatTimeHeatmap: document.getElementById("flat-time-heatmap"),
      flatTimePerfectChart: document.getElementById("flat-time-perfect-chart"),
      flatTimeModeSections: Array.from(document.querySelectorAll("#flat-time-view [data-flat-mode]")),
    };

    const currencyFormatter = new Intl.NumberFormat("en-US", {
      style: "currency",
      currency: "USD",
      maximumFractionDigits: 0,
    });

    const numberFormatter = new Intl.NumberFormat("en-US", {
      maximumFractionDigits: 1,
      minimumFractionDigits: 0,
    });

    const CATEGORY_ORDER = ["Stuck pipe", "Optimization", "Operational Compliance", "Well Control", "Reporting"];
    const THEME_STORAGE_KEY = "weekly-report-theme";
    const FLAT_TIME_SERIES_COLORS = ["#1264d6", "#0f766e", "#c06a0a", "#be123c", "#7c3aed", "#0891b2", "#16a34a", "#dc2626"];
    const FLAT_TIME_ACTIVITY_TRANSLATIONS = dashboardData.activityCodeTranslations || {
      loaded: false,
      source: "",
      wellSections: {},
      operations: {},
      activities: {},
      generic: {},
    };
    const FLAT_TIME_RIG_LOOKUP = buildFlatTimeRigLookup(Array.isArray(dashboardData.interventions) ? dashboardData.interventions : []);
    const flatTimeState = {
      baseDatasets: [],
      uploadedDatasets: [],
      focusWell: "",
      focusActivity: "",
    };

    function getChartTheme() {
      const isCorona = document.body.classList.contains("theme-corona");
      return isCorona
        ? {
            text: "#f5f5f5",
            muted: "#a1aab8",
            grid: "#2c2e33",
            axis: "#3a3d46",
            line: "#0090e7",
            area: "rgba(0, 144, 231, 0.16)",
            pointLabel: "#f5f5f5",
            valueLabel: "#d5d9e0",
          }
        : {
            text: "#1f2d3d",
            muted: "#607085",
            grid: "#d8e2ef",
            axis: "#9fb3c8",
            line: "#1264d6",
            area: "rgba(18, 100, 214, 0.10)",
            pointLabel: "#34475d",
            valueLabel: "#6b7b8d",
          };
    }

    function applyTheme(theme) {
      const resolvedTheme = theme === "corona" ? "corona" : "classic";
      document.body.classList.toggle("theme-corona", resolvedTheme === "corona");
      ui.themeToggleState.textContent = resolvedTheme === "corona" ? "Corona" : "Classic";
      ui.themeToggleThumb.textContent = resolvedTheme === "corona" ? "Corona" : "Classic";
      localStorage.setItem(THEME_STORAGE_KEY, resolvedTheme);
    }

    function toggleTheme() {
      applyTheme(document.body.classList.contains("theme-corona") ? "classic" : "corona");
      applyFilters();
      renderWeeklyReport();
      renderFlatTime();
    }

    function escapeHtml(value) {
      return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#39;");
    }

    function formatCurrency(value) {
      return currencyFormatter.format(Number(value || 0));
    }

    function formatNumber(value) {
      return numberFormatter.format(Number(value || 0));
    }

    function slugify(value) {
      return String(value || "")
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, "-")
        .replace(/^-+|-+$/g, "");
    }

    function normalizeFlatTimeWellToken(value) {
      return String(value || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
    }

    function baseFlatTimeWellToken(value) {
      return normalizeFlatTimeWellToken(String(value || "").replace(/_\\d+$/, ""));
    }

    function buildFlatTimeRigLookup(rows) {
      const lookup = new Map();

      function addRig(key, rigName) {
        if (!key || !rigName) return;
        if (!lookup.has(key)) lookup.set(key, new Set());
        lookup.get(key).add(rigName);
      }

      rows.forEach((row) => {
        const wellName = row.wellName || "";
        const rigName = row.rigName || "";
        addRig(normalizeFlatTimeWellToken(wellName), rigName);
        addRig(baseFlatTimeWellToken(wellName), rigName);
      });

      return lookup;
    }

    function resolveFlatTimeRigLabel(subjectWell) {
      const exact = FLAT_TIME_RIG_LOOKUP.get(normalizeFlatTimeWellToken(subjectWell));
      if (exact && exact.size) return Array.from(exact).sort().join(" / ");

      const base = FLAT_TIME_RIG_LOOKUP.get(baseFlatTimeWellToken(subjectWell));
      if (base && base.size) return Array.from(base).sort().join(" / ");

      return "Rig not mapped";
    }

    function deriveFlatTimeRigLabelFromFileName(fileName) {
      const baseName = String(fileName || "").replace(/\\.[^.]+$/, "");
      const prefix = baseName.split("_")[0].trim();
      if (prefix) return prefix;
      return "Rig not mapped";
    }

    function enrichFlatTimeDataset(dataset) {
      if (!dataset) return dataset;
      return {
        ...dataset,
        rigLabel: dataset.rigLabel || resolveFlatTimeRigLabel(dataset.subjectWell),
      };
    }

    function parseCsvLine(line) {
      const values = [];
      let current = "";
      let inQuotes = false;
      for (let index = 0; index < line.length; index += 1) {
        const char = line[index];
        if (char === '"') {
          if (inQuotes && line[index + 1] === '"') {
            current += '"';
            index += 1;
          } else {
            inQuotes = !inQuotes;
          }
        } else if (char === "," && !inQuotes) {
          values.push(current.trim());
          current = "";
        } else {
          current += char;
        }
      }
      values.push(current.trim());
      return values;
    }

    function parseFlatTimeCsvText(fileName, text) {
      const rows = String(text || "")
        .replace(/\\r\\n/g, "\\n")
        .replace(/\\r/g, "\\n")
        .split("\\n")
        .map((line) => parseCsvLine(line));
      const rigLabel = deriveFlatTimeRigLabelFromFileName(fileName);
      const datasetMap = new Map();
      let currentGroupName = "";
      let currentHeader = null;

      function getDataset(subjectWell) {
        if (!datasetMap.has(subjectWell)) {
          datasetMap.set(subjectWell, {
            id: slugify(fileName + "-" + subjectWell),
            fileName,
            subjectWell,
            rigLabel,
            groupsMap: new Map(),
          });
        }
        return datasetMap.get(subjectWell);
      }

      function getGroup(dataset, groupName) {
        if (!dataset.groupsMap.has(groupName)) {
          dataset.groupsMap.set(groupName, {
            groupName,
            activities: [],
            totalSubjectHours: 0,
            totalMeanHours: 0,
            totalMedianHours: 0,
          });
        }
        return dataset.groupsMap.get(groupName);
      }

      rows.forEach((row) => {
        const first = (row[0] || "").trim();
        if (first === "Group Name") {
          currentGroupName = (row[1] || "Unknown").trim();
          currentHeader = null;
          return;
        }

        if (!currentGroupName) return;
        if (!first || first === "Group Type") return;

        if (first === "Activity") {
          const wellColumns = [];
          let meanIndex = -1;
          let medianIndex = -1;

          row.forEach((cell, index) => {
            const label = String(cell || "").trim();
            if (index === 0 || !label) return;
            if (/^mean/i.test(label)) {
              meanIndex = index;
              return;
            }
            if (/^median/i.test(label)) {
              medianIndex = index;
              return;
            }
            wellColumns.push({ index, label });
          });

          currentHeader = { wellColumns, meanIndex, medianIndex };
          return;
        }

        if (!currentHeader || !currentHeader.wellColumns.length) return;

        if (first === "Total") {
          currentHeader.wellColumns.forEach((column) => {
            const dataset = getDataset(column.label);
            const group = getGroup(dataset, currentGroupName);
            group.totalSubjectHours = Number(row[column.index] || group.totalSubjectHours || 0);
          });
          return;
        }

        currentHeader.wellColumns.forEach((column) => {
          const subjectHours = Number(row[column.index] || 0);
          if (!subjectHours) return;
          const dataset = getDataset(column.label);
          const group = getGroup(dataset, currentGroupName);
          group.activities.push({
            activity: first,
            sectionSize: extractFlatTimeSectionSize(first),
            subjectHours,
            meanHours: currentHeader.meanIndex >= 0 ? Number(row[currentHeader.meanIndex] || 0) : 0,
            medianHours: currentHeader.medianIndex >= 0 ? Number(row[currentHeader.medianIndex] || 0) : 0,
          });
        });
      });

      return Array.from(datasetMap.values()).map((dataset) => {
        const groups = Array.from(dataset.groupsMap.values()).filter((group) => group.activities.length || group.totalSubjectHours);
        groups.forEach((group) => {
          if (!group.totalSubjectHours) {
            group.totalSubjectHours = group.activities.reduce((sum, item) => sum + item.subjectHours, 0);
          }
          // Recompute aggregate benchmarks from activity rows because several CSV
          // exports carry inconsistent group total mean/median values.
          group.totalMeanHours = group.activities.reduce((sum, item) => sum + item.meanHours, 0);
          group.totalMedianHours = group.activities.reduce((sum, item) => sum + item.medianHours, 0);
        });

        return {
          id: dataset.id,
          fileName: dataset.fileName,
          subjectWell: dataset.subjectWell,
          rigLabel: dataset.rigLabel,
          groups,
          totalSubjectHours: groups.reduce((sum, group) => sum + group.totalSubjectHours, 0),
          totalMeanHours: groups.reduce((sum, group) => sum + group.totalMeanHours, 0),
          totalMedianHours: groups.reduce((sum, group) => sum + group.totalMedianHours, 0),
        };
      });
    }

    function getFlatTimeDatasets() {
      return [...flatTimeState.baseDatasets, ...flatTimeState.uploadedDatasets].map(enrichFlatTimeDataset);
    }

    function getAvailableFlatTimeRigs(datasets) {
      return Array.from(new Set(datasets.map((dataset) => dataset.rigLabel || "Rig not mapped"))).sort((left, right) => left.localeCompare(right));
    }

    function populateFlatTimeRigOptions(datasets) {
      const current = ui.flatTimeRig.value;
      const options = ['<option value="">All rigs</option>'];
      const rigs = getAvailableFlatTimeRigs(datasets);
      rigs.forEach((rig) => {
        options.push('<option value="' + escapeHtml(rig) + '">' + escapeHtml(rig) + "</option>");
      });
      ui.flatTimeRig.innerHTML = options.join("");
      if (rigs.includes(current)) {
        ui.flatTimeRig.value = current;
      }
    }

    function filterFlatTimeDatasetsByRig(datasets, rigLabel) {
      if (!rigLabel) return datasets;
      return datasets.filter((dataset) => (dataset.rigLabel || "Rig not mapped") === rigLabel);
    }

    function getFlatTimeMetricKey() {
      const metric = ui.flatTimeMetric.value || "subject";
      return metric === "mean" ? "meanHours" : metric === "median" ? "medianHours" : "subjectHours";
    }

    function getFlatTimeTotalKey() {
      const metric = ui.flatTimeMetric.value || "subject";
      return metric === "mean" ? "totalMeanHours" : metric === "median" ? "totalMedianHours" : "totalSubjectHours";
    }

    function getFlatTimeMode() {
      return ui.flatTimeMode.value || "executive";
    }

    function updateFlatTimeModeVisibility() {
      const mode = getFlatTimeMode();
      ui.flatTimeModeSections.forEach((section) => {
      const allowedModes = String(section.dataset.flatMode || "executive engineering").split(/\\s+/).filter(Boolean);
        section.hidden = !allowedModes.includes(mode);
      });
    }

    function populateFlatTimeWellOptions(datasets, preferredWell) {
      const current = ui.flatTimeWell.value;
      const options = ['<option value="">Auto-select worst well</option>'];
      const wellNames = datasets.map((dataset) => dataset.subjectWell);
      wellNames.forEach((wellName) => {
        const dataset = datasets.find((item) => item.subjectWell === wellName);
        const label = wellName + (dataset && dataset.rigLabel ? " • " + dataset.rigLabel : "");
        options.push('<option value="' + escapeHtml(wellName) + '">' + escapeHtml(label) + "</option>");
      });
      ui.flatTimeWell.innerHTML = options.join("");
      if (wellNames.includes(current)) {
        ui.flatTimeWell.value = current;
      } else if (preferredWell && wellNames.includes(preferredWell)) {
        ui.flatTimeWell.value = preferredWell;
      } else {
        ui.flatTimeWell.value = "";
      }
    }

    function createFlatTimeUploadId(fileName, subjectWell) {
      return slugify(fileName + "-" + subjectWell + "-" + Date.now() + "-" + Math.random().toString(36).slice(2, 8));
    }

    function extractFlatTimeSectionSize(activityName) {
      const match = String(activityName || "").match(/^(\\d+(?:\\.\\d+)?)(?=-)/);
      return match ? match[1] : "__no_section__";
    }

    function formatFlatTimeSectionSize(sectionSize) {
      return sectionSize === "__no_section__" ? "No section size" : sectionSize + '"';
    }

    function compareFlatTimeSectionSizes(left, right) {
      if (left === right) return 0;
      if (left === "__no_section__") return 1;
      if (right === "__no_section__") return -1;
      return Number(left) - Number(right) || left.localeCompare(right);
    }

    function getAvailableFlatTimeSectionSizes(datasets) {
      return Array.from(
        new Set(
          datasets.flatMap((dataset) =>
            dataset.groups.flatMap((group) =>
              group.activities.map((activity) => activity.sectionSize || extractFlatTimeSectionSize(activity.activity))
            )
          )
        )
      ).sort(compareFlatTimeSectionSizes);
    }

    function populateFlatTimeSectionOptions(datasets) {
      const current = ui.flatTimeSection.value;
      const options = ['<option value="">All section sizes</option>'];
      getAvailableFlatTimeSectionSizes(datasets).forEach((sectionSize) => {
        options.push(
          '<option value="' + escapeHtml(sectionSize) + '">' + escapeHtml(formatFlatTimeSectionSize(sectionSize)) + "</option>"
        );
      });
      ui.flatTimeSection.innerHTML = options.join("");
      const available = getAvailableFlatTimeSectionSizes(datasets);
      if (available.includes(current)) {
        ui.flatTimeSection.value = current;
      }
    }

    function filterFlatTimeDatasetsBySection(datasets, sectionSize) {
      if (!sectionSize) return datasets;

      return datasets
        .map((dataset) => {
          const groups = dataset.groups
            .map((group) => {
              const activities = group.activities.filter(
                (activity) => (activity.sectionSize || extractFlatTimeSectionSize(activity.activity)) === sectionSize
              );
              if (!activities.length) return null;
              return {
                groupName: group.groupName,
                activities,
                totalSubjectHours: activities.reduce((sum, activity) => sum + Number(activity.subjectHours || 0), 0),
                totalMeanHours: activities.reduce((sum, activity) => sum + Number(activity.meanHours || 0), 0),
                totalMedianHours: activities.reduce((sum, activity) => sum + Number(activity.medianHours || 0), 0),
              };
            })
            .filter(Boolean);

          if (!groups.length) return null;

          return {
            ...dataset,
            groups,
            totalSubjectHours: groups.reduce((sum, group) => sum + Number(group.totalSubjectHours || 0), 0),
            totalMeanHours: groups.reduce((sum, group) => sum + Number(group.totalMeanHours || 0), 0),
            totalMedianHours: groups.reduce((sum, group) => sum + Number(group.totalMedianHours || 0), 0),
          };
        })
        .filter(Boolean);
    }

    function annotateFlatTimeScopedBenchmarks(datasets) {
      if (!datasets.length) return datasets;

      const activityScopeMap = new Map();

      datasets.forEach((dataset) => {
        dataset.groups.forEach((group) => {
          group.activities.forEach((activity) => {
            const sectionSize = activity.sectionSize || extractFlatTimeSectionSize(activity.activity);
            const scopeKey = [sectionSize, group.groupName, activity.activity].join("||");
            const subjectHours = Number(activity.subjectHours || 0);
            if (subjectHours <= 0) return;
            if (!activityScopeMap.has(scopeKey)) activityScopeMap.set(scopeKey, []);
            activityScopeMap.get(scopeKey).push(subjectHours);
          });
        });
      });

      return datasets.map((dataset) => {
        const groups = dataset.groups.map((group) => {
          const activities = group.activities.map((activity) => {
            const sectionSize = activity.sectionSize || extractFlatTimeSectionSize(activity.activity);
            const scopeKey = [sectionSize, group.groupName, activity.activity].join("||");
            const values = activityScopeMap.get(scopeKey) || [];
            return {
              ...activity,
              meanHours: values.length ? average(values) : 0,
              medianHours: values.length ? percentile(values, 0.5) : 0,
            };
          });

          return {
            ...group,
            activities,
            totalSubjectHours: activities.reduce((sum, item) => sum + Number(item.subjectHours || 0), 0),
            totalMeanHours: activities.reduce((sum, item) => sum + Number(item.meanHours || 0), 0),
            totalMedianHours: activities.reduce((sum, item) => sum + Number(item.medianHours || 0), 0),
          };
        });

        return {
          ...dataset,
          groups,
          totalSubjectHours: groups.reduce((sum, group) => sum + Number(group.totalSubjectHours || 0), 0),
          totalMeanHours: groups.reduce((sum, group) => sum + Number(group.totalMeanHours || 0), 0),
          totalMedianHours: groups.reduce((sum, group) => sum + Number(group.totalMedianHours || 0), 0),
        };
      });
    }

    function buildPerfectFlatTimeSections(datasets) {
      const sectionMap = new Map();

      datasets.forEach((dataset) => {
        dataset.groups.forEach((group) => {
          group.activities.forEach((activity) => {
            const sectionSize = activity.sectionSize || extractFlatTimeSectionSize(activity.activity);
            if (sectionSize === "__no_section__") return;
            if (!sectionMap.has(sectionSize)) sectionMap.set(sectionSize, new Map());
            const activityMap = sectionMap.get(sectionSize);
            const value = Number(activity.subjectHours || 0);
            if (value <= 0) return;
            const current = activityMap.get(activity.activity);
            if (!current || value < current) activityMap.set(activity.activity, value);
          });
        });
      });

      return Array.from(sectionMap.entries())
        .map(([sectionSize, activityMap]) => ({
          sectionSize,
          bestHours: Array.from(activityMap.values()).reduce((sum, value) => sum + value, 0),
        }))
        .filter((item) => item.bestHours > 0)
        .sort((left, right) => Number(right.sectionSize) - Number(left.sectionSize));
    }

    function renderPerfectFlatTimeChart(target, datasets, metricKey) {
      const idealSections = buildPerfectFlatTimeSections(datasets);
      if (!idealSections.length) {
        target.innerHTML = '<div class="empty">No section-sized activities available to draw the perfect flat time curve.</div>';
        return;
      }

      const sectionOrder = Array.from(
        new Set(
          datasets.flatMap((dataset) =>
            dataset.groups.flatMap((group) =>
              group.activities
                .map((activity) => activity.sectionSize || extractFlatTimeSectionSize(activity.activity))
                .filter((sectionSize) => sectionSize && sectionSize !== "__no_section__")
            )
          )
        )
      ).sort((left, right) => Number(right) - Number(left) || left.localeCompare(right));

      const idealMap = new Map(idealSections.map((section) => [section.sectionSize, section.bestHours]));

      function buildSeries(label, color, sectionHours, isIdeal) {
        let cumulativeDays = 0;
        const points = sectionOrder.map((sectionSize, index) => {
          cumulativeDays += Number(sectionHours.get(sectionSize) || 0) / 24;
          return {
            sectionSize,
            cumulativeDays,
            depthIndex: index + 1,
          };
        });
        return { label, color, isIdeal, points };
      }

      const series = [
        buildSeries(
          "Ideal curve",
          "#1264d6",
          new Map(sectionOrder.map((sectionSize) => [sectionSize, Number(idealMap.get(sectionSize) || 0)])),
          true
        ),
        ...datasets.map((dataset, index) => {
          const sectionHours = new Map();
          dataset.groups.forEach((group) => {
            group.activities.forEach((activity) => {
              const sectionSize = activity.sectionSize || extractFlatTimeSectionSize(activity.activity);
              if (!sectionSize || sectionSize === "__no_section__") return;
              const value = Number(activity[metricKey] || 0);
              sectionHours.set(sectionSize, (sectionHours.get(sectionSize) || 0) + value);
            });
          });
          return buildSeries(
            dataset.subjectWell,
            FLAT_TIME_SERIES_COLORS[index % FLAT_TIME_SERIES_COLORS.length],
            sectionHours,
            false
          );
        }),
      ];

      const chartTheme = getChartTheme();
      const width = 960;
      const height = 430;
      const margin = { top: 30, right: 28, bottom: 52, left: 88 };
      const chartWidth = width - margin.left - margin.right;
      const chartHeight = height - margin.top - margin.bottom;
      const maxDays = Math.max(
        ...series.flatMap((line) => line.points.map((point) => point.cumulativeDays)),
        1
      );
      const maxDepth = Math.max(sectionOrder.length, 1);

      const scaledSeries = series.map((line) => ({
        ...line,
        points: line.points.map((point) => ({
          ...point,
          x: margin.left + (point.cumulativeDays / maxDays) * chartWidth,
          y: margin.top + ((point.depthIndex - 1) / Math.max(maxDepth - 1, 1)) * chartHeight,
        })),
      }));

      const xTicks = Array.from({ length: 6 }, (_, index) => {
        const value = (maxDays / 5) * index;
        const x = margin.left + (value / maxDays) * chartWidth;
        return (
          '<g>' +
          '<line x1="' + x.toFixed(2) + '" y1="' + margin.top + '" x2="' + x.toFixed(2) + '" y2="' + (height - margin.bottom) + '" stroke="' + chartTheme.grid + '" stroke-dasharray="4 5"></line>' +
          '<text x="' + x.toFixed(2) + '" y="' + (height - 12) + '" text-anchor="middle" font-size="11" fill="' + chartTheme.valueLabel + '">' + escapeHtml(formatNumber(value)) + "</text>" +
          "</g>"
        );
      }).join("");

      const yTicks = sectionOrder
        .map((sectionSize, index) => {
          const y = margin.top + (index / Math.max(maxDepth - 1, 1)) * chartHeight;
          return (
            '<g>' +
            '<line x1="' + margin.left + '" y1="' + y.toFixed(2) + '" x2="' + (width - margin.right) + '" y2="' + y.toFixed(2) + '" stroke="' + chartTheme.grid + '" stroke-dasharray="4 5"></line>' +
            '<text x="' + (margin.left - 12) + '" y="' + (y + 4).toFixed(2) + '" text-anchor="end" font-size="11" font-weight="700" fill="' + chartTheme.text + '">' + escapeHtml(formatFlatTimeSectionSize(sectionSize)) + "</text>" +
            '</g>'
          );
        })
        .join("");

      const lineSvg = scaledSeries
        .map((line) => {
          const path = line.points.map((point, index) => (index === 0 ? "M" : "L") + point.x.toFixed(2) + " " + point.y.toFixed(2)).join(" ");
          const endPoint = line.points[line.points.length - 1];
          const labelX = Math.min(width - margin.right + 4, endPoint.x + 10);
          const pointsSvg = line.points
            .map((point) => (
              '<circle cx="' + point.x.toFixed(2) + '" cy="' + point.y.toFixed(2) + '" r="' + (line.isIdeal ? "4.5" : "3.5") + '" fill="' + line.color + '" opacity="' + (line.isIdeal ? "1" : "0.85") + '"></circle>'
            ))
            .join("");
          return (
            '<g>' +
            '<path d="' + path + '" fill="none" stroke="' + line.color + '" stroke-width="' + (line.isIdeal ? "4.5" : "2.5") + '" stroke-linecap="round" stroke-linejoin="round" opacity="' + (line.isIdeal ? "1" : "0.9") + '"></path>' +
            pointsSvg +
            '<text x="' + labelX.toFixed(2) + '" y="' + (endPoint.y + (line.isIdeal ? -10 : 10)).toFixed(2) + '" font-size="11" font-weight="700" fill="' + chartTheme.pointLabel + '">' + escapeHtml(line.label + " • " + formatNumber(endPoint.cumulativeDays) + " d") + '</text>' +
            '</g>'
          );
        })
        .join("");

      const legend = series
        .map((line) => (
          '<span class="legend-item" style="margin-right:12px;">' +
          '<span class="legend-dot" style="background:' + line.color + '; width:14px; height:14px;"></span>' +
          escapeHtml(line.label) +
          '</span>'
        ))
        .join("");

      target.innerHTML =
        '<div class="legend" style="margin-bottom:12px; flex-wrap:wrap;">' + legend + '</div>' +
        '<div class="column-chart-wrap">' +
        '<svg class="column-chart-svg" viewBox="0 0 ' + width + " " + height + '" role="img" aria-label="Perfect flat time compared with wells">' +
        xTicks +
        yTicks +
        '<line x1="' + margin.left + '" y1="' + margin.top + '" x2="' + margin.left + '" y2="' + (height - margin.bottom) + '" stroke="' + chartTheme.axis + '"></line>' +
        '<line x1="' + margin.left + '" y1="' + (height - margin.bottom) + '" x2="' + (width - margin.right) + '" y2="' + (height - margin.bottom) + '" stroke="' + chartTheme.axis + '"></line>' +
        lineSvg +
        '<text x="' + (margin.left + chartWidth / 2).toFixed(2) + '" y="' + (height - 6) + '" text-anchor="middle" font-size="12" font-weight="700" fill="' + chartTheme.text + '">Days</text>' +
        '<text x="18" y="' + (margin.top + chartHeight / 2).toFixed(2) + '" text-anchor="middle" font-size="12" font-weight="700" fill="' + chartTheme.text + '" transform="rotate(-90 18 ' + (margin.top + chartHeight / 2).toFixed(2) + ')">Depth / Section Progression</text>' +
        "</svg>" +
        "</div>";
    }

    function average(numbers) {
      const values = numbers.filter((value) => Number.isFinite(value) && value > 0);
      if (!values.length) return 0;
      return values.reduce((sum, value) => sum + value, 0) / values.length;
    }

    function uniqueCount(rows, key) {
      return new Set(rows.map((row) => row[key]).filter(Boolean)).size;
    }

    function buildCounter(rows, key) {
      const counts = new Map();
      rows.forEach((row) => {
        const label = (row[key] || "").trim();
        if (!label) return;
        counts.set(label, (counts.get(label) || 0) + 1);
      });
      return Array.from(counts.entries())
        .map(([label, value]) => ({ label, value }))
        .sort((left, right) => right.value - left.value || left.label.localeCompare(right.label));
    }

    function buildValueCounter(rows, key, valueKey) {
      const counts = new Map();
      rows.forEach((row) => {
        const label = (row[key] || "").trim();
        const value = Number(row[valueKey] || 0);
        if (!label || !Number.isFinite(value)) return;
        counts.set(label, (counts.get(label) || 0) + value);
      });
      return Array.from(counts.entries())
        .map(([label, value]) => ({ label, value }))
        .sort((left, right) => right.value - left.value || left.label.localeCompare(right.label));
    }

    function buildTrend(rows, granularity) {
      const key = granularity === "week" ? "week" : granularity === "month" ? "month" : "date";
      const counts = new Map();
      rows.forEach((row) => {
        const label = row[key];
        if (!label) return;
        counts.set(label, (counts.get(label) || 0) + 1);
      });
      return Array.from(counts.entries())
        .map(([label, value]) => ({ label, value }))
        .sort((left, right) => left.label.localeCompare(right.label))
        .slice(-16);
    }

    function populateSelect(element, options, allLabel) {
      const current = element.value;
      const htmlOptions = ['<option value="">' + escapeHtml(allLabel) + "</option>"];
      options.forEach((option) => {
        htmlOptions.push('<option value="' + escapeHtml(option) + '">' + escapeHtml(option) + "</option>");
      });
      element.innerHTML = htmlOptions.join("");
      if (options.includes(current)) {
        element.value = current;
      }
    }

    function renderBarChart(target, items, color, formatter) {
      if (!items.length) {
        target.innerHTML = '<div class="empty">No data available for the selected filters.</div>';
        return;
      }
      const trimmed = items.slice(0, 8);
      const maxValue = Math.max(...trimmed.map((item) => item.value), 0);
      target.innerHTML =
        '<div class="bar-list">' +
        trimmed
          .map((item) => {
            const width = maxValue === 0 ? 0 : (item.value / maxValue) * 100;
            return (
              '<div class="bar-row">' +
              '<div class="bar-label" title="' + escapeHtml(item.label) + '">' + escapeHtml(item.label) + "</div>" +
              '<div class="bar-track"><div class="bar-fill" style="width:' + width.toFixed(1) + "%; background:" + color + ';"></div></div>' +
              '<div class="bar-value">' + escapeHtml(formatter(item.value)) + "</div>" +
              "</div>"
            );
          })
          .join("") +
        "</div>";
    }

    function renderTrendChart(target, items) {
      if (!items.length) {
        target.innerHTML = '<div class="empty">No trend data available for the selected filters.</div>';
        return;
      }

      const chartTheme = getChartTheme();

      const width = 920;
      const height = 280;
      const paddingX = 48;
      const paddingTop = 24;
      const paddingBottom = 42;
      const chartWidth = width - paddingX * 2;
      const chartHeight = height - paddingTop - paddingBottom;
      const maxValue = Math.max(...items.map((item) => item.value), 1);

      const points = items.map((item, index) => {
        const x = items.length === 1 ? width / 2 : paddingX + (chartWidth * index) / (items.length - 1);
        const y = paddingTop + chartHeight - (item.value / maxValue) * chartHeight;
        return { ...item, x, y };
      });

      const path = points.map((point, index) => (index === 0 ? "M" : "L") + point.x.toFixed(2) + " " + point.y.toFixed(2)).join(" ");
      const areaPath = path + " L " + points[points.length - 1].x.toFixed(2) + " " + (paddingTop + chartHeight) + " L " + points[0].x.toFixed(2) + " " + (paddingTop + chartHeight) + " Z";

      const labels = points
        .map((point) => {
          return '<text x="' + point.x.toFixed(2) + '" y="' + (height - 10) + '" text-anchor="middle" font-size="11" fill="' + chartTheme.muted + '">' + escapeHtml(point.label) + "</text>";
        })
        .join("");

      const dots = points
        .map((point) => {
          return (
            '<g>' +
            '<circle cx="' + point.x.toFixed(2) + '" cy="' + point.y.toFixed(2) + '" r="4.5" fill="' + chartTheme.line + '"></circle>' +
            '<text x="' + point.x.toFixed(2) + '" y="' + (point.y - 10).toFixed(2) + '" text-anchor="middle" font-size="11" fill="' + chartTheme.pointLabel + '">' + escapeHtml(String(point.value)) + "</text>" +
            "</g>"
          );
        })
        .join("");

      const grid = Array.from({ length: 5 }, (_, index) => {
        const y = paddingTop + (chartHeight * index) / 4;
        return '<line x1="' + paddingX + '" y1="' + y.toFixed(2) + '" x2="' + (width - paddingX) + '" y2="' + y.toFixed(2) + '" stroke="' + chartTheme.grid + '" stroke-dasharray="4 5"></line>';
      }).join("");

      target.innerHTML =
        '<svg class="line-svg" viewBox="0 0 ' + width + " " + height + '" role="img" aria-label="Intervention trend">' +
        grid +
        '<path d="' + areaPath + '" fill="' + chartTheme.area + '"></path>' +
        '<path d="' + path + '" fill="none" stroke="' + chartTheme.line + '" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"></path>' +
        dots +
        labels +
        "</svg>";
    }

    function wrapChartLabel(label, maxLineLength) {
      const words = String(label || "").split(/\\s+/).filter(Boolean);
      if (!words.length) return [""];
      const lines = [];
      let current = words[0];
      for (let index = 1; index < words.length; index += 1) {
        const next = words[index];
        if ((current + " " + next).length <= maxLineLength) {
          current += " " + next;
        } else {
          lines.push(current);
          current = next;
        }
      }
      lines.push(current);
      return lines.slice(0, 3);
    }

    function niceMax(value) {
      if (value <= 0) return 1;
      const exponent = Math.floor(Math.log10(value));
      const fraction = value / Math.pow(10, exponent);
      let niceFraction = 1;
      if (fraction <= 1) niceFraction = 1;
      else if (fraction <= 2) niceFraction = 2;
      else if (fraction <= 5) niceFraction = 5;
      else niceFraction = 10;
      return niceFraction * Math.pow(10, exponent);
    }

    function renderMultiSeriesChart(target, items, seriesDefs, options = {}) {
      if (!items.length) {
        target.innerHTML = '<div class="empty">No data available for this block.</div>';
        return;
      }

      const chartTheme = getChartTheme();

      const primaryMax = Math.max(
        ...items.flatMap((item) =>
          seriesDefs
            .filter((series) => !series.isSecondary)
            .map((series) => Number(item[series.key] || 0))
        ),
        0
      );

      const rawMax = Math.max(
        ...items.flatMap((item) => seriesDefs.map((series) => Number(item[series.key] || 0))),
        0
      );

      const context = { primaryMax: primaryMax || rawMax || 1, rawMax: rawMax || 1 };
      const preparedItems = items.map((item) => ({
        ...item,
        series: seriesDefs.map((series) => {
          const rawValue = Number(item[series.key] || 0);
          const scaledValue = series.scale ? Number(series.scale(rawValue, context)) : rawValue;
          return {
            ...series,
            rawValue,
            scaledValue,
            formatted: series.format ? series.format(rawValue) : formatNumber(rawValue),
          };
        }),
      }));

      const maxValue = niceMax(
        Math.max(
          ...preparedItems.flatMap((item) => item.series.map((series) => series.scaledValue)),
          0
        )
      );

      const width = Math.max(options.minWidth || 620, items.length * Math.max(options.groupMinWidth || 150, seriesDefs.length * 44 + 70));
      const height = options.height || 360;
      const margin = { top: 28, right: 18, bottom: 88, left: 52 };
      const chartWidth = width - margin.left - margin.right;
      const chartHeight = height - margin.top - margin.bottom;
      const groupWidth = chartWidth / items.length;
      const clusterWidth = Math.min(groupWidth - 20, seriesDefs.length * 38 + (seriesDefs.length - 1) * 10);
      const barWidth = Math.max(16, Math.min(34, (clusterWidth - (seriesDefs.length - 1) * 10) / seriesDefs.length));
      const clusterOffset = (groupWidth - (barWidth * seriesDefs.length + (seriesDefs.length - 1) * 10)) / 2;
      const tickCount = 4;

      const grid = Array.from({ length: tickCount + 1 }, (_, index) => {
        const value = (maxValue / tickCount) * index;
        const y = margin.top + chartHeight - (value / maxValue) * chartHeight;
        return (
          '<g>' +
          '<line x1="' + margin.left + '" y1="' + y.toFixed(2) + '" x2="' + (width - margin.right) + '" y2="' + y.toFixed(2) + '" stroke="' + chartTheme.grid + '" stroke-dasharray="4 5"></line>' +
          '<text x="' + (margin.left - 10) + '" y="' + (y + 4).toFixed(2) + '" text-anchor="end" font-size="11" fill="' + chartTheme.valueLabel + '">' + escapeHtml(formatNumber(value)) + "</text>" +
          "</g>"
        );
      }).join("");

      const bars = preparedItems
        .map((item, itemIndex) => {
          const groupX = margin.left + itemIndex * groupWidth;
          const labelLines = wrapChartLabel(item.label, 18);
          const labelX = groupX + groupWidth / 2;
          const labelY = height - 38;

          const labelSvg = labelLines
            .map((line, lineIndex) => {
              const dy = lineIndex === 0 ? 0 : 14;
              return '<tspan x="' + labelX.toFixed(2) + '" dy="' + dy + '">' + escapeHtml(line) + "</tspan>";
            })
            .join("");

          const barsSvg = item.series
            .map((series, seriesIndex) => {
              const x = groupX + clusterOffset + seriesIndex * (barWidth + 10);
              const barHeight = maxValue === 0 ? 0 : (series.scaledValue / maxValue) * chartHeight;
              const y = margin.top + chartHeight - barHeight;
              const displayY = Math.max(margin.top + 12, y - 8);
              return (
                '<g>' +
                '<rect x="' + x.toFixed(2) + '" y="' + y.toFixed(2) + '" width="' + barWidth.toFixed(2) + '" height="' + Math.max(2, barHeight).toFixed(2) + '" rx="10" fill="' + series.color + '"></rect>' +
                '<text x="' + (x + barWidth / 2).toFixed(2) + '" y="' + displayY.toFixed(2) + '" text-anchor="middle" font-size="11" font-weight="700" fill="' + chartTheme.pointLabel + '">' + escapeHtml(series.formatted) + "</text>" +
                "</g>"
              );
            })
            .join("");

          return (
            '<g>' +
            barsSvg +
            '<text x="' + labelX.toFixed(2) + '" y="' + labelY + '" text-anchor="middle" font-size="12" font-weight="700" fill="' + chartTheme.text + '">' + labelSvg + "</text>" +
            "</g>"
          );
        })
        .join("");

      target.innerHTML =
        '<div class="column-chart-wrap">' +
        '<svg class="column-chart-svg" viewBox="0 0 ' + width + " " + height + '" role="img" aria-label="Grouped column chart">' +
        grid +
        '<line x1="' + margin.left + '" y1="' + (margin.top + chartHeight).toFixed(2) + '" x2="' + (width - margin.right) + '" y2="' + (margin.top + chartHeight).toFixed(2) + '" stroke="' + chartTheme.axis + '"></line>' +
        bars +
        "</svg>" +
        "</div>";
    }

    function renderTable(target, headers, rows) {
      if (!rows.length) {
        target.innerHTML = '<div class="empty">No rows available for the selected filters.</div>';
        return;
      }

      const headerHtml = headers.map((header) => "<th>" + escapeHtml(header) + "</th>").join("");
      const bodyHtml = rows
        .map((row) => {
          return "<tr>" + row.map((cell) => "<td>" + escapeHtml(cell) + "</td>").join("") + "</tr>";
        })
        .join("");

      target.innerHTML = '<div class="table-wrap"><table><thead><tr>' + headerHtml + "</tr></thead><tbody>" + bodyHtml + "</tbody></table></div>";
    }

    function renderTableHtml(target, headers, rows) {
      if (!rows.length) {
        target.innerHTML = '<div class="empty">No rows available for the selected filters.</div>';
        return;
      }

      const headerHtml = headers.map((header) => "<th>" + escapeHtml(header) + "</th>").join("");
      const bodyHtml = rows
        .map((row) => {
          return "<tr>" + row.map((cell) => "<td>" + cell + "</td>").join("") + "</tr>";
        })
        .join("");

      target.innerHTML = '<div class="table-wrap"><table><thead><tr>' + headerHtml + "</tr></thead><tbody>" + bodyHtml + "</tbody></table></div>";
    }

    function confidenceBadgeHtml(confidence) {
      const tone = String(confidence || "low").trim().toLowerCase();
      return '<span class="confidence-badge ' + escapeHtml(tone) + '">' + escapeHtml(confidence || "Low") + '</span>';
    }

    function flatTimeActionButtonHtml(type, value, label) {
      const title = type === "activity" ? getFlatTimeActivityTooltip(label) : "";
      return '<button type="button" class="table-action" data-flat-focus-' + escapeHtml(type) + '="' + escapeHtml(value) + '"' + (title ? ' title="' + escapeHtml(title) + '"' : "") + '>' + escapeHtml(label) + '</button>';
    }

    function getFlatTimeActivityTooltip(activityLabel) {
      const label = String(activityLabel || "").trim();
      if (!label) return "";

      const translations = FLAT_TIME_ACTIVITY_TRANSLATIONS || {};
      const wellSections = translations.wellSections || {};
      const operations = translations.operations || {};
      const activities = translations.activities || {};
      const generic = translations.generic || {};
      const parts = label.split("-").filter(Boolean);
      const lines = [];

      if (parts.length >= 3) {
        const sectionCode = parts[0];
        const operationCode = parts[1];
        const activityCode = parts[parts.length - 1];

        if (wellSections[sectionCode]) {
          lines.push("Section " + sectionCode + ': ' + wellSections[sectionCode]);
        }
        if (operations[operationCode]) {
          lines.push("Operation " + operationCode + ': ' + operations[operationCode]);
        }
        if (activities[activityCode]) {
          lines.push("Activity " + activityCode + ': ' + activities[activityCode]);
        } else if (generic[activityCode]) {
          lines.push("Activity " + activityCode + ': ' + generic[activityCode]);
        }
      }

      if (!lines.length && generic[label]) {
        lines.push(generic[label]);
      }

      return lines.join("\n");
    }

    function flatTimeActivityLabelHtml(label) {
      const tooltip = getFlatTimeActivityTooltip(label);
      return '<span' + (tooltip ? ' title="' + escapeHtml(tooltip) + '"' : "") + '>' + escapeHtml(label) + '</span>';
    }

    function flatTimeTrendHtml(excessHours) {
      const value = Number(excessHours || 0);
      const isSlow = value > 0.01;
      const tone = isSlow ? "slow" : "fast";
      const arrow = isSlow ? "▲" : "▼";
      const label = isSlow ? "slower" : "faster";
      return (
        '<span class="trend-indicator ' + tone + '">' +
        '<span>' + escapeHtml(formatNumber(value)) + '</span>' +
        '<span class="arrow">' + arrow + '</span>' +
        '<span>' + label + '</span>' +
        '</span>'
      );
    }

    function isYesLike(value) {
      return ["yes", "y", "true", "confirmed"].includes(String(value || "").trim().toLowerCase());
    }

    function normalizeCategory(category) {
      const value = String(category || "").trim().toLowerCase();
      if (value === "stuck pipe") return "Stuck pipe";
      if (value === "optimization") return "Optimization";
      if (value === "operational compliance") return "Operational Compliance";
      if (value === "well control") return "Well Control";
      if (value === "reporting" || value === "reporting ") return "Reporting";
      return category || "Other";
    }

    function formatPercent(value) {
      return (Number(value || 0) * 100).toFixed(1) + "%";
    }

    function formatHoursWithDays(value) {
      const hours = Number(value || 0);
      return formatNumber(hours) + " hr / " + formatNumber(hours / 24) + " d";
    }

    function formatDateHuman(dateString) {
      if (!dateString) return "";
      return new Date(dateString + "T00:00:00").toLocaleDateString("en-GB", {
        day: "2-digit",
        month: "short",
        year: "numeric",
      });
    }

    function weekBounds(week) {
      const dates = dashboardData.interventions
        .filter((row) => row.week === week && row.date)
        .map((row) => row.date)
        .sort();
      return {
        start: dates[0] || "",
        end: dates[dates.length - 1] || "",
      };
    }

    function formatWeekLabel(week) {
      const bounds = weekBounds(week);
      if (!bounds.start || !bounds.end) return week || "Unknown week";
      return formatDateHuman(bounds.start) + " - " + formatDateHuman(bounds.end);
    }

    function toIsoDate(date) {
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, "0");
      const day = String(date.getDate()).padStart(2, "0");
      return year + "-" + month + "-" + day;
    }

    function getDefaultLastTuesdayRange() {
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const dayOfWeek = today.getDay();
      const diffToTuesday = (dayOfWeek - 2 + 7) % 7;
      const end = new Date(today);
      end.setDate(today.getDate() - diffToTuesday);
      const start = new Date(end);
      start.setDate(end.getDate() - 6);
      return {
        start: toIsoDate(start),
        end: toIsoDate(end),
      };
    }

    function getWeeklyReportDateRange() {
      let start = ui.weeklyReportStartDate.value;
      let end = ui.weeklyReportEndDate.value;

      if (start && end && start > end) {
        return { start: end, end: start };
      }
      return { start, end };
    }

    function textBlob(row) {
      return [
        row.description,
        row.recommendation,
        row.parameter,
        row.type,
        row.app,
        row.expected,
        row.actual,
        row.rigComment,
        row.rtocComments,
      ]
        .join(" ")
        .toLowerCase();
    }

    function isWiperTripRow(row) {
      return textBlob(row).includes("wiper");
    }

    function isRopRow(row) {
      return row.type.toLowerCase() === "rop" || row.parameter.toLowerCase() === "rop" || /\brop\b/.test(textBlob(row));
    }

    function isKpiRow(row) {
      const parameter = row.parameter.toLowerCase();
      const type = row.type.toLowerCase();
      return (
        type === "kpi" ||
        ["w2w", "s2s", "tripping speed", "tripping speed ", "connection", "kpi", "time"].includes(parameter) ||
        textBlob(row).includes("kpi")
      );
    }

    function buildCategorySummary(rows) {
      return CATEGORY_ORDER.map((category) => {
        const matches = rows.filter((row) => normalizeCategory(row.category) === category);
        const interventions = matches.length;
        const rigAction = matches.filter((row) => isYesLike(row.rigAction)).length;
        const validated = matches.filter((row) => row.isValidated).length;
        return {
          label: category,
          interventions,
          rigAction,
          validated,
          validationRate: interventions ? validated / interventions : 0,
        };
      });
    }

    function buildRigSummary(rows, matcher) {
      const groups = new Map();
      rows.filter(matcher).forEach((row) => {
        const key = row.rigName + "||" + row.wellName;
        if (!groups.has(key)) {
          groups.set(key, {
            label: row.rigName + " / " + row.wellName,
            rig: row.rigName,
            well: row.wellName,
            interventions: 0,
            rigAction: 0,
            savedTime: 0,
            lossTime: 0,
          });
        }
        const entry = groups.get(key);
        entry.interventions += 1;
        entry.rigAction += isYesLike(row.rigAction) ? 1 : 0;
        entry.savedTime += row.costSavingHours;
        entry.lossTime += row.potentialAvoidanceHours;
      });

      const items = Array.from(groups.values()).sort(
        (left, right) =>
          (right.savedTime + right.lossTime + right.interventions) -
            (left.savedTime + left.lossTime + left.interventions) ||
          left.label.localeCompare(right.label)
      );

      const totals = items.reduce(
        (acc, item) => {
          acc.interventions += item.interventions;
          acc.rigAction += item.rigAction;
          acc.savedTime += item.savedTime;
          acc.lossTime += item.lossTime;
          return acc;
        },
        { interventions: 0, rigAction: 0, savedTime: 0, lossTime: 0 }
      );

      return { items, totals };
    }

    function providerLabel(row) {
      return row.engDept || row.optDept || "N/A";
    }

    function buildHighlightEntries(rows, mode) {
      const isActual = mode === "actual";
      return rows
        .filter((row) =>
          isActual
            ? row.costSavingHours > 0 || row.costSavingValue > 0
            : row.potentialAvoidanceHours > 0 || row.potentialAvoidanceValue > 0
        )
        .sort((left, right) =>
          isActual
            ? right.costSavingValue - left.costSavingValue || (right.date || "").localeCompare(left.date || "")
            : right.potentialAvoidanceValue - left.potentialAvoidanceValue || (right.date || "").localeCompare(left.date || "")
        )
        .map((row) => ({
          week: row.week,
          date: row.date,
          rig: row.rigName,
          well: row.wellName,
          provider: providerLabel(row),
          action: isActual ? row.description || row.recommendation : row.recommendation || row.description,
          hours: isActual ? row.costSavingHours : row.potentialAvoidanceHours,
          value: isActual ? row.costSavingValue : row.potentialAvoidanceValue,
        }));
    }

    function renderWeeklyMetrics(actualWeek, actualYtd, potentialWeek, potentialYtd) {
      const cards = [
        {
          label: "Total This Week",
          value: formatNumber(actualWeek.reduce((sum, row) => sum + row.hours, 0)),
          suffix: "hours saved",
          meta: formatCurrency(actualWeek.reduce((sum, row) => sum + row.value, 0)) + " value realized",
        },
        {
          label: "Total YTD",
          value: formatNumber(actualYtd.reduce((sum, row) => sum + row.hours, 0)),
          suffix: "hours saved",
          meta: formatCurrency(actualYtd.reduce((sum, row) => sum + row.value, 0)) + " value realized",
        },
        {
          label: "Potential This Week",
          value: formatNumber(potentialWeek.reduce((sum, row) => sum + row.hours, 0)),
          suffix: "hours potential",
          meta: formatCurrency(potentialWeek.reduce((sum, row) => sum + row.value, 0)) + " value potential",
        },
        {
          label: "Potential YTD",
          value: formatNumber(potentialYtd.reduce((sum, row) => sum + row.hours, 0)),
          suffix: "hours potential",
          meta: formatCurrency(potentialYtd.reduce((sum, row) => sum + row.value, 0)) + " value potential",
        },
      ];

      ui.weeklyHighlightMetrics.innerHTML = cards
        .map(
          (card) =>
            '<div class="metric-pill">' +
            '<div class="label">' + escapeHtml(card.label) + "</div>" +
            '<div class="value"><span class="value-main">' + escapeHtml(card.value) + '</span><span class="value-suffix">' + escapeHtml(card.suffix) + "</span></div>" +
            '<div class="meta">' + escapeHtml(card.meta) + "</div>" +
            "</div>"
        )
        .join("");
    }

    function renderHighlightTable(target, weekEntries, ytdEntries, isActual) {
      const headers = isActual
        ? ["Rig", "Well", "Provider / Dept", "Operations / Action", "Saved Time (hrs)", "Cost Saving (US$)"]
        : ["Rig", "Well", "Provider / Dept", "Operations / Action", "Potential Saved Time (hrs)", "Potential Cost Saving/Avoidance (US$)"];

      const bodyRows = weekEntries
        .slice(0, 25)
        .map((row) => {
          return (
            "<tr>" +
            "<td>" + escapeHtml(row.rig) + "</td>" +
            "<td>" + escapeHtml(row.well) + "</td>" +
            "<td>" + escapeHtml(row.provider) + "</td>" +
            "<td>" + escapeHtml(row.action) + "</td>" +
            "<td>" + escapeHtml(formatNumber(row.hours)) + "</td>" +
            "<td>" + escapeHtml(formatCurrency(row.value)) + "</td>" +
            "</tr>"
          );
        })
        .join("");

      const totalWeekHours = weekEntries.reduce((sum, row) => sum + row.hours, 0);
      const totalWeekValue = weekEntries.reduce((sum, row) => sum + row.value, 0);
      const totalYtdHours = ytdEntries.reduce((sum, row) => sum + row.hours, 0);
      const totalYtdValue = ytdEntries.reduce((sum, row) => sum + row.value, 0);

      const totalsRows =
        '<tr>' +
        '<td colspan="4" style="text-align:right; font-weight:700;">Total This Week</td>' +
        '<td><strong>' + escapeHtml(formatNumber(totalWeekHours)) + "</strong></td>" +
        '<td><strong>' + escapeHtml(formatCurrency(totalWeekValue)) + "</strong></td>" +
        "</tr>" +
        '<tr>' +
        '<td colspan="4" style="text-align:right; font-weight:700;">Total YTD</td>' +
        '<td><strong>' + escapeHtml(formatNumber(totalYtdHours)) + "</strong></td>" +
        '<td><strong>' + escapeHtml(formatCurrency(totalYtdValue)) + "</strong></td>" +
        "</tr>";

      target.innerHTML =
        '<div class="table-wrap"><table><thead><tr>' +
        headers.map((header) => "<th>" + escapeHtml(header) + "</th>").join("") +
        "</tr></thead><tbody>" +
        bodyRows +
        totalsRows +
        "</tbody></table></div>";
    }

    function renderSummaryBlock(tableTarget, chartTarget, summary, chartSeries) {
      const tableRows = summary.items
        .map((item) => [
          item.rig,
          item.well,
          String(item.interventions),
          String(item.rigAction),
          formatNumber(item.savedTime),
          formatNumber(item.lossTime),
        ])
        .concat([
          [
            "Total",
            "",
            String(summary.totals.interventions),
            String(summary.totals.rigAction),
            formatNumber(summary.totals.savedTime),
            formatNumber(summary.totals.lossTime),
          ],
        ]);

      renderTable(
        tableTarget,
        ["Rig Name", "Well Name", "Number of Interventions", "Rig Action Taken", "Saved Time (hrs)", "Loss Time (hrs)"],
        tableRows
      );

      renderMultiSeriesChart(chartTarget, summary.items, chartSeries);
    }

    function daysBetweenInclusive(start, end) {
      if (!start || !end) return 0;
      const startDate = new Date(start + "T00:00:00");
      const endDate = new Date(end + "T00:00:00");
      const diff = Math.round((endDate - startDate) / 86400000);
      return diff >= 0 ? diff + 1 : 0;
    }

    function buildWeeklyStatsRows(rows, cumulativeRows) {
      const groups = new Map();
      const cumulativeGroups = new Map();

      rows.forEach((row) => {
        const key = row.rigName + "||" + row.wellName;
        if (!groups.has(key)) {
          groups.set(key, { rig: row.rigName, well: row.wellName, rows: [] });
        }
        groups.get(key).rows.push(row);
      });

      cumulativeRows.forEach((row) => {
        const key = row.rigName + "||" + row.wellName;
        if (!cumulativeGroups.has(key)) {
          cumulativeGroups.set(key, []);
        }
        cumulativeGroups.get(key).push(row);
      });

      return Array.from(groups.values())
        .sort((left, right) => left.rig.localeCompare(right.rig) || left.well.localeCompare(right.well))
        .map((group) => {
          const byCategory = (categoryName) => {
            const matches = group.rows.filter((row) => normalizeCategory(row.category) === categoryName);
            const count = matches.length;
            const validated = matches.filter((row) => row.isValidated).length;
            return {
              count,
              validity: count ? formatPercent(validated / count) : "0.0%",
            };
          };

          const totalCount = group.rows.length;
          const totalValidated = group.rows.filter((row) => row.isValidated).length;
          const key = group.rig + "||" + group.well;
          const cumulativeGroupRows = cumulativeGroups.get(key) || [];
          const monitoredThisWeek = new Set(group.rows.map((row) => row.date).filter(Boolean)).size;
          const monitoredSinceStart = new Set(cumulativeGroupRows.map((row) => row.date).filter(Boolean)).size;

          return {
            rig: group.rig,
            well: group.well,
            optimization: byCategory("Optimization"),
            stuckPipe: byCategory("Stuck pipe"),
            wellControl: byCategory("Well Control"),
            operationalCompliance: byCategory("Operational Compliance"),
            reporting: byCategory("Reporting"),
            total: {
              count: totalCount,
              validity: totalCount ? formatPercent(totalValidated / totalCount) : "0.0%",
            },
            monitoredThisWeek,
            monitoredSinceStart,
          };
        });
    }

    function renderWeeklyStatsMetrics(target, thisWeekDays, totalDays) {
      const cards = [
        {
          label: "Days Monitored This Week",
          value: String(thisWeekDays),
          suffix: "days",
          meta: "Selected reporting range",
        },
        {
          label: "Days Monitored Since Start",
          value: String(totalDays),
          suffix: "days",
          meta: "From monitoring start date to selected end date",
        },
      ];

      target.innerHTML = cards
        .map(
          (card) =>
            '<div class="metric-pill">' +
            '<div class="label">' + escapeHtml(card.label) + "</div>" +
            '<div class="value"><span class="value-main">' + escapeHtml(card.value) + '</span><span class="value-suffix">' + escapeHtml(card.suffix) + "</span></div>" +
            '<div class="meta">' + escapeHtml(card.meta) + "</div>" +
            "</div>"
        )
        .join("");
    }

    function renderWeeklyStatsTable(target, rows) {
      if (!rows.length) {
        target.innerHTML = '<div class="empty">No intervention statistics available for the selected period.</div>';
        return;
      }

      const headerHtml =
        "<thead>" +
        '<tr><th colspan="14">Weekly Intervention Statistics and Analysis</th></tr>' +
        '<tr>' +
        '<th colspan="2"></th>' +
        '<th colspan="2">Optimization</th>' +
        '<th colspan="2">Stuck Pipe</th>' +
        '<th colspan="2">Well Control</th>' +
        '<th colspan="2">Operational Compliance</th>' +
        '<th colspan="2">Reporting</th>' +
        '<th>Days Monitored This Week</th>' +
        '<th>Days Monitored Since Start</th>' +
        "</tr>" +
        '<tr>' +
        '<th>Rig Name</th><th>Well Name</th>' +
        '<th># Interventions</th><th>Validity %</th>' +
        '<th># Interventions</th><th>Validity %</th>' +
        '<th># Interventions</th><th>Validity %</th>' +
        '<th># Interventions</th><th>Validity %</th>' +
        '<th># Interventions</th><th>Validity %</th>' +
        '<th>Days</th><th>Days</th>' +
        "</tr>" +
        "</thead>";

      const totalRow = rows.reduce(
        (acc, row) => {
          acc.optimization.count += row.optimization.count;
          acc.optimization.validityCount += row.optimization.count ? Number(row.optimization.validity.replace("%", "")) * row.optimization.count : 0;
          acc.optimization.den += row.optimization.count;
          acc.stuckPipe.count += row.stuckPipe.count;
          acc.stuckPipe.validityCount += row.stuckPipe.count ? Number(row.stuckPipe.validity.replace("%", "")) * row.stuckPipe.count : 0;
          acc.stuckPipe.den += row.stuckPipe.count;
          acc.wellControl.count += row.wellControl.count;
          acc.wellControl.validityCount += row.wellControl.count ? Number(row.wellControl.validity.replace("%", "")) * row.wellControl.count : 0;
          acc.wellControl.den += row.wellControl.count;
          acc.operationalCompliance.count += row.operationalCompliance.count;
          acc.operationalCompliance.validityCount += row.operationalCompliance.count ? Number(row.operationalCompliance.validity.replace("%", "")) * row.operationalCompliance.count : 0;
          acc.operationalCompliance.den += row.operationalCompliance.count;
          acc.reporting.count += row.reporting.count;
          acc.reporting.validityCount += row.reporting.count ? Number(row.reporting.validity.replace("%", "")) * row.reporting.count : 0;
          acc.reporting.den += row.reporting.count;
          acc.monitoredThisWeek += row.monitoredThisWeek;
          acc.monitoredSinceStart += row.monitoredSinceStart;
          return acc;
        },
        {
          optimization: { count: 0, validityCount: 0, den: 0 },
          stuckPipe: { count: 0, validityCount: 0, den: 0 },
          wellControl: { count: 0, validityCount: 0, den: 0 },
          operationalCompliance: { count: 0, validityCount: 0, den: 0 },
          reporting: { count: 0, validityCount: 0, den: 0 },
          monitoredThisWeek: 0,
          monitoredSinceStart: 0,
        }
      );

      const pct = (bucket) => bucket.den ? (bucket.validityCount / bucket.den).toFixed(1) + "%" : "0.0%";

      const bodyHtml = rows
        .map((row) => {
          return (
            "<tr>" +
            "<td>" + escapeHtml(row.rig) + "</td>" +
            "<td>" + escapeHtml(row.well) + "</td>" +
            "<td>" + escapeHtml(String(row.optimization.count)) + "</td>" +
            "<td>" + escapeHtml(row.optimization.validity) + "</td>" +
            "<td>" + escapeHtml(String(row.stuckPipe.count)) + "</td>" +
            "<td>" + escapeHtml(row.stuckPipe.validity) + "</td>" +
            "<td>" + escapeHtml(String(row.wellControl.count)) + "</td>" +
            "<td>" + escapeHtml(row.wellControl.validity) + "</td>" +
            "<td>" + escapeHtml(String(row.operationalCompliance.count)) + "</td>" +
            "<td>" + escapeHtml(row.operationalCompliance.validity) + "</td>" +
            "<td>" + escapeHtml(String(row.reporting.count)) + "</td>" +
            "<td>" + escapeHtml(row.reporting.validity) + "</td>" +
            "<td><strong>" + escapeHtml(String(row.monitoredThisWeek)) + "</strong></td>" +
            "<td><strong>" + escapeHtml(String(row.monitoredSinceStart)) + "</strong></td>" +
            "</tr>"
          );
        })
        .join("") +
        (
          "<tr>" +
          "<td><strong>Total</strong></td>" +
          "<td></td>" +
          "<td><strong>" + escapeHtml(String(totalRow.optimization.count)) + "</strong></td>" +
          "<td><strong>" + escapeHtml(pct(totalRow.optimization)) + "</strong></td>" +
          "<td><strong>" + escapeHtml(String(totalRow.stuckPipe.count)) + "</strong></td>" +
          "<td><strong>" + escapeHtml(pct(totalRow.stuckPipe)) + "</strong></td>" +
          "<td><strong>" + escapeHtml(String(totalRow.wellControl.count)) + "</strong></td>" +
          "<td><strong>" + escapeHtml(pct(totalRow.wellControl)) + "</strong></td>" +
          "<td><strong>" + escapeHtml(String(totalRow.operationalCompliance.count)) + "</strong></td>" +
          "<td><strong>" + escapeHtml(pct(totalRow.operationalCompliance)) + "</strong></td>" +
          "<td><strong>" + escapeHtml(String(totalRow.reporting.count)) + "</strong></td>" +
          "<td><strong>" + escapeHtml(pct(totalRow.reporting)) + "</strong></td>" +
          "<td><strong>" + escapeHtml(String(totalRow.monitoredThisWeek)) + "</strong></td>" +
          "<td><strong>" + escapeHtml(String(totalRow.monitoredSinceStart)) + "</strong></td>" +
          "</tr>"
        );

      target.innerHTML = '<div class="table-wrap"><table class="stats-table">' + headerHtml + "<tbody>" + bodyHtml + "</tbody></table></div>";
    }

    function buildFlatTimeGroupItems(datasets, totalKey) {
      const groupNames = Array.from(
        new Set(datasets.flatMap((dataset) => dataset.groups.map((group) => group.groupName)))
      );

      return groupNames
        .map((groupName) => {
          const item = { label: groupName };
          datasets.forEach((dataset) => {
            const match = dataset.groups.find((group) => group.groupName === groupName);
            item[dataset.id] = match ? Number(match[totalKey] || 0) : 0;
          });
          item.total = datasets.reduce((sum, dataset) => sum + Number(item[dataset.id] || 0), 0);
          return item;
        })
        .sort((left, right) => right.total - left.total || left.label.localeCompare(right.label));
    }

    function buildFlatTimeActivityItems(datasets, metricKey) {
      const activityMap = new Map();
      datasets.forEach((dataset) => {
        dataset.groups.forEach((group) => {
          group.activities.forEach((activity) => {
            const key = activity.activity;
            if (!activityMap.has(key)) {
              activityMap.set(key, {
                label: key,
                groupLabel: group.groupName,
                sectionSize: activity.sectionSize || extractFlatTimeSectionSize(activity.activity),
              });
            }
            activityMap.get(key)[dataset.id] = Number(activity[metricKey] || 0);
          });
        });
      });

      return Array.from(activityMap.values())
        .map((item) => {
          datasets.forEach((dataset) => {
            item[dataset.id] = Number(item[dataset.id] || 0);
          });
          item.total = datasets.reduce((sum, dataset) => sum + Number(item[dataset.id] || 0), 0);
          return item;
        })
        .sort((left, right) => right.total - left.total || left.label.localeCompare(right.label));
    }

    function percentile(numbers, ratio) {
      const values = numbers.filter((value) => Number.isFinite(value)).slice().sort((left, right) => left - right);
      if (!values.length) return 0;
      if (values.length === 1) return values[0];
      const position = (values.length - 1) * ratio;
      const lower = Math.floor(position);
      const upper = Math.ceil(position);
      if (lower === upper) return values[lower];
      const weight = position - lower;
      return values[lower] * (1 - weight) + values[upper] * weight;
    }

    function standardDeviation(numbers) {
      const values = numbers.filter((value) => Number.isFinite(value));
      if (values.length <= 1) return 0;
      const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
      const variance = values.reduce((sum, value) => sum + Math.pow(value - mean, 2), 0) / values.length;
      return Math.sqrt(variance);
    }

    function flatTimeConfidence(sampleSize, cv) {
      if (sampleSize >= 5 && cv <= 0.2) return "High";
      if (sampleSize >= 3 && cv <= 0.35) return "Medium";
      return "Low";
    }

    function flatTimeVariabilityLabel(cv) {
      if (cv <= 0.2) return "Low";
      if (cv <= 0.35) return "Moderate";
      return "High";
    }

    function computeFlatTimeOpportunity(item, datasets) {
      const ranked = datasets
        .map((dataset) => ({
          datasetId: dataset.id,
          label: dataset.subjectWell,
          rigLabel: dataset.rigLabel || "Rig not mapped",
          value: Number(item[dataset.id] || 0),
        }))
        .filter((entry) => entry.value > 0)
        .sort((left, right) => right.value - left.value);

      const values = ranked.map((entry) => entry.value);
      const occurrenceCount = values.length;
      const topEntry = ranked[0] || { label: "N/A", rigLabel: "Rig not mapped", value: 0 };
      const meanValue = occurrenceCount ? values.reduce((sum, value) => sum + value, 0) / occurrenceCount : 0;
      const peerValues = ranked.slice(1).map((entry) => entry.value).filter((value) => value > 0);
      const peerAverage = peerValues.length ? average(peerValues) : 0;
      const fastestTime = occurrenceCount ? Math.min(...values) : 0;
      const p25Value = percentile(values, 0.25);
      const sortedValues = values.slice().sort((left, right) => left - right);
      const medianValue = occurrenceCount
        ? (occurrenceCount % 2
            ? sortedValues[(occurrenceCount - 1) / 2]
            : (sortedValues[occurrenceCount / 2 - 1] + sortedValues[occurrenceCount / 2]) / 2)
        : 0;
      const stdDev = standardDeviation(values);
      const cv = meanValue > 0 ? stdDev / meanValue : 0;

      let idealTime = fastestTime;
      let idealRule = "fastest";

      if (occurrenceCount >= 3 && fastestTime > 0) {
        const meanGapRatio = meanValue > 0 ? Math.abs(meanValue - fastestTime) / meanValue : 0;
        const medianGapRatio = medianValue > 0 ? Math.abs(medianValue - fastestTime) / medianValue : 0;
        if (meanGapRatio > 0.35 || medianGapRatio > 0.35) {
          idealTime = Math.min(
            ...[p25Value, medianValue, meanValue].filter((value) => Number.isFinite(value) && value > 0)
          );
          idealRule = "stable benchmark";
        }
      }

      if (!Number.isFinite(idealTime) || idealTime <= 0) {
        idealTime = fastestTime || p25Value || meanValue || medianValue || 0;
      }

      const gapToIdeal = Math.max(topEntry.value - idealTime, 0);
      const gapVsPeerAverage = peerAverage > 0 ? Math.max(topEntry.value - peerAverage, 0) : 0;
      const totalRecoverableHours = ranked.reduce((sum, entry) => sum + Math.max(entry.value - idealTime, 0), 0);
      const confidence = flatTimeConfidence(occurrenceCount, cv);
      const variability = flatTimeVariabilityLabel(cv);

      return {
        sectionSize: item.sectionSize || "__no_section__",
        groupLabel: item.groupLabel || "Unknown",
        activityLabel: item.label,
        totalTime: item.total,
        averagePerWell: occurrenceCount ? item.total / occurrenceCount : 0,
        peerAverage,
        meanValue,
        p25Value,
        medianValue,
        stdDev,
        cv,
        fastestTime,
        idealTime,
        idealRule,
        gapToIdeal,
        gapVsPeerAverage,
        totalRecoverableHours,
        confidence,
        variability,
        topEntry,
        occurrenceCount,
        values,
        ranked,
        summaryText:
          occurrenceCount >= 2
            ? occurrenceCount + " wells; peers avg " + formatNumber(peerAverage || meanValue || 0) + " hr, top well " + topEntry.label + " ran " + formatNumber(topEntry.value) + " hr"
            : "Only one well observed for this activity",
      };
    }

    function buildWellRanking(datasets, opportunities) {
      return datasets
        .map((dataset) => {
          let actualTotal = 0;
          let idealTotal = 0;
          let excessTotal = 0;
          const drivers = [];

          opportunities.forEach((opportunity) => {
            const actual = Number(opportunity.ranked.find((entry) => entry.datasetId === dataset.id)?.value || 0);
            if (!actual) return;
            actualTotal += actual;
            idealTotal += opportunity.idealTime;
            const gap = Math.max(actual - opportunity.idealTime, 0);
            excessTotal += gap;
            if (gap > 0) {
              drivers.push({
                activity: opportunity.activityLabel,
                group: opportunity.groupLabel,
                gap,
              });
            }
          });

          drivers.sort((left, right) => right.gap - left.gap || left.activity.localeCompare(right.activity));

          return {
            rigLabel: dataset.rigLabel || "Rig not mapped",
            wellLabel: dataset.subjectWell,
            actualTotal,
            idealTotal,
            excessTotal,
            topDriver: drivers[0] || null,
            topDrivers: drivers.slice(0, 3),
          };
        })
        .sort((left, right) => right.excessTotal - left.excessTotal || right.actualTotal - left.actualTotal || left.wellLabel.localeCompare(right.wellLabel));
    }

    function buildSectionBenchmarkItems(datasets, metricKey, opportunities) {
      const sectionMap = new Map();
      datasets.forEach((dataset) => {
        const totalsBySection = new Map();
        dataset.groups.forEach((group) => {
          group.activities.forEach((activity) => {
            const sectionSize = activity.sectionSize || extractFlatTimeSectionSize(activity.activity);
            if (!sectionSize || sectionSize === "__no_section__") return;
            if (!sectionMap.has(sectionSize)) sectionMap.set(sectionSize, { values: [], ideal: 0 });
            totalsBySection.set(sectionSize, (totalsBySection.get(sectionSize) || 0) + Number(activity[metricKey] || 0));
          });
        });
        totalsBySection.forEach((value, sectionSize) => {
          sectionMap.get(sectionSize).values.push(value);
        });
      });

      opportunities.forEach((opportunity) => {
        const sectionSize = opportunity.sectionSize;
        if (!sectionSize || sectionSize === "__no_section__") return;
        if (!sectionMap.has(sectionSize)) sectionMap.set(sectionSize, { values: [], ideal: 0 });
        sectionMap.get(sectionSize).ideal += Number(opportunity.idealTime || 0);
      });

      return Array.from(sectionMap.entries())
        .map(([sectionSize, bucket]) => {
          const actualAverage = average(bucket.values);
          const idealTime = Number(bucket.ideal || 0);
          return {
            label: formatFlatTimeSectionSize(sectionSize),
            sectionSize,
            actualAverage,
            idealTime,
            spread: Math.max(actualAverage - idealTime, 0),
          };
        })
        .sort((left, right) => Number(right.sectionSize) - Number(left.sectionSize) || left.label.localeCompare(right.label));
    }

    function buildRigBenchmarkSummary(datasets, opportunities) {
      const rigMap = new Map();

      datasets.forEach((dataset) => {
        const rigLabel = dataset.rigLabel || "Rig not mapped";
        if (!rigMap.has(rigLabel)) {
          rigMap.set(rigLabel, {
            rigLabel,
            datasets: [],
            gapByActivity: new Map(),
          });
        }
        rigMap.get(rigLabel).datasets.push(dataset);
      });

      return Array.from(rigMap.values())
        .map((bucket) => {
          let actualTotal = 0;
          let idealTotal = 0;
          let excessTotal = 0;
          opportunities.forEach((opportunity) => {
            bucket.datasets.forEach((dataset) => {
              const actual = Number(opportunity.ranked.find((entry) => entry.datasetId === dataset.id)?.value || 0);
              if (!actual) return;
              actualTotal += actual;
              idealTotal += opportunity.idealTime;
              const gap = Math.max(actual - opportunity.idealTime, 0);
              excessTotal += gap;
              if (gap > 0) {
                bucket.gapByActivity.set(
                  opportunity.activityLabel,
                  (bucket.gapByActivity.get(opportunity.activityLabel) || 0) + gap
                );
              }
            });
          });

          const mainRepeatingActivity = Array.from(bucket.gapByActivity.entries())
            .sort((left, right) => right[1] - left[1] || left[0].localeCompare(right[0]))[0];

          return {
            rigLabel: bucket.rigLabel,
            wellCount: bucket.datasets.length,
            averageFlatTime: bucket.datasets.length ? actualTotal / bucket.datasets.length : 0,
            averageIdealTime: bucket.datasets.length ? idealTotal / bucket.datasets.length : 0,
            excessTime: excessTotal,
            mainRepeatingActivity: mainRepeatingActivity ? mainRepeatingActivity[0] : "No repeated excess",
          };
        })
        .sort((left, right) => right.excessTime - left.excessTime || right.averageFlatTime - left.averageFlatTime || left.rigLabel.localeCompare(right.rigLabel));
    }

    function buildOpportunityPipeline(opportunities) {
      return opportunities
        .map((opportunity) => {
          const wellsImpacted = opportunity.ranked.filter((entry) => entry.value > opportunity.idealTime).length;
          let priority = "Monitor";
          if (opportunity.totalRecoverableHours >= 40 && opportunity.confidence === "High") priority = "Act now";
          else if (opportunity.totalRecoverableHours >= 20 || opportunity.confidence === "Medium") priority = "Next wave";

          return {
            activityLabel: opportunity.activityLabel,
            groupLabel: opportunity.groupLabel,
            occurrenceCount: opportunity.occurrenceCount,
            wellsImpacted,
            idealTime: opportunity.idealTime,
            totalRecoverableHours: opportunity.totalRecoverableHours,
            priority,
          };
        })
        .sort((left, right) => right.totalRecoverableHours - left.totalRecoverableHours || right.wellsImpacted - left.wellsImpacted || left.activityLabel.localeCompare(right.activityLabel));
    }

    function renderWaterfallChart(target, dataset, opportunities) {
      if (!dataset || !opportunities.length) {
        target.innerHTML = '<div class="empty">Choose a well to draw the waterfall.</div>';
        return;
      }

      const chartTheme = getChartTheme();
      const contributions = opportunities
        .map((opportunity) => {
          const actual = Number(opportunity.ranked.find((entry) => entry.datasetId === dataset.id)?.value || 0);
          return {
            label: opportunity.activityLabel,
            gap: Math.max(actual - opportunity.idealTime, 0),
          };
        })
        .filter((item) => item.gap > 0)
        .sort((left, right) => right.gap - left.gap || left.label.localeCompare(right.label));

      const topDrivers = contributions.slice(0, 5);
      const otherGap = contributions.slice(5).reduce((sum, item) => sum + item.gap, 0);
      if (otherGap > 0) topDrivers.push({ label: "Other gaps", gap: otherGap });

      const actualTotal = topDrivers.reduce((sum, item) => sum + item.gap, 0) + opportunities.reduce((sum, opportunity) => {
        const actual = Number(opportunity.ranked.find((entry) => entry.datasetId === dataset.id)?.value || 0);
        return sum + Math.min(actual, opportunity.idealTime || 0);
      }, 0);
      const idealTotal = Math.max(actualTotal - topDrivers.reduce((sum, item) => sum + item.gap, 0), 0);

      const steps = [{ label: "Actual total", start: 0, end: actualTotal, type: "total" }];
      let running = actualTotal;
      topDrivers.forEach((driver) => {
        steps.push({ label: driver.label, start: running, end: running - driver.gap, type: "delta", delta: -driver.gap });
        running -= driver.gap;
      });
      steps.push({ label: "Ideal total", start: 0, end: idealTotal, type: "total-ideal" });

      const maxValue = niceMax(Math.max(...steps.map((step) => Math.max(step.start, step.end)), 1));
      const width = Math.max(780, steps.length * 120);
      const height = 360;
      const margin = { top: 24, right: 24, bottom: 86, left: 56 };
      const chartWidth = width - margin.left - margin.right;
      const chartHeight = height - margin.top - margin.bottom;
      const barWidth = Math.min(78, chartWidth / steps.length - 18);

      const grid = Array.from({ length: 5 }, (_, index) => {
        const value = (maxValue / 4) * index;
        const y = margin.top + chartHeight - (value / maxValue) * chartHeight;
        return (
          '<g>' +
          '<line x1="' + margin.left + '" y1="' + y.toFixed(2) + '" x2="' + (width - margin.right) + '" y2="' + y.toFixed(2) + '" stroke="' + chartTheme.grid + '" stroke-dasharray="4 5"></line>' +
          '<text x="' + (margin.left - 10) + '" y="' + (y + 4).toFixed(2) + '" text-anchor="end" font-size="11" fill="' + chartTheme.valueLabel + '">' + escapeHtml(formatNumber(value)) + '</text>' +
          '</g>'
        );
      }).join("");

      const bars = steps.map((step, index) => {
        const x = margin.left + index * (chartWidth / steps.length) + ((chartWidth / steps.length) - barWidth) / 2;
        const topValue = Math.max(step.start, step.end);
        const bottomValue = Math.min(step.start, step.end);
        const y = margin.top + chartHeight - (topValue / maxValue) * chartHeight;
        const yBottom = margin.top + chartHeight - (bottomValue / maxValue) * chartHeight;
        const heightValue = Math.max(yBottom - y, 2);
        const fill = step.type === "delta" ? "#be123c" : step.type === "total-ideal" ? "#0f766e" : "#1264d6";
        const centerX = x + barWidth / 2;
        const labelLines = wrapChartLabel(step.label, 16);
        const labelSvg = labelLines.map((line, lineIndex) => '<tspan x="' + centerX.toFixed(2) + '" dy="' + (lineIndex === 0 ? 0 : 13) + '">' + escapeHtml(line) + '</tspan>').join("");
        const displayValue = step.type === "delta" ? formatNumber(Math.abs(step.delta || 0)) : formatNumber(step.end);
        return (
          '<g>' +
          '<rect x="' + x.toFixed(2) + '" y="' + y.toFixed(2) + '" width="' + barWidth.toFixed(2) + '" height="' + heightValue.toFixed(2) + '" rx="10" fill="' + fill + '"></rect>' +
          '<text x="' + centerX.toFixed(2) + '" y="' + Math.max(margin.top + 12, y - 8).toFixed(2) + '" text-anchor="middle" font-size="11" font-weight="700" fill="' + chartTheme.pointLabel + '">' + escapeHtml(displayValue) + '</text>' +
          '<text x="' + centerX.toFixed(2) + '" y="' + (height - 42) + '" text-anchor="middle" font-size="11" font-weight="700" fill="' + chartTheme.text + '">' + labelSvg + '</text>' +
          '</g>'
        );
      }).join("");

      target.innerHTML =
        '<div class="legend" style="margin-bottom:10px;">' +
        '<span class="legend-item"><span class="legend-dot" style="background:#1264d6;"></span>' + escapeHtml(dataset.subjectWell + " actual") + '</span>' +
        '<span class="legend-item"><span class="legend-dot" style="background:#be123c;"></span>Recoverable gap</span>' +
        '<span class="legend-item"><span class="legend-dot" style="background:#0f766e;"></span>Recommended ideal</span>' +
        '</div>' +
        '<div class="column-chart-wrap">' +
        '<svg class="column-chart-svg" viewBox="0 0 ' + width + ' ' + height + '" role="img" aria-label="Well versus ideal waterfall chart">' +
        grid +
        bars +
        '<line x1="' + margin.left + '" y1="' + (margin.top + chartHeight).toFixed(2) + '" x2="' + (width - margin.right) + '" y2="' + (margin.top + chartHeight).toFixed(2) + '" stroke="' + chartTheme.axis + '"></line>' +
        '</svg></div>';
    }

    function renderVariabilityChart(target, opportunities, topN) {
      const items = opportunities
        .filter((opportunity) => opportunity.occurrenceCount >= 2)
        .slice()
        .sort((left, right) => right.cv - left.cv || right.totalTime - left.totalTime)
        .slice(0, topN);

      if (!items.length) {
        target.innerHTML = '<div class="empty">Need at least two wells per activity to show variability.</div>';
        return;
      }

      const chartTheme = getChartTheme();
      const width = 920;
      const height = Math.max(340, items.length * 48 + 70);
      const margin = { top: 26, right: 28, bottom: 34, left: 190 };
      const chartWidth = width - margin.left - margin.right;
      const chartHeight = height - margin.top - margin.bottom;
      const maxValue = niceMax(Math.max(...items.map((item) => Math.max(...item.values, item.idealTime)), 1));
      const rowHeight = chartHeight / items.length;

      const grid = Array.from({ length: 6 }, (_, index) => {
        const value = (maxValue / 5) * index;
        const x = margin.left + (value / maxValue) * chartWidth;
        return (
          '<g>' +
          '<line x1="' + x.toFixed(2) + '" y1="' + margin.top + '" x2="' + x.toFixed(2) + '" y2="' + (height - margin.bottom) + '" stroke="' + chartTheme.grid + '" stroke-dasharray="4 5"></line>' +
          '<text x="' + x.toFixed(2) + '" y="' + (height - 10) + '" text-anchor="middle" font-size="11" fill="' + chartTheme.valueLabel + '">' + escapeHtml(formatNumber(value)) + '</text>' +
          '</g>'
        );
      }).join("");

      const rows = items.map((item, index) => {
        const y = margin.top + index * rowHeight + rowHeight / 2;
        const minValue = Math.min(...item.values);
        const maxObserved = Math.max(...item.values);
        const q1 = percentile(item.values, 0.25);
        const q3 = percentile(item.values, 0.75);
        const xMin = margin.left + (minValue / maxValue) * chartWidth;
        const xQ1 = margin.left + (q1 / maxValue) * chartWidth;
        const xMedian = margin.left + (item.medianValue / maxValue) * chartWidth;
        const xQ3 = margin.left + (q3 / maxValue) * chartWidth;
        const xMax = margin.left + (maxObserved / maxValue) * chartWidth;
        const xIdeal = margin.left + (item.idealTime / maxValue) * chartWidth;
        return (
          '<g>' +
          '<line x1="' + xMin.toFixed(2) + '" y1="' + y.toFixed(2) + '" x2="' + xMax.toFixed(2) + '" y2="' + y.toFixed(2) + '" stroke="' + chartTheme.axis + '" stroke-width="2"></line>' +
          '<rect x="' + xQ1.toFixed(2) + '" y="' + (y - 10).toFixed(2) + '" width="' + Math.max(xQ3 - xQ1, 2).toFixed(2) + '" height="20" rx="8" fill="rgba(18, 100, 214, 0.25)" stroke="#1264d6"></rect>' +
          '<line x1="' + xMedian.toFixed(2) + '" y1="' + (y - 12).toFixed(2) + '" x2="' + xMedian.toFixed(2) + '" y2="' + (y + 12).toFixed(2) + '" stroke="#1264d6" stroke-width="3"></line>' +
          '<circle cx="' + xIdeal.toFixed(2) + '" cy="' + y.toFixed(2) + '" r="5" fill="#c06a0a"></circle>' +
          '<text x="' + (margin.left - 12) + '" y="' + (y + 4).toFixed(2) + '" text-anchor="end" font-size="11" font-weight="700" fill="' + chartTheme.text + '">' + escapeHtml(item.activityLabel) + '</text>' +
          '</g>'
        );
      }).join("");

      target.innerHTML =
        '<div class="legend" style="margin-bottom:10px;">' +
        '<span class="legend-item"><span class="legend-dot" style="background:#1264d6;"></span>Interquartile range / median</span>' +
        '<span class="legend-item"><span class="legend-dot" style="background:#c06a0a;"></span>Recommended ideal</span>' +
        '</div>' +
        '<div class="column-chart-wrap">' +
        '<svg class="column-chart-svg" viewBox="0 0 ' + width + ' ' + height + '" role="img" aria-label="Variability box plot chart">' +
        grid +
        rows +
        '</svg></div>';
    }

    function wireFlatTimeFocusActions() {
      Array.from(document.querySelectorAll("[data-flat-focus-well]")).forEach((button) => {
        button.addEventListener("click", () => {
          flatTimeState.focusWell = button.dataset.flatFocusWell || "";
          renderFlatTime();
        });
      });

      Array.from(document.querySelectorAll("[data-flat-focus-activity]")).forEach((button) => {
        button.addEventListener("click", () => {
          flatTimeState.focusActivity = button.dataset.flatFocusActivity || "";
          renderFlatTime();
        });
      });
    }

    function renderFlatTimeDrilldown(datasets, opportunities, selectedWell, selectedActivity) {
      const selectedDataset = datasets.find((dataset) => dataset.subjectWell === selectedWell) || datasets[0];
      const selectedOpportunity = opportunities.find((opportunity) => opportunity.activityLabel === selectedActivity) || opportunities[0];

      if (!selectedDataset || !selectedOpportunity) {
        ui.flatTimeWellDrilldown.innerHTML = '<div class="empty">Select more CSVs to unlock the drill-down.</div>';
        ui.flatTimeActivityDrilldown.innerHTML = '<div class="empty">Select more CSVs to unlock the activity benchmark.</div>';
        return;
      }

      const allWellDrivers = opportunities
        .map((opportunity) => {
          const actual = Number(opportunity.ranked.find((entry) => entry.datasetId === selectedDataset.id)?.value || 0);
          const gap = Math.max(actual - opportunity.idealTime, 0);
          return {
            activityLabel: opportunity.activityLabel,
            groupLabel: opportunity.groupLabel,
            actual,
            idealTime: opportunity.idealTime,
            gap,
            peerAverage: opportunity.peerAverage || opportunity.meanValue || 0,
          };
        })
        .filter((item) => item.actual > 0)
        .sort((left, right) => right.gap - left.gap || right.actual - left.actual || left.activityLabel.localeCompare(right.activityLabel));

      const wellDrivers = allWellDrivers.slice(0, 6);

      const wellActualTotal = allWellDrivers.reduce((sum, item) => sum + item.actual, 0);
      const wellIdealTotal = allWellDrivers.reduce((sum, item) => sum + item.idealTime, 0);
      const wellExcess = allWellDrivers.reduce((sum, item) => sum + item.gap, 0);

      ui.flatTimeWellDrilldown.innerHTML =
        '<div class="metric-strip" style="margin-bottom:14px;">' +
        '<div class="metric-pill"><div class="label">Selected Well</div><div class="value"><span class="value-main">' + escapeHtml(selectedDataset.subjectWell) + '</span></div><div class="meta">' + escapeHtml(selectedDataset.rigLabel || "Rig not mapped") + '</div></div>' +
        '<div class="metric-pill"><div class="label">Actual Flat Time</div><div class="value"><span class="value-main">' + escapeHtml(formatNumber(wellActualTotal)) + '</span><span class="value-suffix">' + escapeHtml("hr / " + formatNumber(wellActualTotal / 24) + " d") + '</span></div><div class="meta">All activities in the selected filter context</div></div>' +
        '<div class="metric-pill"><div class="label">Ideal Flat Time</div><div class="value"><span class="value-main">' + escapeHtml(formatNumber(wellIdealTotal)) + '</span><span class="value-suffix">' + escapeHtml("hr / " + formatNumber(wellIdealTotal / 24) + " d") + '</span></div><div class="meta">Recommended achievable total for the same activities</div></div>' +
        '<div class="metric-pill"><div class="label">Recoverable Gap</div><div class="value"><span class="value-main">' + escapeHtml(formatNumber(wellExcess)) + '</span><span class="value-suffix">' + escapeHtml("hr / " + formatNumber(wellExcess / 24) + " d") + '</span></div><div class="meta">Time above the recommended ideal</div></div>' +
        '</div>' +
        '<div class="drill-list">' +
        wellDrivers.map((item) =>
          '<div class="drill-item">' +
          '<strong>' + flatTimeActionButtonHtml("activity", item.activityLabel, item.activityLabel) + '</strong>' +
          '<span>' + escapeHtml(item.groupLabel + " • actual " + formatHoursWithDays(item.actual) + " • peers avg " + formatHoursWithDays(item.peerAverage) + " • ideal " + formatHoursWithDays(item.idealTime) + " • gap " + formatHoursWithDays(item.gap)) + '</span>' +
          '</div>'
        ).join('') +
        '</div>';

      const peerRows = selectedOpportunity.ranked
        .slice()
        .sort((left, right) => right.value - left.value || left.label.localeCompare(right.label))
        .map((entry) => ({
          rigLabel: entry.rigLabel || "Rig not mapped",
          label: entry.label,
          value: entry.value,
          gap: Math.max(entry.value - selectedOpportunity.idealTime, 0),
        }));

      ui.flatTimeActivityDrilldown.innerHTML =
        '<div class="metric-strip" style="margin-bottom:14px;">' +
        '<div class="metric-pill"><div class="label">Selected Activity</div><div class="value"><span class="value-main">' + flatTimeActivityLabelHtml(selectedOpportunity.activityLabel) + '</span></div><div class="meta">' + escapeHtml(selectedOpportunity.groupLabel + " • " + formatFlatTimeSectionSize(selectedOpportunity.sectionSize)) + '</div></div>' +
        '<div class="metric-pill"><div class="label">Recommended Ideal</div><div class="value"><span class="value-main">' + escapeHtml(formatNumber(selectedOpportunity.idealTime)) + '</span><span class="value-suffix">' + escapeHtml("hr / " + formatNumber(selectedOpportunity.idealTime / 24) + " d") + '</span></div><div class="meta">' + escapeHtml(selectedOpportunity.idealRule) + '</div></div>' +
        '<div class="metric-pill"><div class="label">Confidence</div><div class="value">' + confidenceBadgeHtml(selectedOpportunity.confidence) + '</div><div class="meta">' + escapeHtml(selectedOpportunity.variability + " variability • sample " + selectedOpportunity.occurrenceCount) + '</div></div>' +
        '<div class="metric-pill"><div class="label">Observed Range</div><div class="value"><span class="value-main">' + escapeHtml(formatNumber(selectedOpportunity.fastestTime)) + " - " + escapeHtml(formatNumber(Math.max(...selectedOpportunity.values, 0))) + '</span><span class="value-suffix">' + escapeHtml("hr / " + formatNumber(selectedOpportunity.fastestTime / 24) + " - " + formatNumber(Math.max(...selectedOpportunity.values, 0) / 24) + " d") + '</span></div><div class="meta">Fastest to slowest observed execution</div></div>' +
        '</div>' +
        '<div class="table-wrap"><table><thead><tr><th>Rig</th><th>Well</th><th>Observed Time (hr)</th><th>Gap vs Ideal (hr)</th></tr></thead><tbody>' +
        peerRows.map((row) =>
          '<tr>' +
          '<td>' + escapeHtml(row.rigLabel) + '</td>' +
          '<td>' + flatTimeActionButtonHtml("well", row.label, row.label) + '</td>' +
          '<td>' + escapeHtml(formatHoursWithDays(row.value)) + '</td>' +
          '<td>' + escapeHtml(formatHoursWithDays(row.gap)) + '</td>' +
          '</tr>'
        ).join('') +
        '</tbody></table></div>';

      ui.flatTimeDrilldownNote.textContent =
        'Selected well: ' + selectedDataset.subjectWell + ' • selected activity: ' + selectedOpportunity.activityLabel + '. ' +
        'Headline well totals are now calculated from all activities in the selected filter context, while the list below keeps only the top loss drivers. ' +
        'Depth-based drill-down is still limited because the uploaded CSVs do not contain true depth fields such as section top/bottom, measured depth or TD.';
    }

    function renderParetoChart(target, opportunities) {
      if (!opportunities.length) {
        target.innerHTML = '<div class="empty">No recoverable-hour opportunities available.</div>';
        return;
      }

      const items = opportunities
        .slice()
        .sort((left, right) => right.totalRecoverableHours - left.totalRecoverableHours)
        .slice(0, 10);
      const totalRecoverable = items.reduce((sum, item) => sum + item.totalRecoverableHours, 0) || 1;
      const chartTheme = getChartTheme();
      const width = Math.max(820, items.length * 120);
      const height = 360;
      const margin = { top: 24, right: 40, bottom: 86, left: 52 };
      const chartWidth = width - margin.left - margin.right;
      const chartHeight = height - margin.top - margin.bottom;
      const maxBar = niceMax(Math.max(...items.map((item) => item.totalRecoverableHours), 1));
      const groupWidth = chartWidth / items.length;
      let cumulative = 0;
      const points = [];

      const grid = Array.from({ length: 5 }, (_, index) => {
        const value = (maxBar / 4) * index;
        const y = margin.top + chartHeight - (value / maxBar) * chartHeight;
        return (
          '<g>' +
          '<line x1="' + margin.left + '" y1="' + y.toFixed(2) + '" x2="' + (width - margin.right) + '" y2="' + y.toFixed(2) + '" stroke="' + chartTheme.grid + '" stroke-dasharray="4 5"></line>' +
          '<text x="' + (margin.left - 10) + '" y="' + (y + 4).toFixed(2) + '" text-anchor="end" font-size="11" fill="' + chartTheme.valueLabel + '">' + escapeHtml(formatNumber(value)) + "</text>" +
          "</g>"
        );
      }).join("");

      const bars = items.map((item, index) => {
        const x = margin.left + index * groupWidth + groupWidth * 0.18;
        const barWidth = groupWidth * 0.64;
        const barHeight = (item.totalRecoverableHours / maxBar) * chartHeight;
        const y = margin.top + chartHeight - barHeight;
        cumulative += item.totalRecoverableHours;
        const cumulativePct = (cumulative / totalRecoverable) * 100;
        const pointX = x + barWidth / 2;
        const pointY = margin.top + chartHeight - (cumulativePct / 100) * chartHeight;
        points.push({ x: pointX, y: pointY, pct: cumulativePct });
        const labelLines = wrapChartLabel(item.activityLabel, 16);
        const labelSvg = labelLines
          .map((line, lineIndex) => '<tspan x="' + pointX.toFixed(2) + '" dy="' + (lineIndex === 0 ? 0 : 13) + '">' + escapeHtml(line) + "</tspan>")
          .join("");
        return (
          '<g>' +
          '<rect x="' + x.toFixed(2) + '" y="' + y.toFixed(2) + '" width="' + barWidth.toFixed(2) + '" height="' + Math.max(barHeight, 2).toFixed(2) + '" rx="10" fill="#1264d6"></rect>' +
          '<text x="' + pointX.toFixed(2) + '" y="' + Math.max(margin.top + 12, y - 8).toFixed(2) + '" text-anchor="middle" font-size="11" font-weight="700" fill="' + chartTheme.pointLabel + '">' + escapeHtml(formatNumber(item.totalRecoverableHours)) + "</text>" +
          '<text x="' + pointX.toFixed(2) + '" y="' + (height - 42) + '" text-anchor="middle" font-size="11" font-weight="700" fill="' + chartTheme.text + '">' + labelSvg + "</text>" +
          "</g>"
        );
      }).join("");

      const linePath = points.map((point, index) => (index === 0 ? "M" : "L") + point.x.toFixed(2) + " " + point.y.toFixed(2)).join(" ");
      const lineDots = points
        .map((point) => (
          '<g>' +
          '<circle cx="' + point.x.toFixed(2) + '" cy="' + point.y.toFixed(2) + '" r="4" fill="#c06a0a"></circle>' +
          '<text x="' + point.x.toFixed(2) + '" y="' + (point.y - 10).toFixed(2) + '" text-anchor="middle" font-size="10" fill="' + chartTheme.pointLabel + '">' + escapeHtml(formatNumber(point.pct) + "%") + "</text>" +
          '</g>'
        ))
        .join("");

      target.innerHTML =
        '<div class="column-chart-wrap">' +
        '<svg class="column-chart-svg" viewBox="0 0 ' + width + ' ' + height + '" role="img" aria-label="Pareto recoverable hours chart">' +
        grid +
        bars +
        '<path d="' + linePath + '" fill="none" stroke="#c06a0a" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"></path>' +
        lineDots +
        '<line x1="' + margin.left + '" y1="' + (margin.top + chartHeight).toFixed(2) + '" x2="' + (width - margin.right) + '" y2="' + (margin.top + chartHeight).toFixed(2) + '" stroke="' + chartTheme.axis + '"></line>' +
        '</svg></div>';
    }

    function renderHeatmap(target, datasets, opportunities, topN) {
      if (!datasets.length || !opportunities.length) {
        target.innerHTML = '<div class="empty">No heatmap data available.</div>';
        return;
      }

      const rows = datasets.slice();
      const columns = opportunities
        .slice()
        .sort((left, right) => right.totalRecoverableHours - left.totalRecoverableHours)
        .slice(0, Math.max(6, topN));
      const maxGap = Math.max(
        ...rows.flatMap((dataset) =>
          columns.map((opportunity) => {
            const actual = Number(opportunity.ranked.find((entry) => entry.datasetId === dataset.id)?.value || 0);
            return Math.max(actual - opportunity.idealTime, 0);
          })
        ),
        0
      );

      const headerHtml =
        '<tr><th>Well</th>' +
        columns.map((column) => '<th title="' + escapeHtml(getFlatTimeActivityTooltip(column.activityLabel) || column.activityLabel) + '">' + escapeHtml(column.activityLabel) + '</th>').join("") +
        '</tr>';
      const bodyHtml = rows.map((dataset) => {
        const cells = columns.map((opportunity) => {
          const actual = Number(opportunity.ranked.find((entry) => entry.datasetId === dataset.id)?.value || 0);
          const gap = Math.max(actual - opportunity.idealTime, 0);
          const opacity = maxGap > 0 ? Math.max(0.12, gap / maxGap) : 0;
          const bg = gap > 0 ? 'rgba(200, 30, 90, ' + opacity.toFixed(2) + ')' : 'rgba(18, 100, 214, 0.06)';
          const color = gap > 0.01 ? '#ffffff' : 'var(--ink)';
          return '<td style="background:' + bg + '; color:' + color + '; font-weight:700; text-align:center;">' + escapeHtml(formatNumber(gap)) + '</td>';
        }).join("");
        return '<tr><td><strong>' + escapeHtml(dataset.subjectWell) + '</strong><br><span style="color:var(--muted); font-size:12px;">' + escapeHtml(dataset.rigLabel || '') + '</span></td>' + cells + '</tr>';
      }).join("");

      target.innerHTML = '<div class="table-wrap"><table><thead>' + headerHtml + '</thead><tbody>' + bodyHtml + '</tbody></table></div>';
    }

    function renderFlatTimeDatasetTags(datasets) {
      if (!datasets.length) {
        ui.flatTimeDatasetTags.innerHTML = '<span class="tag tag-muted">No flat time CSV datasets loaded yet.</span>';
        return;
      }

      ui.flatTimeDatasetTags.innerHTML = datasets
        .map((dataset) => {
          const removable = flatTimeState.uploadedDatasets.some((item) => item.id === dataset.id);
          return (
            '<span class="tag">' +
            '<strong>' + escapeHtml(dataset.subjectWell) + "</strong>" +
            '<span class="tag-muted">' + escapeHtml(dataset.rigLabel || "Rig not mapped") + " • " + escapeHtml(dataset.fileName) + "</span>" +
            (removable ? '<button type="button" data-flat-time-remove="' + escapeHtml(dataset.id) + '">Remove</button>' : "") +
            "</span>"
          );
        })
        .join("");

      Array.from(ui.flatTimeDatasetTags.querySelectorAll("[data-flat-time-remove]")).forEach((button) => {
        button.addEventListener("click", () => {
          flatTimeState.uploadedDatasets = flatTimeState.uploadedDatasets.filter((item) => item.id !== button.dataset.flatTimeRemove);
          renderFlatTime();
        });
      });
    }

    function renderFlatTimeSummary(datasets, metricKey, totalKey, activityItems, groupItems) {
      if (!datasets.length) {
        ui.flatTimeSummary.innerHTML = '<div class="empty">Upload flat time CSV files to start the comparison.</div>';
        return;
      }

      const topActivity = activityItems[0];
      const topGroup = groupItems[0];
      const highestDataset = datasets
        .map((dataset) => ({ label: dataset.subjectWell, rigLabel: dataset.rigLabel || "Rig not mapped", value: Number(dataset[totalKey] || 0) }))
        .sort((left, right) => right.value - left.value)[0];
      const opportunities = activityItems.map((item) => computeFlatTimeOpportunity(item, datasets));
      const topSpread = opportunities
        .slice()
        .sort((left, right) => right.gapToIdeal - left.gapToIdeal || right.gapVsPeerAverage - left.gapVsPeerAverage)[0];
      const totalRecoverable = opportunities.reduce((sum, opportunity) => sum + opportunity.totalRecoverableHours, 0);
      const mostReliableIdeal = opportunities
        .filter((opportunity) => opportunity.occurrenceCount >= 3)
        .slice()
        .sort((left, right) => {
          const confidenceScore = { High: 3, Medium: 2, Low: 1 };
          return (
            (confidenceScore[right.confidence] || 0) - (confidenceScore[left.confidence] || 0) ||
            left.cv - right.cv ||
            right.occurrenceCount - left.occurrenceCount
          );
        })[0];
      const overallHours = datasets.reduce((sum, dataset) => sum + Number(dataset[totalKey] || 0), 0);

      const cards = [
        { label: "Benchmarks Compared", value: String(datasets.length), meta: datasets.map((dataset) => dataset.subjectWell).join(", ") },
        { label: "Top Consuming Activity", value: topActivity ? topActivity.label : "N/A", valueHtml: topActivity ? flatTimeActivityLabelHtml(topActivity.label) : escapeHtml("N/A"), meta: topActivity ? formatNumber(topActivity.total) + " hr total" : "No activity data" },
        { label: "Largest Group", value: topGroup ? topGroup.label : "N/A", meta: topGroup ? formatNumber(topGroup.total) + " hr total" : "No group data" },
        { label: "Highest Burden Well", value: highestDataset ? highestDataset.label : "N/A", meta: highestDataset ? (highestDataset.rigLabel + " • " + formatNumber(highestDataset.value) + " hr total") : "No dataset totals" },
        {
          label: "Best Reduction Opportunity",
          value: topSpread ? topSpread.topEntry.rigLabel : "N/A",
          meta: topSpread ? (topSpread.topEntry.label + " • " + topSpread.activityLabel + " • gap " + formatNumber(topSpread.gapToIdeal) + " hr vs ideal") : "Need more than one dataset",
          metaHtml: topSpread ? (escapeHtml(topSpread.topEntry.label + " • ") + flatTimeActivityLabelHtml(topSpread.activityLabel) + escapeHtml(" • gap " + formatNumber(topSpread.gapToIdeal) + " hr vs ideal")) : escapeHtml("Need more than one dataset"),
        },
        {
          label: "Total Recoverable Hours",
          value: formatNumber(totalRecoverable),
          meta: "Sum of time above the recommended ideal across the comparison set",
        },
        {
          label: "Most Reliable Ideal",
          value: mostReliableIdeal ? mostReliableIdeal.activityLabel : "N/A",
          valueHtml: mostReliableIdeal ? flatTimeActivityLabelHtml(mostReliableIdeal.activityLabel) : escapeHtml("N/A"),
          meta: mostReliableIdeal ? (mostReliableIdeal.confidence + " confidence • target " + formatNumber(mostReliableIdeal.idealTime) + " hr") : "Need at least 3 wells for a strong benchmark",
        },
        { label: "Total Compared Time", value: formatNumber(overallHours), meta: metricKey === "subjectHours" ? "Subject well hours" : metricKey === "meanHours" ? "Mean hours" : "Median hours" },
      ];

      ui.flatTimeSummary.innerHTML = cards
        .map((card) => (
          '<div class="metric-pill">' +
          '<div class="label">' + escapeHtml(card.label) + "</div>" +
          '<div class="value"><span class="value-main">' + (card.valueHtml || escapeHtml(card.value)) + '</span></div>' +
          '<div class="meta">' + (card.metaHtml || escapeHtml(card.meta)) + "</div>" +
          "</div>"
        ))
        .join("");
    }

    function renderFlatTime() {
      const allDatasets = getFlatTimeDatasets();
      const metricKey = getFlatTimeMetricKey();
      const totalKey = getFlatTimeTotalKey();
      const topN = Number(ui.flatTimeTopN.value || 10);
      const selectedRig = ui.flatTimeRig.value || "";
      const selectedSectionSize = ui.flatTimeSection.value || "";

      renderFlatTimeDatasetTags(allDatasets);
      populateFlatTimeRigOptions(allDatasets);
      updateFlatTimeModeVisibility();

      if (!allDatasets.length) {
        ui.flatTimeTitle.textContent = "No flat time datasets loaded";
        ui.flatTimeSubtitle.textContent = "Use the CSV uploader to compare benchmark files.";
        ui.flatTimeSummary.innerHTML = '<div class="empty">Upload flat time CSV files to start the comparison.</div>';
        ui.flatTimeWellRanking.innerHTML = '<div class="empty">Upload flat time CSV files to rank wells by excess time.</div>';
        ui.flatTimeParetoChart.innerHTML = '<div class="empty">Upload flat time CSV files to build a Pareto of recoverable hours.</div>';
        ui.flatTimeWaterfallChart.innerHTML = '<div class="empty">Upload flat time CSV files to build the waterfall.</div>';
        ui.flatTimeSectionBenchmarkChart.innerHTML = '<div class="empty">Upload flat time CSV files to compare sections.</div>';
        ui.flatTimeRigSummary.innerHTML = '<div class="empty">Upload flat time CSV files to summarize rigs.</div>';
        ui.flatTimeOpportunityPipeline.innerHTML = '<div class="empty">Upload flat time CSV files to build the opportunity pipeline.</div>';
        ui.flatTimeGroupChart.innerHTML = '<div class="empty">No flat time group data available.</div>';
        ui.flatTimeActivityChart.innerHTML = '<div class="empty">No flat time activity data available.</div>';
        ui.flatTimeBenchmarkTable.innerHTML = '<div class="empty">No activity benchmark table available.</div>';
        ui.flatTimeWellDrilldown.innerHTML = '<div class="empty">Upload flat time CSV files to inspect a well.</div>';
        ui.flatTimeActivityDrilldown.innerHTML = '<div class="empty">Upload flat time CSV files to inspect an activity.</div>';
        ui.flatTimeDrilldownNote.textContent = "Click a well or activity in the tables above to open the benchmark, peer comparison and ideal-time logic.";
        ui.flatTimeOpportunityTable.innerHTML = '<div class="empty">No flat time comparison table available.</div>';
        ui.flatTimeGroupTable.innerHTML = '<div class="empty">No flat time group table available.</div>';
        ui.flatTimeLossDrivers.innerHTML = '<div class="empty">Upload flat time CSV files to list top loss drivers by well.</div>';
        ui.flatTimeVariabilityChart.innerHTML = '<div class="empty">Upload flat time CSV files to show variability.</div>';
        ui.flatTimeHeatmap.innerHTML = '<div class="empty">No heatmap data available.</div>';
        ui.flatTimePerfectChart.innerHTML = '<div class="empty">Upload flat time CSV files to draw the perfect flat time curve.</div>';
        populateFlatTimeWellOptions([], "");
        return;
      }

      const rigDatasets = filterFlatTimeDatasetsByRig(allDatasets, selectedRig);
      populateFlatTimeSectionOptions(rigDatasets);
      const datasets = annotateFlatTimeScopedBenchmarks(
        filterFlatTimeDatasetsBySection(rigDatasets, selectedSectionSize)
      );

      if (!datasets.length) {
        ui.flatTimeTitle.textContent = "No data for selected section size";
        ui.flatTimeSubtitle.textContent = "Try another rig or section size, or switch back to all sections.";
        ui.flatTimeSummary.innerHTML = '<div class="empty">No benchmark activities match the selected section size.</div>';
        ui.flatTimeWellRanking.innerHTML = '<div class="empty">No well ranking available for this section size.</div>';
        ui.flatTimeParetoChart.innerHTML = '<div class="empty">No Pareto data available for this section size.</div>';
        ui.flatTimeWaterfallChart.innerHTML = '<div class="empty">No waterfall available for this section size.</div>';
        ui.flatTimeSectionBenchmarkChart.innerHTML = '<div class="empty">No section benchmark available for this section size.</div>';
        ui.flatTimeRigSummary.innerHTML = '<div class="empty">No rig benchmark summary available for this section size.</div>';
        ui.flatTimeOpportunityPipeline.innerHTML = '<div class="empty">No opportunity pipeline available for this section size.</div>';
        ui.flatTimeGroupChart.innerHTML = '<div class="empty">No flat time group data available for this section size.</div>';
        ui.flatTimeActivityChart.innerHTML = '<div class="empty">No flat time activity data available for this section size.</div>';
        ui.flatTimeBenchmarkTable.innerHTML = '<div class="empty">No benchmark table available for this section size.</div>';
        ui.flatTimeWellDrilldown.innerHTML = '<div class="empty">No well drill-down available for this section size.</div>';
        ui.flatTimeActivityDrilldown.innerHTML = '<div class="empty">No activity drill-down available for this section size.</div>';
        ui.flatTimeDrilldownNote.textContent = "Click a well or activity in the tables above to open the benchmark, peer comparison and ideal-time logic.";
        ui.flatTimeOpportunityTable.innerHTML = '<div class="empty">No flat time comparison table available for this section size.</div>';
        ui.flatTimeGroupTable.innerHTML = '<div class="empty">No flat time group table available for this section size.</div>';
        ui.flatTimeLossDrivers.innerHTML = '<div class="empty">No loss driver ranking available for this section size.</div>';
        ui.flatTimeVariabilityChart.innerHTML = '<div class="empty">No variability view available for this section size.</div>';
        ui.flatTimeHeatmap.innerHTML = '<div class="empty">No heatmap data available for this section size.</div>';
        ui.flatTimePerfectChart.innerHTML = '<div class="empty">No section-sized activities available for the perfect flat time curve.</div>';
        populateFlatTimeWellOptions([], "");
        return;
      }

      const groupItems = buildFlatTimeGroupItems(datasets, totalKey);
      const activityItems = buildFlatTimeActivityItems(datasets, metricKey);
      const sectionLabel = selectedSectionSize ? formatFlatTimeSectionSize(selectedSectionSize) : "All section sizes";

      ui.flatTimeTitle.textContent = datasets.length + " wells compared";
      ui.flatTimeSubtitle.textContent = (selectedRig || "All rigs") + " • " + sectionLabel + " • " + datasets.map((dataset) => dataset.subjectWell).join(" vs ");

      renderFlatTimeSummary(datasets, metricKey, totalKey, activityItems, groupItems);

      const seriesDefs = datasets.map((dataset, index) => ({
        key: dataset.id,
        label: dataset.subjectWell,
        color: FLAT_TIME_SERIES_COLORS[index % FLAT_TIME_SERIES_COLORS.length],
        format: (value) => formatNumber(value),
      }));

      renderMultiSeriesChart(ui.flatTimeGroupChart, groupItems.slice(0, 8), seriesDefs, {
        height: 430,
        minWidth: 880,
        groupMinWidth: 180,
      });
      renderMultiSeriesChart(ui.flatTimeActivityChart, activityItems.slice(0, topN), seriesDefs, {
        height: 430,
        minWidth: 880,
        groupMinWidth: 180,
      });

      const allOpportunities = activityItems
        .map((item) => computeFlatTimeOpportunity(item, datasets))
        .sort((left, right) => right.gapToIdeal - left.gapToIdeal || right.totalRecoverableHours - left.totalRecoverableHours || right.totalTime - left.totalTime);
      const rankedOpportunities = allOpportunities.slice(0, topN);
      const wellRanking = buildWellRanking(datasets, allOpportunities);
      const worstWell = wellRanking[0] ? wellRanking[0].wellLabel : "";
      const topActivityFocus = rankedOpportunities[0] ? rankedOpportunities[0].activityLabel : "";
      if (!flatTimeState.focusWell || !datasets.some((dataset) => dataset.subjectWell === flatTimeState.focusWell)) {
        flatTimeState.focusWell = worstWell;
      }
      if (!flatTimeState.focusActivity || !allOpportunities.some((item) => item.activityLabel === flatTimeState.focusActivity)) {
        flatTimeState.focusActivity = topActivityFocus;
      }
      populateFlatTimeWellOptions(datasets, flatTimeState.focusWell || worstWell);
      const selectedWell = ui.flatTimeWell.value || flatTimeState.focusWell || worstWell;
      const selectedDataset = datasets.find((dataset) => dataset.subjectWell === selectedWell) || datasets[0];
      const selectedActivity = flatTimeState.focusActivity || topActivityFocus;
      const sectionBenchmarkItems = buildSectionBenchmarkItems(datasets, metricKey, allOpportunities);
      const rigBenchmarkRows = buildRigBenchmarkSummary(datasets, allOpportunities);
      const opportunityPipeline = buildOpportunityPipeline(allOpportunities).slice(0, Math.max(topN, 8));

      renderTableHtml(
        ui.flatTimeWellRanking,
        ["Rig", "Well", "Actual Total (hr)", "Ideal Total (hr)", "Excess Time (hr)", "Top Drivers"],
        wellRanking.map((row) => [
          escapeHtml(row.rigLabel),
          flatTimeActionButtonHtml("well", row.wellLabel, row.wellLabel),
          escapeHtml(formatNumber(row.actualTotal)),
          escapeHtml(formatNumber(row.idealTotal)),
          flatTimeTrendHtml(row.excessTotal),
          row.topDrivers.length
            ? row.topDrivers.map((driver) => flatTimeActivityLabelHtml(driver.activity) + escapeHtml(" (" + driver.group + ", +" + formatNumber(driver.gap) + " hr)")).join("<br>")
            : escapeHtml("No excess detected"),
        ])
      );

      renderParetoChart(ui.flatTimeParetoChart, rankedOpportunities);
      renderWaterfallChart(ui.flatTimeWaterfallChart, selectedDataset, allOpportunities);
      renderMultiSeriesChart(
        ui.flatTimeSectionBenchmarkChart,
        sectionBenchmarkItems,
        [
          { key: "actualAverage", label: "Actual avg", color: "#1264d6", format: (value) => formatNumber(value) },
          { key: "idealTime", label: "Recommended ideal", color: "#0f766e", format: (value) => formatNumber(value) },
          { key: "spread", label: "Spread", color: "#c06a0a", format: (value) => formatNumber(value) },
        ],
        {
          height: 420,
          minWidth: 760,
          groupMinWidth: 150,
        }
      );

      renderTable(
        ui.flatTimeRigSummary,
        ["Rig", "# Wells", "Avg Flat Time (hr)", "Ideal Flat Time (hr)", "Excess Time (hr)", "Main Repeating Activity"],
        rigBenchmarkRows.map((row) => [
          row.rigLabel,
          String(row.wellCount),
          formatNumber(row.averageFlatTime),
          formatNumber(row.averageIdealTime),
          formatNumber(row.excessTime),
          row.mainRepeatingActivity,
        ])
      );

      renderTableHtml(
        ui.flatTimeOpportunityPipeline,
        ["Activity", "Group", "Occurrences", "Wells Impacted", "Ideal Time (hr)", "Recoverable Hours", "Priority"],
        opportunityPipeline.map((row) => [
          flatTimeActionButtonHtml("activity", row.activityLabel, row.activityLabel),
          escapeHtml(row.groupLabel),
          escapeHtml(String(row.occurrenceCount)),
          escapeHtml(String(row.wellsImpacted)),
          escapeHtml(formatNumber(row.idealTime)),
          escapeHtml(formatNumber(row.totalRecoverableHours)),
          escapeHtml(row.priority),
        ])
      );

      renderTableHtml(
        ui.flatTimeBenchmarkTable,
        ["Section", "Group", "Activity", "Sample", "Fastest", "P25", "Median", "Mean", "Recommended Ideal", "Variability", "Confidence", "Recoverable Hours", "Highest Well"],
        allOpportunities
          .slice(0, Math.max(topN * 2, 20))
          .map((opportunity) => [
            escapeHtml(formatFlatTimeSectionSize(opportunity.sectionSize)),
            escapeHtml(opportunity.groupLabel),
            flatTimeActionButtonHtml("activity", opportunity.activityLabel, opportunity.activityLabel),
            escapeHtml(String(opportunity.occurrenceCount)),
            escapeHtml(formatNumber(opportunity.fastestTime)),
            escapeHtml(formatNumber(opportunity.p25Value)),
            escapeHtml(formatNumber(opportunity.medianValue)),
            escapeHtml(formatNumber(opportunity.meanValue)),
            escapeHtml(formatNumber(opportunity.idealTime) + " (" + opportunity.idealRule + ")"),
            escapeHtml(opportunity.variability),
            confidenceBadgeHtml(opportunity.confidence),
            escapeHtml(formatNumber(opportunity.totalRecoverableHours)),
            flatTimeActionButtonHtml("well", opportunity.topEntry.label || "", (opportunity.topEntry.rigLabel || "Rig not mapped") + " • " + (opportunity.topEntry.label || "N/A")),
          ])
      );

      renderTableHtml(
        ui.flatTimeOpportunityTable,
        ["Section", "Group", "Activity", "Sample", "Highest Rig", "Highest Well", "Actual Time (hr)", "Peer Avg (hr)", "Ideal Time (hr)", "Gap To Ideal (hr)", "How Gap Was Calculated"],
        rankedOpportunities.map((opportunity) => {
          const peerReference = opportunity.peerAverage || opportunity.meanValue || opportunity.medianValue || 0;
          const explanation =
            opportunity.occurrenceCount >= 2
              ? (
                  (opportunity.occurrenceCount - 1) + " peer wells avg " + formatNumber(peerReference) +
                  " hr; " + opportunity.topEntry.label + " ran " + formatNumber(opportunity.topEntry.value) +
                  " hr; ideal = " + formatNumber(opportunity.idealTime) + " hr; gap = " + formatNumber(opportunity.gapToIdeal) + " hr"
                )
              : "Only one well available, so no peer comparison yet";

          return [
            escapeHtml(formatFlatTimeSectionSize(opportunity.sectionSize)),
            escapeHtml(opportunity.groupLabel),
            flatTimeActionButtonHtml("activity", opportunity.activityLabel, opportunity.activityLabel),
            escapeHtml(String(opportunity.occurrenceCount)),
            escapeHtml(opportunity.topEntry.rigLabel || "Rig not mapped"),
            flatTimeActionButtonHtml("well", opportunity.topEntry.label || "", opportunity.topEntry.label || "N/A"),
            escapeHtml(formatNumber(opportunity.topEntry.value)),
            escapeHtml(formatNumber(peerReference)),
            escapeHtml(formatNumber(opportunity.idealTime)),
            escapeHtml(formatNumber(opportunity.gapToIdeal)),
            escapeHtml(explanation + " (" + opportunity.idealRule + ")"),
          ];
        })
      );

      renderTable(
        ui.flatTimeGroupTable,
        ["Group", ...datasets.map((dataset) => dataset.subjectWell), "Total"],
        groupItems.map((item) => [
          item.label,
          ...datasets.map((dataset) => formatNumber(item[dataset.id] || 0)),
          formatNumber(item.total),
        ])
      );

      renderTable(
        ui.flatTimeLossDrivers,
        ["Rig", "Well", "Top Driver 1", "Top Driver 2", "Top Driver 3", "Excess Time (hr)"],
        wellRanking.map((row) => [
          row.rigLabel,
          row.wellLabel,
          row.topDrivers[0] ? row.topDrivers[0].activity + " (+" + formatNumber(row.topDrivers[0].gap) + " hr)" : "-",
          row.topDrivers[1] ? row.topDrivers[1].activity + " (+" + formatNumber(row.topDrivers[1].gap) + " hr)" : "-",
          row.topDrivers[2] ? row.topDrivers[2].activity + " (+" + formatNumber(row.topDrivers[2].gap) + " hr)" : "-",
          formatNumber(row.excessTotal),
        ])
      );

      renderVariabilityChart(ui.flatTimeVariabilityChart, allOpportunities, topN);
      renderHeatmap(ui.flatTimeHeatmap, datasets, allOpportunities, topN);
      renderPerfectFlatTimeChart(ui.flatTimePerfectChart, datasets, metricKey);
      renderFlatTimeDrilldown(datasets, allOpportunities, selectedWell, selectedActivity);
      wireFlatTimeFocusActions();
    }

    function renderWeeklyReport() {
      const range = getWeeklyReportDateRange();
      if (!range.start && !range.end) {
        return;
      }

      const bounds = {
        start: range.start,
        end: range.end,
      };

      const periodRows = dashboardData.interventions.filter((row) => {
        if (!row.date) return false;
        if (bounds.start && row.date < bounds.start) return false;
        if (bounds.end && row.date > bounds.end) return false;
        return true;
      });
      const cumulativeRows = dashboardData.interventions.filter((row) => row.date && bounds.end && row.date <= bounds.end);

      const weeklyCategories = buildCategorySummary(periodRows);
      const cumulativeCategories = buildCategorySummary(cumulativeRows);

      ui.weeklyReportTitle.textContent = "Weekly report for " + (bounds.start && bounds.end ? formatDateHuman(bounds.start) + " - " + formatDateHuman(bounds.end) : "selected range");
      ui.weeklyReportSubtitle.textContent = periodRows.length + " interventions, " + uniqueCount(periodRows, "rigName") + " rigs, " + uniqueCount(periodRows, "wellName") + " wells";
      ui.weeklyReportRange.textContent = bounds.start && bounds.end ? formatDateHuman(bounds.start) + " - " + formatDateHuman(bounds.end) : "Range pending";
      ui.weeklyBannerCopy.textContent = "Selected period: " + ui.weeklyReportRange.textContent + ". The blocks below follow the same report storytelling used in the weekly Excel workbook.";
      ui.weeklyBannerChip1.textContent = periodRows.length + " interventions";
      ui.weeklyBannerChip2.textContent = uniqueCount(periodRows, "rigName") + " active rigs";
      ui.weeklyBannerChip3.textContent = formatCurrency(periodRows.reduce((sum, row) => sum + row.costSavingValue + row.potentialAvoidanceValue, 0)) + " total impact";

      renderTable(
        ui.weeklyCategoryTable,
        ["Category", "Number of Interventions", "Rig Action Taken", "Validation %"],
        weeklyCategories.map((item) => [
          item.label,
          String(item.interventions),
          String(item.rigAction),
          formatPercent(item.validationRate),
        ])
      );

      renderMultiSeriesChart(ui.weeklyCategoryChart, weeklyCategories, [
        { key: "interventions", label: "Interventions", color: "#1264d6" },
        { key: "rigAction", label: "Rig Action", color: "#0f766e" },
        { key: "validationRate", label: "Validation %", color: "#c06a0a", format: (value) => formatPercent(value), scale: (value, context) => value * context.primaryMax, isSecondary: true },
      ]);

      renderTable(
        ui.cumulativeCategoryTable,
        ["Category", "# of Interventions", "Rig Action Taken", "RTOC Validation", "Validation %"],
        cumulativeCategories.map((item) => [
          item.label,
          String(item.interventions),
          String(item.rigAction),
          String(item.validated),
          formatPercent(item.validationRate),
        ])
      );

      renderMultiSeriesChart(ui.cumulativeCategoryChart, cumulativeCategories, [
        { key: "interventions", label: "Interventions", color: "#1264d6" },
        { key: "rigAction", label: "Rig Action", color: "#0f766e" },
        { key: "validated", label: "Validated", color: "#be123c" },
        { key: "validationRate", label: "Validation %", color: "#c06a0a", format: (value) => formatPercent(value), scale: (value, context) => value * context.primaryMax, isSecondary: true },
      ]);

      const wiperSummary = buildRigSummary(periodRows, isWiperTripRow);
      const ropSummary = buildRigSummary(periodRows, isRopRow);
      const kpiSummary = buildRigSummary(periodRows, isKpiRow);
      const summarySeries = [
        { key: "savedTime", label: "Saved Time", color: "#1264d6", format: (value) => formatNumber(value) },
        { key: "lossTime", label: "Loss Time", color: "#c81e5a", format: (value) => formatNumber(value) },
      ];

      renderSummaryBlock(ui.wiperSummaryTable, ui.wiperSummaryChart, wiperSummary, summarySeries);
      renderSummaryBlock(ui.ropSummaryTable, ui.ropSummaryChart, ropSummary, summarySeries);
      renderSummaryBlock(ui.kpiSummaryTable, ui.kpiSummaryChart, kpiSummary, summarySeries);

      const actualWeek = buildHighlightEntries(periodRows, "actual");
      const actualYtd = buildHighlightEntries(cumulativeRows, "actual");
      const potentialWeek = buildHighlightEntries(periodRows, "potential");
      const potentialYtd = buildHighlightEntries(cumulativeRows, "potential");
      const weeklyStatsRows = buildWeeklyStatsRows(periodRows, cumulativeRows);
      const monitoredThisWeekDays = daysBetweenInclusive(bounds.start, bounds.end);
      const monitoredTotalDays = daysBetweenInclusive(dashboardData.meta.monitoringStartDate, bounds.end);

      renderWeeklyMetrics(actualWeek, actualYtd, potentialWeek, potentialYtd);
      renderHighlightTable(ui.actualHighlightsTable, actualWeek, actualYtd, true);
      renderHighlightTable(ui.potentialHighlightsTable, potentialWeek, potentialYtd, false);
      renderWeeklyStatsMetrics(ui.weeklyStatsMetrics, monitoredThisWeekDays, monitoredTotalDays);
      renderWeeklyStatsTable(ui.weeklyStatsTable, weeklyStatsRows);
    }

    function setActiveView(viewId) {
      ui.viewPanels.forEach((panel) => {
        panel.hidden = panel.id !== viewId;
      });
      ui.viewTabs.forEach((button) => {
        button.classList.toggle("is-active", button.dataset.view === viewId);
      });
    }

    function exportWeeklyReportPdf() {
      setActiveView("weekly-report-view");
      renderWeeklyReport();
      window.setTimeout(() => {
        window.print();
      }, 80);
    }

    function getActiveFilters() {
      return {
        startDate: ui.startDate.value,
        endDate: ui.endDate.value,
        week: ui.week.value,
        month: ui.month.value,
        rig: ui.rig.value,
        field: ui.field.value,
        well: ui.well.value,
        category: ui.category.value,
        type: ui.type.value,
        app: ui.app.value,
        rep: ui.rep.value,
        validation: ui.validation.value,
        granularity: ui.granularity.value,
        search: ui.search.value.trim().toLowerCase(),
      };
    }

    function rowMatches(row, filters) {
      if (filters.startDate && row.date && row.date < filters.startDate) return false;
      if (filters.endDate && row.date && row.date > filters.endDate) return false;
      if (filters.week && row.week !== filters.week) return false;
      if (filters.month && row.month !== filters.month) return false;
      if (filters.rig && row.rigName !== filters.rig) return false;
      if (filters.field && row.field !== filters.field) return false;
      if (filters.well && row.wellName !== filters.well) return false;
      if (filters.category && row.category !== filters.category) return false;
      if (filters.type && row.type !== filters.type) return false;
      if (filters.app && row.app !== filters.app) return false;
      if (filters.rep && row.rtesRep !== filters.rep) return false;
      if (filters.validation === "validated" && !row.isValidated) return false;
      if (filters.validation === "not_validated" && row.isValidated) return false;
      if (filters.search && !row.searchText.includes(filters.search)) return false;
      return true;
    }

    function costAvoidanceMatches(row, filters) {
      if (filters.startDate && row.endDate && row.endDate < filters.startDate) return false;
      if (filters.endDate && row.startDate && row.startDate > filters.endDate) return false;
      if (filters.rig && row.rig !== filters.rig) return false;
      if (filters.well && row.well !== filters.well) return false;
      if (filters.search && !row.searchText.includes(filters.search)) return false;
      return true;
    }

    function resetFilters() {
      ui.startDate.value = "";
      ui.endDate.value = "";
      ui.week.value = "";
      ui.month.value = "";
      ui.rig.value = "";
      ui.field.value = "";
      ui.well.value = "";
      ui.category.value = "";
      ui.type.value = "";
      ui.app.value = "";
      ui.rep.value = "";
      ui.validation.value = "";
      ui.granularity.value = "day";
      ui.search.value = "";
      ui.presetButtons.forEach((button) => button.classList.remove("is-active"));
      applyFilters();
    }

    function applyPreset(preset) {
      const defaultRange = getDefaultLastTuesdayRange();
      const referenceEndDate = ui.endDate.value || defaultRange.end;
      if (!referenceEndDate) return;
      const weeklyAllStartDate = dashboardData.meta.monitoringStartDate || dashboardData.meta.minDate || "";
      const end = new Date(referenceEndDate + "T00:00:00");
      let start = null;

      if (preset === "last7") start = new Date(end.getTime() - 6 * 86400000);
      if (preset === "last30") start = new Date(end.getTime() - 29 * 86400000);
      if (preset === "last90") start = new Date(end.getTime() - 89 * 86400000);

      ui.presetButtons.forEach((button) => button.classList.toggle("is-active", button.dataset.preset === preset));
      ui.week.value = "";
      ui.month.value = "";

      if (preset === "all") {
        ui.startDate.value = "";
        ui.endDate.value = "";
        ui.weeklyReportStartDate.value = weeklyAllStartDate;
        ui.weeklyReportEndDate.value = referenceEndDate;
      } else if (start) {
        const startDate = start.toISOString().slice(0, 10);
        ui.startDate.value = startDate;
        ui.endDate.value = referenceEndDate;
        ui.weeklyReportStartDate.value = startDate;
        ui.weeklyReportEndDate.value = referenceEndDate;
      }

      applyFilters();
      renderWeeklyReport();
    }

    function updateSectionVisibility() {
      ui.toggles.forEach((toggle) => {
        const target = document.getElementById(toggle.dataset.target);
        if (target) target.hidden = !toggle.checked;
      });
    }

    function renderKpis(filteredRows, filteredCostAvoidance) {
      const validated = filteredRows.filter((row) => row.isValidated).length;
      const validationRate = filteredRows.length ? (validated / filteredRows.length) * 100 : 0;
      const costSavingHours = filteredRows.reduce((sum, row) => sum + row.costSavingHours, 0);
      const potentialAvoidanceHours = filteredRows.reduce((sum, row) => sum + row.potentialAvoidanceHours, 0);
      const costSavingValue = filteredRows.reduce((sum, row) => sum + row.costSavingValue, 0);
      const potentialAvoidanceValue = filteredRows.reduce((sum, row) => sum + row.potentialAvoidanceValue, 0);
      const caDaysSaved = filteredCostAvoidance.reduce((sum, row) => sum + row.daysSaved, 0);
      const caValue = filteredCostAvoidance.reduce((sum, row) => sum + row.costAvoidanceValue, 0);
      const avgSpreadRate = average(filteredRows.map((row) => row.rigSpreadRate));

      const cards = [
        {
          label: "Interventions",
          value: String(filteredRows.length),
          meta: validated + " validated (" + validationRate.toFixed(1) + "%)",
        },
        {
          label: "Coverage",
          value: String(uniqueCount(filteredRows, "rigName")),
          meta: uniqueCount(filteredRows, "field") + " fields and " + uniqueCount(filteredRows, "wellName") + " wells",
        },
        {
          label: "Hours Impact",
          value: formatNumber(costSavingHours + potentialAvoidanceHours),
          meta: formatNumber(costSavingHours) + " saved + " + formatNumber(potentialAvoidanceHours) + " avoided",
        },
        {
          label: "Financial Impact",
          value: formatCurrency(costSavingValue + potentialAvoidanceValue),
          meta: formatCurrency(costSavingValue) + " saved + " + formatCurrency(potentialAvoidanceValue) + " avoided",
        },
        {
          label: "RTES CA",
          value: formatCurrency(caValue),
          meta: formatNumber(caDaysSaved) + " days saved",
        },
        {
          label: "Avg Spread Rate",
          value: formatCurrency(avgSpreadRate),
          meta: "Average across filtered intervention rows",
        },
      ];

      ui.kpiGrid.innerHTML = cards
        .map((card) => {
          return (
            '<article class="card">' +
            '<div class="card-label">' + escapeHtml(card.label) + "</div>" +
            '<div class="card-value">' + escapeHtml(card.value) + "</div>" +
            '<div class="card-meta">' + escapeHtml(card.meta) + "</div>" +
            "</article>"
          );
        })
        .join("");
    }

    function renderFilterChips(filters) {
      const chips = [];
      if (filters.startDate) chips.push("Start: " + filters.startDate);
      if (filters.endDate) chips.push("End: " + filters.endDate);
      if (filters.week) chips.push("Week: " + filters.week);
      if (filters.month) chips.push("Month: " + filters.month);
      if (filters.rig) chips.push("Rig: " + filters.rig);
      if (filters.field) chips.push("Field: " + filters.field);
      if (filters.well) chips.push("Well: " + filters.well);
      if (filters.category) chips.push("Category: " + filters.category);
      if (filters.type) chips.push("Type: " + filters.type);
      if (filters.app) chips.push("App: " + filters.app);
      if (filters.rep) chips.push("RTES Rep: " + filters.rep);
      if (filters.validation === "validated") chips.push("Validated only");
      if (filters.validation === "not_validated") chips.push("Not validated only");
      if (filters.search) chips.push('Search: "' + filters.search + '"');
      ui.activeFilters.innerHTML = chips.length
        ? chips.map((chip) => '<span class="chip">' + escapeHtml(chip) + "</span>").join("")
        : '<span class="chip">No active filters. Showing the full dataset.</span>';
    }

    function renderRankings(filteredRows) {
      renderTable(
        ui.categoryTable,
        ["Category", "Count"],
        buildCounter(filteredRows, "category").slice(0, 10).map((item) => [item.label, String(item.value)])
      );
      renderTable(
        ui.repTable,
        ["RTES Rep", "Count"],
        buildCounter(filteredRows, "rtesRep").slice(0, 10).map((item) => [item.label, String(item.value)])
      );
      renderTable(
        ui.fieldTable,
        ["Field", "Count"],
        buildCounter(filteredRows, "field").slice(0, 10).map((item) => [item.label, String(item.value)])
      );
      renderTable(
        ui.wellTable,
        ["Well", "Count"],
        buildCounter(filteredRows, "wellName").slice(0, 10).map((item) => [item.label, String(item.value)])
      );
    }

    function renderDetails(filteredRows) {
      const rows = filteredRows
        .slice()
        .sort((left, right) => (right.date || "").localeCompare(left.date || ""))
        .slice(0, 150)
        .map((row) => [
          row.index,
          row.date,
          row.week,
          row.rigName,
          row.field,
          row.wellName,
          row.category,
          row.type,
          row.app,
          row.rtesRep,
          row.isValidated ? "Yes" : "No",
          row.description.slice(0, 90),
        ]);

      renderTable(
        ui.interventionTable,
        ["#", "Date", "Week", "Rig", "Field", "Well", "Category", "Type", "App", "RTES Rep", "Validated", "Description"],
        rows
      );
    }

    function renderCostAvoidance(filteredRows) {
      renderBarChart(
        ui.caChart,
        buildValueCounter(filteredRows, "rig", "costAvoidanceValue"),
        "#c81e5a",
        (value) => formatCurrency(value)
      );

      const rows = filteredRows
        .slice()
        .sort((left, right) => (right.startDate || "").localeCompare(left.startDate || ""))
        .slice(0, 80)
        .map((row) => [
          row.startDate,
          row.endDate,
          row.rig,
          row.well,
          formatNumber(row.daysSaved),
          formatCurrency(row.costAvoidanceValue),
        ]);

      renderTable(ui.caTable, ["Start", "End", "Rig", "Well", "Days Saved", "Cost Avoidance"], rows);
    }

    function applyFilters() {
      const filters = getActiveFilters();
      const filteredRows = dashboardData.interventions.filter((row) => rowMatches(row, filters));
      const filteredCostAvoidance = dashboardData.costAvoidance.filter((row) => costAvoidanceMatches(row, filters));

      ui.resultsTitle.textContent = filteredRows.length + " interventions in view";
      ui.resultsSubtitle.textContent =
        uniqueCount(filteredRows, "rigName") + " rigs, " +
        uniqueCount(filteredRows, "wellName") + " wells, " +
        filteredCostAvoidance.length + " RTES CA rows";

      renderFilterChips(filters);
      renderKpis(filteredRows, filteredCostAvoidance);
      renderTrendChart(ui.trendChart, buildTrend(filteredRows, filters.granularity));
      renderBarChart(ui.categoryChart, buildCounter(filteredRows, "category"), "#1264d6", (value) => String(value));
      renderBarChart(ui.rigChart, buildCounter(filteredRows, "rigName"), "#0f766e", (value) => String(value));
      renderBarChart(ui.typeChart, buildCounter(filteredRows, "type"), "#c06a0a", (value) => String(value));
      renderBarChart(ui.appChart, buildCounter(filteredRows, "app"), "#7c3aed", (value) => String(value));
      renderRankings(filteredRows);
      renderDetails(filteredRows);
      renderCostAvoidance(filteredCostAvoidance);
      updateSectionVisibility();
    }

    function wireEvents() {
      [
        ui.startDate,
        ui.endDate,
        ui.week,
        ui.month,
        ui.rig,
        ui.field,
        ui.well,
        ui.category,
        ui.type,
        ui.app,
        ui.rep,
        ui.validation,
        ui.granularity,
      ].forEach((element) => {
        element.addEventListener("change", () => {
          ui.presetButtons.forEach((button) => button.classList.remove("is-active"));
          applyFilters();
        });
      });

      ui.search.addEventListener("input", applyFilters);
      ui.reset.addEventListener("click", resetFilters);
      ui.themeToggle.addEventListener("click", toggleTheme);
      ui.presetButtons.forEach((button) => {
        button.addEventListener("click", () => applyPreset(button.dataset.preset));
      });
      ui.toggles.forEach((toggle) => {
        toggle.addEventListener("change", updateSectionVisibility);
      });

      ui.viewTabs.forEach((button) => {
        button.addEventListener("click", () => setActiveView(button.dataset.view));
      });

      ui.weeklyExportPdf.addEventListener("click", exportWeeklyReportPdf);
      ui.weeklyReportStartDate.addEventListener("change", () => {
        ui.presetButtons.forEach((button) => button.classList.remove("is-active"));
        renderWeeklyReport();
      });
      ui.weeklyReportEndDate.addEventListener("change", () => {
        ui.presetButtons.forEach((button) => button.classList.remove("is-active"));
        renderWeeklyReport();
      });
      ui.flatTimeRig.addEventListener("change", renderFlatTime);
      ui.flatTimeSection.addEventListener("change", renderFlatTime);
      ui.flatTimeMetric.addEventListener("change", renderFlatTime);
      ui.flatTimeTopN.addEventListener("change", renderFlatTime);
      ui.flatTimeMode.addEventListener("change", renderFlatTime);
      ui.flatTimeWell.addEventListener("change", () => {
        flatTimeState.focusWell = ui.flatTimeWell.value || "";
        renderFlatTime();
      });
      ui.flatTimeRecalculate.addEventListener("click", renderFlatTime);
      ui.flatTimeClearUploads.addEventListener("click", () => {
        flatTimeState.uploadedDatasets = [];
        ui.flatTimeUpload.value = "";
        ui.flatTimeRig.value = "";
        ui.flatTimeSection.value = "";
        ui.flatTimeWell.value = "";
        flatTimeState.focusWell = "";
        flatTimeState.focusActivity = "";
        renderFlatTime();
      });
      ui.flatTimeUpload.addEventListener("change", async () => {
        const files = Array.from(ui.flatTimeUpload.files || []);
        if (!files.length) return;
        const parsed = await Promise.all(
          files.map(async (file) => parseFlatTimeCsvText(file.name, await file.text()))
        );
        const prepared = parsed
          .flat()
          .filter((item) => item && item.groups && item.groups.length)
          .map((item) => ({ ...item, id: createFlatTimeUploadId(item.fileName, item.subjectWell) }));
        flatTimeState.uploadedDatasets = [
          ...flatTimeState.uploadedDatasets,
          ...prepared,
        ];
        renderFlatTime();
      });
    }

    function initialize() {
      applyTheme(localStorage.getItem(THEME_STORAGE_KEY) || "classic");
      populateSelect(ui.week, dashboardData.filters.weeks, "All weeks");
      populateSelect(ui.month, dashboardData.filters.months, "All months");
      populateSelect(ui.rig, dashboardData.filters.rigs, "All rigs");
      populateSelect(ui.field, dashboardData.filters.fields, "All fields");
      populateSelect(ui.well, dashboardData.filters.wells, "All wells");
      populateSelect(ui.category, dashboardData.filters.categories, "All categories");
      populateSelect(ui.type, dashboardData.filters.types, "All types");
      populateSelect(ui.app, dashboardData.filters.apps, "All apps");
      populateSelect(ui.rep, dashboardData.filters.reps, "All RTES reps");
      if (dashboardData.meta.minDate) ui.startDate.min = dashboardData.meta.minDate;
      if (dashboardData.meta.maxDate) ui.startDate.max = dashboardData.meta.maxDate;
      if (dashboardData.meta.minDate) ui.endDate.min = dashboardData.meta.minDate;
      if (dashboardData.meta.maxDate) ui.endDate.max = dashboardData.meta.maxDate;
      if (dashboardData.meta.minDate) ui.weeklyReportStartDate.min = dashboardData.meta.minDate;
      if (dashboardData.meta.maxDate) ui.weeklyReportStartDate.max = dashboardData.meta.maxDate;
      if (dashboardData.meta.minDate) ui.weeklyReportEndDate.min = dashboardData.meta.minDate;
      if (dashboardData.meta.maxDate) ui.weeklyReportEndDate.max = dashboardData.meta.maxDate;

      const defaultRange = getDefaultLastTuesdayRange();
      ui.startDate.value = defaultRange.start;
      ui.endDate.value = defaultRange.end;
      ui.weeklyReportStartDate.value = defaultRange.start;
      ui.weeklyReportEndDate.value = defaultRange.end;

      wireEvents();
      setActiveView("dashboard-view");
      updateSectionVisibility();
      applyFilters();
      renderWeeklyReport();
      renderFlatTime();
    }

    initialize();
  </script>
</body>
</html>
"""


@dataclass
class SheetData:
    name: str
    headers: list[str]
    rows: list[dict[str, str]]


def excel_serial_to_datetime(value: str) -> datetime | None:
    try:
        serial = float(value)
    except (TypeError, ValueError):
        return None
    return datetime(1899, 12, 30) + timedelta(days=serial)


def parse_datetime_value(value: str | None) -> datetime | None:
    text = str(value or "").strip()
    if not text:
        return None

    serial_date = excel_serial_to_datetime(text)
    if serial_date:
        return serial_date

    normalized = (
        text.replace(".", "/")
        .replace("-", "/")
        .replace("\\", "/")
        .replace("T", " ")
    )
    normalized = re.sub(r"\s+", " ", normalized).strip()
    normalized = normalized.split(" ")[0]

    for fmt in (
        "%m/%d/%Y",
        "%m/%d/%y",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%Y/%m/%d",
    ):
        try:
            return datetime.strptime(normalized, fmt)
        except ValueError:
            continue
    return None


def column_letters_to_index(cell_ref: str) -> int:
    letters = "".join(char for char in cell_ref if char.isalpha())
    result = 0
    for char in letters:
        result = (result * 26) + (ord(char.upper()) - 64)
    return result - 1


def load_shared_strings(workbook: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in workbook.namelist():
        return []
    root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
    items = []
    for item in root.findall("a:si", NS):
        text = "".join(node.text or "" for node in item.findall(".//a:t", NS))
        items.append(text)
    return items


def read_cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(".//a:t", NS)).strip()

    value_node = cell.find("a:v", NS)
    if value_node is None or value_node.text is None:
        return ""

    raw_value = value_node.text.strip()
    if cell_type == "s":
        try:
            return shared_strings[int(raw_value)]
        except (ValueError, IndexError):
            return raw_value
    if cell_type == "b":
        return "TRUE" if raw_value == "1" else "FALSE"
    return raw_value


def load_sheet_rows(workbook: ZipFile, sheet_path: str, shared_strings: list[str]) -> list[list[str]]:
    root = ET.fromstring(workbook.read(sheet_path))
    rows: list[list[str]] = []
    for row in root.findall(".//a:sheetData/a:row", NS):
        values: list[str] = []
        for cell in row.findall("a:c", NS):
            ref = cell.attrib.get("r", "")
            index = column_letters_to_index(ref) if ref else len(values)
            while len(values) <= index:
                values.append("")
            values[index] = read_cell_value(cell, shared_strings)
        rows.append(values)
    return rows


def normalize_header(value: str, index: int) -> str:
    text = " ".join((value or "").replace("\xa0", " ").split()).strip()
    return text or f"Column {index + 1}"


def infer_header_row(rows: list[list[str]]) -> int:
    best_index = 0
    best_score = -1
    for index, row in enumerate(rows[:10]):
        normalized = [normalize_header(cell, idx) for idx, cell in enumerate(row)]
        non_empty = sum(1 for cell in row if str(cell).strip())
        unique = len(set(cell for cell in normalized if cell))
        score = (non_empty * 10) + unique
        if score > best_score:
            best_index = index
            best_score = score
    return best_index


def rows_to_dicts(name: str, rows: list[list[str]]) -> SheetData:
    if not rows:
        return SheetData(name=name, headers=[], rows=[])

    header_index = infer_header_row(rows)
    headers = [normalize_header(cell, idx) for idx, cell in enumerate(rows[header_index])]
    width = len(headers)
    records: list[dict[str, str]] = []

    for raw_row in rows[header_index + 1 :]:
        padded = list(raw_row[:width]) + [""] * max(0, width - len(raw_row))
        record = {headers[idx]: (padded[idx] or "").strip() for idx in range(width)}
        if any(value for value in record.values()):
            records.append(record)

    return SheetData(name=name, headers=headers, rows=records)


def load_workbook(source: Path | bytes | BytesIO) -> list[SheetData]:
    workbook_source = BytesIO(source) if isinstance(source, bytes) else source
    with ZipFile(workbook_source) as workbook:
        shared_strings = load_shared_strings(workbook)
        workbook_root = ET.fromstring(workbook.read("xl/workbook.xml"))
        rels_root = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
        rel_map = {
            rel.attrib["Id"]: f"xl/{rel.attrib['Target']}"
            for rel in rels_root.findall("pr:Relationship", NS)
        }

        sheets = []
        sheets_node = workbook_root.find("a:sheets", NS)
        for sheet in list(sheets_node) if sheets_node is not None else []:
            name = sheet.attrib["name"]
            rel_id = sheet.attrib[f"{{{DOCUMENT_REL_NS}}}id"]
            rows = load_sheet_rows(workbook, rel_map[rel_id], shared_strings)
            sheets.append(rows_to_dicts(name, rows))
        return sheets


def as_number(value: str | None) -> float:
    if value is None:
        return 0.0
    text = str(value).replace(",", "").replace("$", "").replace("\xa0", "").strip()
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def iso_date(value: str | None) -> str:
    date_value = parse_datetime_value(value)
    if not date_value:
        return ""
    return date_value.strftime("%Y-%m-%d")


def iso_week(value: str | None) -> str:
    date_value = parse_datetime_value(value)
    if not date_value:
        return ""
    iso_year, iso_week_number, _ = date_value.isocalendar()
    return f"{iso_year}-W{iso_week_number:02d}"


def iso_month(value: str | None) -> str:
    date_value = parse_datetime_value(value)
    if not date_value:
        return ""
    return date_value.strftime("%Y-%m")


def normalize_text(value: str | None) -> str:
    return " ".join(str(value or "").replace("\xa0", " ").split()).strip()


def build_intervention_rows(rows: list[dict[str, str]]) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    for row in rows:
        date = iso_date(row.get("Date"))
        record = {
            "index": normalize_text(row.get("Interventions Count")),
            "date": date,
            "week": iso_week(row.get("Date")),
            "month": iso_month(row.get("Date")),
            "rigName": normalize_text(row.get("Rig Name")),
            "field": normalize_text(row.get("Field")),
            "wellName": normalize_text(row.get("Well Name")),
            "holeSize": normalize_text(row.get("Hole Size")),
            "engDept": normalize_text(row.get("ENG. Dept")),
            "optDept": normalize_text(row.get("Opt. Dept")),
            "category": normalize_text(row.get("Intervention Category")),
            "type": normalize_text(row.get("Intervention Type")),
            "eventIndex": normalize_text(row.get("Event Index")),
            "app": normalize_text(row.get("Corva App")),
            "parameter": normalize_text(row.get("Parameter")),
            "expected": normalize_text(row.get("Expected")),
            "actual": normalize_text(row.get("Actual")),
            "description": normalize_text(row.get("Intervention Description")),
            "recommendation": normalize_text(row.get("Recommendation")),
            "validationText": normalize_text(row.get("RTOC/RDH Validation (Y/N)")),
            "isValidated": normalize_text(row.get("RTOC/RDH Validation (Y/N)")).lower() in {"yes", "y", "true"},
            "rtocComments": normalize_text(row.get("RTOC Comments")),
            "rtocCommunication": normalize_text(row.get("RTOC to Rig Communication")),
            "rigAction": normalize_text(row.get("Rig Taken Action")),
            "rigComment": normalize_text(row.get("Rig Comment")),
            "rtocLeadName": normalize_text(row.get("RTOC lead name")),
            "rtesRep": normalize_text(row.get("RTES Rep")),
            "costSavingHours": as_number(row.get("Cost Saving (hrs)")),
            "potentialAvoidanceHours": as_number(row.get("Potential Cost Avoidance (hrs)")),
            "rigSpreadRate": as_number(row.get("Rig Spread Rate ($/day)")),
            "costSavingValue": as_number(row.get("$ Cost Saving")),
            "potentialAvoidanceValue": as_number(row.get("$ Potential Cost Avoidance")),
        }
        search_fields = [
            record["date"],
            record["week"],
            record["month"],
            record["rigName"],
            record["field"],
            record["wellName"],
            record["category"],
            record["type"],
            record["app"],
            record["parameter"],
            record["description"],
            record["recommendation"],
            record["rtocComments"],
            record["rigComment"],
            record["rtesRep"],
        ]
        record["searchText"] = " ".join(value.lower() for value in search_fields if value)
        if any(
            [
                record["date"],
                record["rigName"],
                record["field"],
                record["wellName"],
                record["category"],
                record["type"],
                record["app"],
                record["description"],
                record["rtesRep"],
            ]
        ):
            output.append(record)
    return output


def build_cost_avoidance_rows(rows: list[dict[str, str]]) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    for row in rows:
        record = {
            "startDate": iso_date(row.get("MONITOR_STA_DTTM")),
            "endDate": iso_date(row.get("MONITOR_END_DTTM")),
            "rig": normalize_text(row.get("Rig")),
            "well": normalize_text(row.get("Well")),
            "daysSaved": as_number(row.get("Days Saved")),
            "costAvoidanceValue": as_number(row.get("Cost Avoidance US$")),
            "costIncluded": normalize_text(row.get("cost included")),
            "caDisplay": normalize_text(row.get("RTOC CA")),
        }
        search_fields = [
            record["startDate"],
            record["endDate"],
            record["rig"],
            record["well"],
            record["costIncluded"],
            record["caDisplay"],
        ]
        record["searchText"] = " ".join(value.lower() for value in search_fields if value)
        if any([record["startDate"], record["endDate"], record["rig"], record["well"]]):
            output.append(record)
    return output


def sorted_unique(values: list[str]) -> list[str]:
    return sorted({value for value in values if value})


def extract_flat_time_section_size(activity_name: str) -> str:
    match = re.match(r"^(\d+(?:\.\d+)?)(?=-)", activity_name or "")
    return match.group(1) if match else "__no_section__"


def parse_flat_time_rows(file_name: str, rows: list[list[str]]) -> dict[str, Any] | None:
    normalized_rows = [[(cell or "").strip() for cell in row] for row in rows]
    subject_well = next((row[1].strip() for row in normalized_rows if len(row) > 1 and row[0] == "Subject Well"), file_name)
    groups: list[dict[str, Any]] = []
    current_group: dict[str, Any] | None = None

    for row in normalized_rows:
        first = row[0] if row else ""
        if first == "Group Name":
            if current_group and current_group["activities"]:
                groups.append(current_group)
            current_group = {
                "groupName": row[1] if len(row) > 1 else "Unknown",
                "activities": [],
                "totalSubjectHours": 0.0,
                "totalMeanHours": 0.0,
                "totalMedianHours": 0.0,
            }
            continue

        if current_group is None or not first or first in {"Group Type", "Activity"}:
            continue

        if first == "Total":
            current_group["totalSubjectHours"] = as_number(row[1] if len(row) > 1 else 0)
            continue

        current_group["activities"].append(
            {
                "activity": first,
                "sectionSize": extract_flat_time_section_size(first),
                "subjectHours": as_number(row[1] if len(row) > 1 else 0),
                "meanHours": as_number(row[2] if len(row) > 2 else 0),
                "medianHours": as_number(row[3] if len(row) > 3 else 0),
            }
        )

    if current_group and current_group["activities"]:
        groups.append(current_group)

    if not groups:
        return None

    for group in groups:
        if not group["totalSubjectHours"]:
            group["totalSubjectHours"] = sum(activity["subjectHours"] for activity in group["activities"])
        group["totalMeanHours"] = sum(activity["meanHours"] for activity in group["activities"])
        group["totalMedianHours"] = sum(activity["medianHours"] for activity in group["activities"])

    return {
        "id": f"{Path(file_name).stem.lower().replace(' ', '-').replace('_', '-')}-{subject_well.lower().replace(' ', '-')}",
        "fileName": file_name,
        "subjectWell": subject_well,
        "groups": groups,
        "totalSubjectHours": sum(group["totalSubjectHours"] for group in groups),
        "totalMeanHours": sum(group["totalMeanHours"] for group in groups),
        "totalMedianHours": sum(group["totalMedianHours"] for group in groups),
    }


def load_flat_time_directory(directory: Path) -> dict[str, Any]:
    datasets: list[dict[str, Any]] = []
    for path in sorted(directory.glob("*.csv")):
        if path.name.startswith("."):
            continue
        try:
            with path.open(newline="", encoding="utf-8-sig") as handle:
                rows = list(csv.reader(handle))
        except OSError:
            continue
        parsed = parse_flat_time_rows(path.name, rows)
        if parsed:
            datasets.append(parsed)
    return {"datasets": datasets}


def load_activity_code_translations() -> dict[str, Any]:
    candidates = [
        Path.cwd() / "Aramco Activity Codes.csv",
        Path(__file__).with_name("Aramco Activity Codes.csv"),
        Path.home() / "Downloads" / "Aramco Activity Codes.csv",
    ]

    source_path = next((path for path in candidates if path.exists()), None)
    if source_path is None:
        return {
            "loaded": False,
            "source": "",
            "wellSections": {},
            "operations": {},
            "activities": {},
            "generic": {},
        }

    well_sections: dict[str, str] = {}
    operations: dict[str, str] = {}
    activities: dict[str, str] = {}
    generic: dict[str, str] = {}

    try:
        with source_path.open(newline="", encoding="utf-8-sig") as handle:
            for row in csv.DictReader(handle):
                code = (row.get("Client Code") or "").strip()
                description = (row.get("Client Code Description") or "").strip()
                code_type = (row.get("Corva Code Type") or "").strip().lower()
                if not code or not description:
                    continue
                generic.setdefault(code, description)
                if code_type == "well sections":
                    well_sections.setdefault(code, description)
                elif code_type == "operation":
                    operations.setdefault(code, description)
                elif code_type == "activity":
                    activities.setdefault(code, description)
    except OSError:
        return {
            "loaded": False,
            "source": "",
            "wellSections": {},
            "operations": {},
            "activities": {},
            "generic": {},
        }

    return {
        "loaded": True,
        "source": source_path.name,
        "wellSections": well_sections,
        "operations": operations,
        "activities": activities,
        "generic": generic,
    }


def build_dashboard_payload(
    spreadsheet_name: str,
    generated_at: str,
    intervention_rows: list[dict[str, Any]],
    cost_avoidance_rows: list[dict[str, Any]],
    flat_time_payload: dict[str, Any] | None = None,
    activity_code_translations: dict[str, Any] | None = None,
) -> dict[str, Any]:
    dates = [row["date"] for row in intervention_rows if row["date"]]
    payload = {
        "meta": {
            "sourceFile": spreadsheet_name,
            "generatedAt": generated_at,
            "minDate": min(dates) if dates else "",
            "maxDate": max(dates) if dates else "",
            "monitoringStartDate": "2025-11-11",
        },
        "filters": {
            "weeks": sorted_unique([row["week"] for row in intervention_rows]),
            "months": sorted_unique([row["month"] for row in intervention_rows]),
            "rigs": sorted_unique([row["rigName"] for row in intervention_rows]),
            "fields": sorted_unique([row["field"] for row in intervention_rows]),
            "wells": sorted_unique([row["wellName"] for row in intervention_rows]),
            "categories": sorted_unique([row["category"] for row in intervention_rows]),
            "types": sorted_unique([row["type"] for row in intervention_rows]),
            "apps": sorted_unique([row["app"] for row in intervention_rows]),
            "reps": sorted_unique([row["rtesRep"] for row in intervention_rows]),
        },
        "interventions": intervention_rows,
        "costAvoidance": cost_avoidance_rows,
    }
    if flat_time_payload is not None:
        payload["flatTime"] = flat_time_payload
    if activity_code_translations is not None:
        payload["activityCodeTranslations"] = activity_code_translations
    return payload


def build_html(title: str, source_file: str, generated_at: str, payload: dict[str, Any]) -> str:
    output = HTML_TEMPLATE
    output = output.replace("__TITLE__", html.escape(title))
    output = output.replace("__SOURCE_FILE__", html.escape(source_file))
    output = output.replace("__GENERATED_AT__", html.escape(generated_at))
    output = output.replace("__DATA_JSON__", json.dumps(payload, separators=(",", ":"), ensure_ascii=True))
    return output


def build_report_html(spreadsheet_name: str, spreadsheet_bytes: bytes, flat_time_payload: dict[str, Any] | None = None) -> str:
    sheets = load_workbook(spreadsheet_bytes)
    sheet_map = {sheet.name: sheet for sheet in sheets}

    interventions_source = sheet_map.get("Intervention Log", SheetData("Intervention Log", [], [])).rows
    cost_avoidance_source = sheet_map.get("RTES CA", SheetData("RTES CA", [], [])).rows

    intervention_rows = build_intervention_rows(interventions_source)
    cost_avoidance_rows = build_cost_avoidance_rows(cost_avoidance_source)

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    payload = build_dashboard_payload(
        spreadsheet_name=spreadsheet_name,
        generated_at=generated_at,
        intervention_rows=intervention_rows,
        cost_avoidance_rows=cost_avoidance_rows,
        flat_time_payload=flat_time_payload,
        activity_code_translations=load_activity_code_translations(),
    )

    return build_html(
        title=f"Interactive Report - {spreadsheet_name}",
        source_file=spreadsheet_name,
        generated_at=generated_at,
        payload=payload,
    )


def build_csv_sheet(csv_name: str, csv_bytes: bytes) -> SheetData:
    rows = list(csv.reader(csv_bytes.decode("utf-8-sig", errors="ignore").splitlines()))
    return rows_to_dicts(csv_name, rows)


def build_intervention_csv_report_html(csv_name: str, csv_bytes: bytes, flat_time_payload: dict[str, Any] | None = None) -> str:
    csv_sheet = build_csv_sheet(csv_name, csv_bytes)
    intervention_rows = build_intervention_rows(csv_sheet.rows)
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    payload = build_dashboard_payload(
        spreadsheet_name=f"{csv_name} (Intervention Log CSV only)",
        generated_at=generated_at,
        intervention_rows=intervention_rows,
        cost_avoidance_rows=[],
        flat_time_payload=flat_time_payload,
        activity_code_translations=load_activity_code_translations(),
    )

    return build_html(
        title=f"Interactive Report - {csv_name}",
        source_file=f"{csv_name} (CSV mode)",
        generated_at=generated_at,
        payload=payload,
    )


def build_empty_report_html(flat_time_payload: dict[str, Any] | None = None) -> str:
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    payload = build_dashboard_payload(
        spreadsheet_name="No workbook loaded",
        generated_at=generated_at,
        intervention_rows=[],
        cost_avoidance_rows=[],
        flat_time_payload=flat_time_payload if flat_time_payload is not None else {"datasets": []},
        activity_code_translations=load_activity_code_translations(),
    )

    return build_html(
        title="Interactive Report - No workbook loaded",
        source_file="No workbook loaded",
        generated_at=generated_at,
        payload=payload,
    )


def build_report(spreadsheet: Path, output_path: Path) -> None:
    html_output = build_report_html(
        spreadsheet.name,
        spreadsheet.read_bytes(),
        flat_time_payload={"datasets": []},
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html_output, encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate an interactive HTML report from an Excel spreadsheet.")
    parser.add_argument(
        "spreadsheet",
        nargs="?",
        help="Path to the .xlsx file. Defaults to the first .xlsx found in the current directory.",
    )
    parser.add_argument(
        "--output",
        help="Path to the generated HTML report. Defaults to report_output/report.html",
    )
    return parser.parse_args()


def find_default_spreadsheet(cwd: Path) -> Path:
    candidates = sorted(cwd.glob("*.xlsx"))
    if not candidates:
        raise FileNotFoundError("No .xlsx files found in the current directory.")
    return candidates[0]


def main() -> int:
    args = parse_args()
    cwd = Path.cwd()
    spreadsheet = Path(args.spreadsheet).expanduser() if args.spreadsheet else find_default_spreadsheet(cwd)
    if not spreadsheet.is_absolute():
        spreadsheet = cwd / spreadsheet
    if not spreadsheet.exists():
        raise FileNotFoundError(f"Spreadsheet not found: {spreadsheet}")

    output = Path(args.output).expanduser() if args.output else cwd / "report_output" / "report.html"
    if not output.is_absolute():
        output = cwd / output

    build_report(spreadsheet, output)
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
