#!/bin/bash
# Adisyoapp - SessionStart hook
# Auto-detects the project's tech stack and installs dependencies so that
# tests and linters are ready when a Claude Code (web) session starts.
#
# The repo is currently a fresh scaffold. This hook is stack-agnostic and
# no-ops gracefully when a given toolchain or manifest is not present, so it
# will "just work" once you add Node, .NET, Flutter, or Python code.
set -euo pipefail

ROOT="${CLAUDE_PROJECT_DIR:-$(git rev-parse --show-toplevel 2>/dev/null || pwd)}"
cd "$ROOT"

log() { printf '[session-start] %s\n' "$1"; }

have() { command -v "$1" >/dev/null 2>&1; }

# --- Node.js (npm / yarn / pnpm) -------------------------------------------
if [ -f package.json ]; then
  if [ -f pnpm-lock.yaml ] && have pnpm; then
    log "Detected pnpm project -> pnpm install"
    pnpm install
  elif [ -f yarn.lock ] && have yarn; then
    log "Detected yarn project -> yarn install"
    yarn install
  elif have npm; then
    log "Detected npm project -> npm install"
    npm install
  else
    log "package.json found but no Node package manager available; skipping"
  fi
fi

# --- .NET (C#) --------------------------------------------------------------
if ls ./*.sln >/dev/null 2>&1 || find . -maxdepth 3 -name '*.csproj' -print -quit | grep -q .; then
  if have dotnet; then
    log "Detected .NET project -> dotnet restore"
    dotnet restore
  else
    log ".NET project found but dotnet SDK not available; skipping"
  fi
fi

# --- Flutter / Dart ---------------------------------------------------------
if [ -f pubspec.yaml ]; then
  if have flutter; then
    log "Detected Flutter project -> flutter pub get"
    flutter pub get
  elif have dart; then
    log "Detected Dart project -> dart pub get"
    dart pub get
  else
    log "pubspec.yaml found but Flutter/Dart SDK not available; skipping"
  fi
fi

# --- Python -----------------------------------------------------------------
if [ -f requirements.txt ] || [ -f pyproject.toml ]; then
  if have python3; then
    log "Detected Python project -> installing dependencies"
    if [ -f requirements.txt ]; then
      python3 -m pip install --quiet -r requirements.txt || log "pip install (requirements.txt) failed; continuing"
    fi
    if [ -f pyproject.toml ] && have poetry; then
      poetry install || log "poetry install failed; continuing"
    fi
    echo 'export PYTHONPATH="."' >> "${CLAUDE_ENV_FILE:-/dev/null}" 2>/dev/null || true
  else
    log "Python project found but python3 not available; skipping"
  fi
fi

log "Dependency setup complete."
