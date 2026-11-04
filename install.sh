#!/usr/bin/env bash
# ==============================================================================
# PQP Drawing, Spec & PCB Quality Inspection Suite - Linux / macOS Installer
# Supported Environments: Antigravity / Claude Code / Codex / Cursor
# Skills: pqp-drawing-msa-auditor, drawing-spec-pqp-comparison, pcb-pqp-spec-comparison
# ==============================================================================

set -euo pipefail

TARGET_DIR=""
MODE="default"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --target)
      TARGET_DIR="$2"
      shift 2
      ;;
    --workspace)
      MODE="workspace"
      shift
      ;;
    --claude)
      MODE="claude"
      shift
      ;;
    *)
      echo "Unknown option: $1"
      exit 1
      ;;
  esac
done

echo "=========================================================="
echo "  PQP Drawing, Spec & PCB Inspection Suite Installer      "
echo "=========================================================="

if [[ -n "$TARGET_DIR" ]]; then
  DEST_BASE="$TARGET_DIR"
elif [[ "$MODE" == "workspace" ]]; then
  DEST_BASE="$(pwd)/.agents/skills"
elif [[ "$MODE" == "claude" ]]; then
  DEST_BASE="$HOME/.claude/skills"
else
  # Default Antigravity global skills directory
  DEST_BASE="$HOME/.gemini/config/skills"
fi

echo "[1/3] Destination: $DEST_BASE"
mkdir -p "$DEST_BASE"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILLS=("pqp-drawing-msa-auditor" "drawing-spec-pqp-comparison" "pcb-pqp-spec-comparison")

echo "[2/3] Installing skills..."
for skill in "${SKILLS[@]}"; do
  SRC_PATH="$SCRIPT_DIR/$skill"
  TARGET_PATH="$DEST_BASE/$skill"

  if [[ ! -d "$SRC_PATH" ]]; then
    echo "Error: Source directory not found: $SRC_PATH"
    continue
  fi

  if [[ -d "$TARGET_PATH" ]]; then
    echo "  -> Updating existing: $skill"
    rm -rf "$TARGET_PATH"
  fi

  cp -r "$SRC_PATH" "$TARGET_PATH"
  echo "  + Installed: $skill"
done

echo "[3/3] Verifying installation..."
ALL_OK=true
for skill in "${SKILLS[@]}"; do
  SKILL_MD="$DEST_BASE/$skill/SKILL.md"
  if [[ -f "$SKILL_MD" ]]; then
    echo "  + Verified: $SKILL_MD"
  else
    echo "  x Verification failed: $SKILL_MD missing"
    ALL_OK=false
  fi
done

echo "=========================================================="
if [[ "$ALL_OK" == true ]]; then
  echo "SUCCESS: PQP Drawing, Spec & PCB Skills installed successfully!"
  echo ""
  echo "How to use in Antigravity / Claude Code / Codex:"
  echo "  1. Slash commands:"
  echo "     /pqp-drawing-msa-auditor"
  echo "     /drawing-spec-pqp-comparison"
  echo "     /pcb-pqp-spec-comparison"
  echo "  2. CLI direct execution via uv:"
  echo "     uv run ~/.gemini/config/skills/pqp-drawing-msa-auditor/scripts/pqp_dwg_msa_auditor.py all --pqp PQP.xlsx --dwgs Drawing.pdf"
  echo "     uv run ~/.gemini/config/skills/drawing-spec-pqp-comparison/scripts/compare_dwg_pqp.py compare --pqp PQP.xlsx --dwgs Drawing.pdf --output report.md"
  echo "     uv run ~/.gemini/config/skills/pcb-pqp-spec-comparison/scripts/compare_pcb_pqp.py compare --spec Spec.pdf --pqp PQP.xlsx --qa QA.xls --fab FAB.pdf --output pcb_report.md"
else
  echo "WARNING: Some skill files could not be verified."
fi
echo "=========================================================="