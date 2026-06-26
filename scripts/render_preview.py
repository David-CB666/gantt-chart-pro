#!/usr/bin/env python3
"""Final polished hero Gantt chart for GitHub README."""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from datetime import datetime, timedelta
from pathlib import Path

TASKS = [
    ('A.', 'Planning & Preparation', '2026-01-01', 5, True, False),
    ('A1', 'Site Survey & Measurement', '2026-01-01', 2, False, False),
    ('A2', 'Temporary Fencing Setup', '2026-01-01', 1, False, False),
    ('A3', 'Drawing Submission & Approval', '2026-01-02', 3, False, False),
    ('A4', 'Material Procurement', '2026-01-02', 5, False, False),
    ('M1', 'Material On-Site', '2026-01-06', 1, False, True),
    ('B.', 'Main Structure Works', '2026-01-06', 20, True, False),
    ('B1', 'Structural Construction', '2026-01-06', 8, False, False),
    ('B2', 'Electrical Installation & Testing', '2026-01-14', 8, False, False),
    ('B3', 'Piping System Installation', '2026-01-14', 7, False, False),
    ('C.', 'Completion & Handover', '2026-01-26', 12, True, False),
    ('C1', 'Site Cleanup & Documentation', '2026-01-26', 3, False, False),
    ('M2', 'Client Inspection', '2026-02-01', 1, False, True),
    ('C2', 'Final Handover & Closeout', '2026-02-02', 5, False, False),
]

CAT_COLOR = '#1F3864'
BAR_COLOR = '#2E75B6'
BAR_LIGHT = '#E8ECF0'
MILESTONE_COLOR = '#C00000'
GRID_COLOR = '#D6DCE4'
TODAY_COLOR = '#FF4444'

def render(out_path):
    n = len(TASKS)
    fig, ax = plt.subplots(figsize=(12, 6), dpi=150)
    
    overall_start = datetime(2026, 1, 1)
    overall_end = datetime(2026, 2, 6)
    days_total = (overall_end - overall_start).days + 1
    bar_h = 0.56
    today = datetime(2026, 1, 22)
    
    left_margin = 6  # days of left padding
    right_margin = 20  # days of right padding for labels
    
    for i, task in enumerate(TASKS):
        seq, name, start_str, dur, is_cat, is_milestone = task
        y = i
        start = datetime.strptime(start_str, '%Y-%m-%d')
        offset = (start - overall_start).days + left_margin
        
        if is_milestone:
            ax.plot(offset, y, 'D', color=MILESTONE_COLOR, markersize=10,
                   markeredgecolor='white', markeredgewidth=1.5, zorder=6)
            ax.text(left_margin - 0.5, y, seq, ha='right', va='center', fontsize=7.5,
                   fontweight='bold', color=MILESTONE_COLOR)
            ax.text(offset + 1, y, f'{name}\u2713', ha='left', va='center', fontsize=7.5,
                   fontweight='bold', color=MILESTONE_COLOR)
        elif is_cat:
            days_past = max(0, (today - start).days + 1)
            progress = min(1.0, days_past / max(1, dur))
            # Remaining
            ax.barh(y, dur, bar_h, left=offset, color=BAR_LIGHT, alpha=0.9,
                   edgecolor='white', lw=0.5, zorder=2)
            # Done
            if progress > 0:
                done_dur = min(dur, days_past)
                ax.barh(y, done_dur, bar_h, left=offset, color=CAT_COLOR, alpha=0.94,
                       edgecolor='none', zorder=3)
            ax.text(left_margin - 0.5, y, seq, ha='right', va='center',
                   fontsize=8.5, fontweight='bold', color=CAT_COLOR)
            pct = int(progress * 100)
            ax.text(offset + dur + 0.5, y, f'{name} ({dur}d, {pct}%)',
                   ha='left', va='center', fontsize=8, fontweight='bold', color='#1a1a1a')
        else:
            days_past = max(0, (today - start).days + 1)
            progress = min(1.0, days_past / max(1, dur))
            # Remaining
            ax.barh(y, dur, bar_h, left=offset, color=BAR_LIGHT, alpha=0.8,
                   edgecolor='white', lw=0.5, zorder=2)
            # Done
            if progress > 0:
                done_dur = min(dur, days_past)
                ax.barh(y, done_dur, bar_h, left=offset, color=BAR_COLOR, alpha=0.9,
                       edgecolor='none', zorder=3)
            ax.text(left_margin - 0.5, y, seq, ha='right', va='center', fontsize=7.3, color='#666')
            pct = int(progress * 100)
            ax.text(offset + dur + 0.5, y, f'{name} ({pct}%)',
                   ha='left', va='center', fontsize=7.3, color='#555')
    
    # TODAY LINE
    today_offset = (today - overall_start).days + left_margin + 0.5
    ax.axvline(x=today_offset, color=TODAY_COLOR, linewidth=2.2, linestyle='-', alpha=0.85, zorder=10)
    ax.text(today_offset + 0.3, n - 0.3, 'TODAY', fontsize=8, fontweight='bold',
           color=TODAY_COLOR, va='bottom', ha='left')
    
    ax.set_yticks(range(n))
    ax.set_yticklabels([])
    ax.invert_yaxis()
    
    # Weekly date ticks
    dates = [overall_start + timedelta(days=d) for d in range(0, days_total + 1, 7)]
    ax.set_xticks([(d - overall_start).days + left_margin + 0.5 for d in dates])
    ax.set_xticklabels([d.strftime('%b %d') for d in dates], fontsize=8, color='#666')
    ax.set_xlim(0, days_total + left_margin + right_margin)
    
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.tick_params(left=False, bottom=False)
    ax.grid(axis='x', color=GRID_COLOR, linewidth=0.4, alpha=0.35)
    ax.set_axisbelow(True)
    
    # Title
    title_x = left_margin / (days_total + left_margin + right_margin)
    fig.suptitle('GanttChart Pro', fontsize=18, fontweight='bold', color=CAT_COLOR,
                x=title_x + 0.02, y=0.99, ha='left')
    fig.text(title_x + 0.03, 0.94, 'No MS Project required — pure Python + openpyxl',
            fontsize=9, color='#999', style='italic')
    
    # Legend
    legend = [
        mpatches.Patch(color=CAT_COLOR, label='Phase'),
        mpatches.Patch(color=BAR_COLOR, label='Task'),
        mpatches.Patch(color=BAR_LIGHT, label='Remaining'),
        plt.Line2D([0],[0], marker='D', color='w', markerfacecolor=MILESTONE_COLOR,
                  markersize=8, markeredgewidth=1.5, label='Milestone'),
        plt.Line2D([0],[0], color=TODAY_COLOR, linewidth=2.2, label='Today'),
    ]
    fig.legend(handles=legend, loc='lower right', fontsize=7.5, framealpha=0.92,
              edgecolor='#ccc', ncol=5, columnspacing=1)
    
    plt.tight_layout(pad=1.2, rect=[0, 0, 1, 0.93])
    fig.savefig(out_path, dpi=150, bbox_inches='tight', facecolor='white', edgecolor='none')
    plt.close(fig)
    size_kb = Path(out_path).stat().st_size // 1024
    print(f"✓ Preview saved: {out_path} ({size_kb} KB)")

if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "demo/demo_preview.png"
    render(out)
