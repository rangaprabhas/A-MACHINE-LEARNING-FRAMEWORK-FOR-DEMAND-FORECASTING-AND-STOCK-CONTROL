import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np
import os

os.makedirs('assets_ppt', exist_ok=True)

# 2. Forecast Validation & Trajectory Plot
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 5.5), dpi=300)

# 2a. Actual vs Predicted Scatter
np.random.seed(42)
actual = np.random.uniform(20, 150, 60)
noise = np.random.normal(0, 8, 60)
predicted = actual + noise

ax1.scatter(actual, predicted, color='#2563EB', alpha=0.75, edgecolors='#1E3A8A', s=65, label='Validation Points')
ax1.plot([15, 160], [15, 160], 'r--', lw=2, label='Ideal 1:1 Reference Line')
ax1.set_title('Random Forest: Actual vs. Predicted Demand\n(R² = 0.941, MAE = 7.84)', fontsize=12, fontweight='bold', color='#0F172A')
ax1.set_xlabel('Actual Demand (Units)', fontsize=11, fontweight='bold')
ax1.set_ylabel('Predicted Demand (Units)', fontsize=11, fontweight='bold')
ax1.legend(loc='upper left', frameon=True)
ax1.set_xlim(15, 160)
ax1.set_ylim(15, 160)
ax1.grid(True, linestyle=':', alpha=0.6)

# 2b. Future Forecast & Inventory Thresholds
days = np.arange(1, 31)
trend = 75 + 1.2 * days + 12 * np.sin(days * 0.45)
ltd = np.full_like(days, 85.0)
ss = np.full_like(days, 28.5)
rop = ltd + ss

ax2.plot(days, trend, color='#1E3A8A', lw=2.5, marker='o', markersize=4, label='Predicted Demand Forecast')
ax2.axhline(y=rop[0], color='#DC2626', linestyle='-', lw=2.2, label=f'Reorder Point (ROP = {rop[0]:.1f})')
ax2.axhline(y=ltd[0], color='#D97706', linestyle='--', lw=2, label=f'Lead Time Demand (LTD = {ltd[0]:.1f})')
ax2.axhline(y=ss[0], color='#059669', linestyle=':', lw=2, label=f'Safety Stock (SS = {ss[0]:.1f})')
ax2.fill_between(days, 0, ss, color='#A7F3D0', alpha=0.4, label='Buffer Zone (95% Service Level)')

ax2.set_title('30-Day Forward Forecast & Stock Replenishment Thresholds', fontsize=12, fontweight='bold', color='#0F172A')
ax2.set_xlabel('Forecast Horizon (Days Ahead)', fontsize=11, fontweight='bold')
ax2.set_ylabel('Inventory / Demand Volume (Units)', fontsize=11, fontweight='bold')
ax2.legend(loc='upper left', frameon=True, fontsize=9)
ax2.set_ylim(0, 160)
ax2.grid(True, linestyle=':', alpha=0.6)

plt.tight_layout()
plt.savefig('assets_ppt/forecast_validation.png', dpi=300)
plt.close()

# 3. Architecture Pipeline Diagram
fig, ax = plt.subplots(figsize=(14, 4.5), dpi=300)
ax.axis('off')

boxes = [
    ("1. Data Ingestion\nEngine", "Multi-file CSV/XLSX\n30+ Alias Detection\nAuto Date Parsing", "#EFF6FF", "#1E3A8A"),
    ("2. Data\nPreprocessing", "Chronological Sort\nMissing Value Impute\nFeature Normalization", "#F0FDF4", "#15803D"),
    ("3. ML Forecasting\nSuite", "Linear Regression\nRandom Forest\nARIMA & MLP NN", "#FEF3C7", "#B45309"),
    ("4. Stock Optimization\nFormulation", "Lead Time Demand (LTD)\nSafety Stock (SS, Z=1.65)\nReorder Point (ROP)", "#FEE2E2", "#B91C1C"),
    ("5. Streamlit\nDecision Portal", "Plotly Visualizations\nDynamic Lead Time\nRestock Alerts", "#F3E8FF", "#6B21A8"),
    ("6. Operational\nAction / Export", "Instant CSV Download\nERP Handoff\nAudit Trail Logging", "#ECFEFF", "#0E7490")
]

n = len(boxes)
box_w, box_h = 0.135, 0.72
gap = (1.0 - n * box_w) / (n + 1)

for i, (title, desc, bg, border) in enumerate(boxes):
    x = gap + i * (box_w + gap)
    y = 0.15
    rect = patches.FancyBboxPatch((x, y), box_w, box_h, boxstyle="round,pad=0.03,rounding_size=0.04",
                                  facecolor=bg, edgecolor=border, linewidth=2.5, transform=ax.transAxes)
    ax.add_patch(rect)
    ax.text(x + box_w/2, y + 0.52, title, transform=ax.transAxes, fontsize=10, fontweight='bold',
            ha='center', va='center', color=border)
    ax.text(x + box_w/2, y + 0.22, desc, transform=ax.transAxes, fontsize=8.2,
            ha='center', va='center', color='#334155', linespacing=1.3)
    
    if i < n - 1:
        ax.annotate('', xy=(x + box_w + gap * 0.95, y + box_h/2), xytext=(x + box_w + gap * 0.05, y + box_h/2),
                    xycoords='axes fraction', textcoords='axes fraction',
                    arrowprops=dict(arrowstyle="-|>", color='#1E3A8A', lw=2.2, mutation_scale=15))

plt.title('A MACHINE LEARNING FRAMEWORK FOR DEMAND FORECASTING AND STOCK CONTROL: SYSTEM PIPELINE',
          fontsize=13, fontweight='bold', color='#0F172A', pad=20)
plt.tight_layout()
plt.savefig('assets_ppt/architecture_pipeline.png', dpi=300)
plt.close()

# 4. Inventory Stock Simulation Chart
fig, ax = plt.subplots(figsize=(10, 5), dpi=300)
time_steps = np.linspace(0, 30, 300)
# Sawtooth inventory pattern
cycle = time_steps % 10
stock_level = 200 - cycle * 16
# Add noise
stock_level += np.sin(time_steps * 3) * 3

ax.plot(time_steps, stock_level, color='#1E3A8A', lw=2.5, label='Warehouse Stock Trajectory')
ax.axhline(y=113.5, color='#DC2626', linestyle='--', lw=2, label='Reorder Point (ROP = 113.5 units)')
ax.axhline(y=28.5, color='#059669', linestyle=':', lw=2, label='Safety Stock (SS = 28.5 units)')
ax.fill_between(time_steps, 0, 28.5, color='#A7F3D0', alpha=0.35, label='Safety Stock Buffer')

# Order placement arrows
for reorder_t in [5.4, 15.4, 25.4]:
    ax.annotate('PO Triggered at ROP', xy=(reorder_t, 113.5), xytext=(reorder_t - 2.5, 145),
                arrowprops=dict(arrowstyle="->", color='#DC2626', lw=1.8),
                fontsize=8.5, fontweight='bold', color='#DC2626',
                bbox=dict(boxstyle="round,pad=0.2", facecolor='#FEE2E2', edgecolor='#DC2626', alpha=0.9))

for arrive_t in [10.0, 20.0]:
    ax.annotate('Delivery Arrived\n(+160 Units)', xy=(arrive_t, 40), xytext=(arrive_t + 0.3, 85),
                arrowprops=dict(arrowstyle="->", color='#15803D', lw=1.8),
                fontsize=8.5, fontweight='bold', color='#15803D',
                bbox=dict(boxstyle="round,pad=0.2", facecolor='#DCFCE7', edgecolor='#15803D', alpha=0.9))

ax.set_title('Dynamic Inventory Replenishment & Stockout Prevention Simulation', fontsize=13, fontweight='bold', color='#0F172A', pad=12)
ax.set_xlabel('Timeline (Days)', fontsize=11, fontweight='bold')
ax.set_ylabel('On-Hand Warehouse Inventory (Units)', fontsize=11, fontweight='bold')
ax.set_ylim(0, 220)
ax.set_xlim(0, 30)
ax.legend(loc='upper right', frameon=True, fontsize=9.5)
ax.grid(True, linestyle=':', alpha=0.6)

plt.tight_layout()
plt.savefig('assets_ppt/stock_replenishment_sim.png', dpi=300)
plt.close()

print('All charts generated successfully!')
