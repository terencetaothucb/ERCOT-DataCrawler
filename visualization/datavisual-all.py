import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import timedelta
import pytz

# 1) Define data folder and columns
DATA_FOLDER = 'data'
COLUMNS = [
    'Time',
    'REG-UP-Deployed', 'REG-UP-Undeployed',
    'REG-DOWN-Deployed', 'REG-DOWN-Undeployed',
    'RRS', 'NON-SPIN', 'ECRS', 'Frequency'
]

# 2) Read all Excel files and parse Time
files = sorted(f for f in os.listdir(DATA_FOLDER) if f.endswith('.xlsx'))
print(f"Found {len(files)} .xlsx files in '{DATA_FOLDER}'")

all_data = []
for i, fn in enumerate(files, 1):
    print(f"[{i}/{len(files)}] Loading {fn}")
    df = pd.read_excel(
        os.path.join(DATA_FOLDER, fn),
        usecols=COLUMNS,
        engine='openpyxl'
    )
    # parse Time column as naive datetime (assumed UTC)
    df['Time'] = pd.to_datetime(df['Time'], errors='coerce')
    all_data.append(df.dropna(subset=['Time']))
print("All files loaded.\n")

# 3) Concatenate, sort, drop duplicates
data = pd.concat(all_data, ignore_index=True)
data = (
    data
    .sort_values('Time')
    .drop_duplicates(subset=['Time'])
    .reset_index(drop=True)
)
print(f"After merge & dedupe: {data.shape}")

# 4) Compute median sampling interval (seconds)
median_delta = data['Time'].diff().dt.total_seconds().dropna().median()
interval = int(median_delta)
print(f"Median interval ≈ {interval} seconds")

# 5) Identify segments where gap > 2 * median_delta
data['gap'] = data['Time'].diff() > pd.Timedelta(seconds=interval * 2)
data['segment'] = data['gap'].cumsum()

# 6) Resample and interpolate each segment separately
segments = []
for seg_id, grp in data.groupby('segment'):
    start_time = grp['Time'].iloc[0]
    end_time   = grp['Time'].iloc[-1]
    # create uniform time index for this segment
    idx = pd.date_range(start=start_time, end=end_time, freq=f'{interval}S')
    seg = (
        grp
        .set_index('Time')
        .reindex(idx)
        .interpolate(method='time')
        .reset_index()
        .rename(columns={'index': 'Time'})
    )
    segments.append(seg)

# combine all segments back together
data = pd.concat(segments, ignore_index=True)
print(f"After segmented interpolation: {data.shape}\n")

# 7) Convert timestamps from UTC to US/Central
UTC = pytz.UTC
US_CENTRAL = pytz.timezone('America/Chicago')
data['Time'] = (
    data['Time']
    .dt.tz_localize(UTC)
    .dt.tz_convert(US_CENTRAL)
)
print("Timestamps converted to US/Central.\n")

# 8) Plot each variable with local time on x-axis
OUTPUT = 'plots-all'
os.makedirs(OUTPUT, exist_ok=True)

variables = [
    'REG-UP-Deployed', 'REG-UP-Undeployed',
    'REG-DOWN-Deployed', 'REG-DOWN-Undeployed',
    'RRS', 'NON-SPIN', 'ECRS', 'Frequency'
]

for var in variables:
    print(f"Plotting {var} …")
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(data['Time'], data[var], linewidth=1.5)

    # set title and axis labels
    ax.set_title(var, fontsize=14)
    ax.set_xlabel('Time (US/Central)', fontsize=12)
    ax.set_ylabel('Frequency (Hz)' if var == 'Frequency' else f'{var} (MW)', fontsize=12)

    # format x-axis as local date & time
    ax.xaxis.set_major_locator(mdates.AutoDateLocator())
    ax.xaxis.set_major_formatter(
        mdates.DateFormatter('%m-%d %H:%M', tz=US_CENTRAL)
    )
    fig.autofmt_xdate(rotation=30, ha='right')

    ax.grid(True)
    ax.tick_params(axis='both', labelsize=10)
    plt.tight_layout(pad=2.0)

    safe = var.replace('/', '-').replace(' ', '_')
    fig.savefig(os.path.join(OUTPUT, f'{safe}.png'), dpi=300)
    plt.close(fig)

print(f"All plots saved to '{OUTPUT}'.")