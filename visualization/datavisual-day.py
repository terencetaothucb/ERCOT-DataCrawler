import os
import pandas as pd
import matplotlib.pyplot as plt
from datetime import timedelta
import pytz

# --- 1) Read and concatenate all Excel files with progress logging ---
data_folder = 'data'
columns_of_interest = [
    'Time',
    'REG-UP-Deployed', 'REG-UP-Undeployed',
    'REG-DOWN-Deployed', 'REG-DOWN-Undeployed',
    'RRS', 'NON-SPIN', 'ECRS', 'Frequency'
]

file_list = sorted(f for f in os.listdir(data_folder) if f.endswith('.xlsx'))
total_files = len(file_list)
print(f"Found {total_files} Excel files in '{data_folder}'.")

all_data = []
for idx, filename in enumerate(file_list, start=1):
    print(f"[{idx}/{total_files}] Reading file: {filename}")
    df = pd.read_excel(
        os.path.join(data_folder, filename),
        usecols=columns_of_interest,
        engine='openpyxl'
    )
    # parse Time column as naive datetime (UTC assumed)
    df['Time'] = pd.to_datetime(df['Time'], errors='coerce')
    all_data.append(df.dropna(subset=['Time']))
print("All files read and appended.\n")

# --- 2) Merge, sort, and dedupe timestamps ---
print("Merging dataframes...")
data = pd.concat(all_data, ignore_index=True)
print(f"  -> Merged shape: {data.shape}")

print("Sorting by Time and dropping duplicates/missing timestamps...")
data = (
    data
    .sort_values('Time')
    .drop_duplicates(subset=['Time'])
    .reset_index(drop=True)
)
print(f"  -> After sort & dedupe: {data.shape}\n")

# --- 3) Compute median interval and interpolate per continuous segment ---
print("Calculating median sampling interval...")
median_delta = data['Time'].diff().dt.total_seconds().dropna().median()
interval = int(median_delta)
print(f"  -> Median interval ≈ {interval} seconds")

print("Identifying gaps and segmenting data...")
# flag gaps larger than twice the median interval
data['is_gap'] = data['Time'].diff() > pd.Timedelta(seconds=interval * 2)
data['segment'] = data['is_gap'].cumsum()

print("Interpolating each segment separately...")
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
print(f"  -> After segmented interpolation: {data.shape}\n")

# --- 4) Convert timestamps from UTC to US/Central and drop tzinfo ---
UTC     = pytz.UTC
CENTRAL = pytz.timezone('America/Chicago')
data['Time'] = (
    data['Time']
    .dt.tz_localize(UTC)
    .dt.tz_convert(CENTRAL)
    .dt.tz_localize(None)
)
print("Timestamps converted to US/Central.\n")

# --- 5) Prepare output directories for daily plots ---
output_base = 'plots-day'
os.makedirs(output_base, exist_ok=True)
print(f"Output base directory: '{output_base}'\n")

variables = [
    'REG-UP-Deployed', 'REG-UP-Undeployed',
    'REG-DOWN-Deployed', 'REG-DOWN-Undeployed',
    'RRS', 'NON-SPIN', 'ECRS', 'Frequency'
]

unique_dates = sorted(data['Time'].dt.date.unique())
print(f"Found data for {len(unique_dates)} unique dates: {unique_dates}\n")

# --- 6) Generate daily plots with uniform hourly x-axis ---
for var in variables:
    safe_var = var.replace('/', '-').replace(' ', '_')
    var_folder = os.path.join(output_base, safe_var)
    os.makedirs(var_folder, exist_ok=True)
    print(f"Processing variable '{var}' -> folder '{var_folder}'")

    for date_idx, current_date in enumerate(unique_dates, start=1):
        print(f"  [{date_idx}/{len(unique_dates)}] Plotting {var} for {current_date}")
        day_start = pd.Timestamp(current_date)
        day_end   = day_start + timedelta(days=1)

        # select data within the 24-hour window
        df_day = data[(data['Time'] >= day_start) & (data['Time'] < day_end)]
        if df_day.empty:
            hours = []
            values = []
        else:
            # convert timestamps to hours since midnight
            hours  = (df_day['Time'] - day_start).dt.total_seconds() / 3600
            values = df_day[var]

        fig, ax = plt.subplots(figsize=(10, 6))
        ax.plot(hours, values, linewidth=1.5)

        date_str = day_start.strftime('%Y-%m-%d')
        ax.set_title(f'{var} — {date_str}', fontsize=14)
        ax.set_xlabel('Time (h)', fontsize=12)
        ax.set_ylabel(
            'Frequency (Hz)' if var == 'Frequency' else f'{var} (MW)',
            fontsize=12
        )

        # fix x-axis from 0 to 24 hours with 2-hour ticks
        ax.set_xlim(0, 24)
        ax.set_xticks(range(0, 25, 2))

        ax.grid(True)
        ax.tick_params(axis='x', rotation=0, labelsize=10)
        ax.tick_params(axis='y', labelsize=10)

        plt.tight_layout(pad=2.0)
        save_name = f'{safe_var}_{date_str}.png'
        fig_path = os.path.join(var_folder, save_name)
        plt.savefig(fig_path, dpi=300)
        plt.close(fig)

    print(f"  -> Completed plots for '{var}'.\n")

print("All daily plots have been generated and saved.")