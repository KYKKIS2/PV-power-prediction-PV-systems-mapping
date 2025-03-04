import pandas as pd
import numpy as np
from scipy.signal import savgol_filter
import matplotlib.pyplot as plt
from tkinter import Tk, filedialog, Button, Label
from pathlib import Path
import tempfile,os
import matplotlib
matplotlib.use('TkAgg')  # Use TkAgg backend for GUI

import matplotlib.pyplot as plt
import numpy as np

# Function to convert Excel to CSV
def convert_excel_to_csv(excel_filepath, csv_filepath):
    # Read and store content 
    # of an excel file  
    read_file = pd.read_excel (excel_filepath) 
    
    # Write the dataframe object 
    # into csv file 
    read_file.to_csv (csv_filepath,  
                    index = None, 
                    header=True) 
        
    # read csv file and convert  
    # into a dataframe object 
    df = pd.DataFrame(pd.read_csv(csv_filepath)) 
    
    # show the dataframe 
    df

# Function to read data from CSV file with skipped rows
def read_data(filepath, skiprows=8):
    df = pd.read_csv(filepath, skiprows=skiprows)
    print("Initial DataFrame after reading CSV:")
    print(df[df.sum(axis=1) != 0].head(50))  # Print the first 50 rows with non-zero values

    df = df.rename(columns=lambda x: x.strip())
    time_df = df.iloc[:, 0]
    
    # Determine the number of days based on the actual number of columns (excluding the timestamp column)
    num_days = df.shape[1] - 1  # Subtracting 1 for the first column which is the timestamp
    
    energy_df = df.iloc[:, 1:num_days + 1]  # Select only the columns for the actual days present
    energy_df.columns = range(1, num_days + 1)  # Day columns as 1, 2, ..., num_days
    
    melted_df = pd.melt(energy_df.reset_index(), id_vars=['index'], value_vars=list(energy_df.columns),
                        var_name='Day', value_name='Power Output')
    melted_df['Timestamp'] = np.tile(time_df, len(energy_df.columns))
    melted_df['Day'] = melted_df['Day'].astype(int)
    melted_df['Year'] = 2018
    melted_df['Month'] = 3
    melted_df['Hour'] = pd.to_datetime(melted_df['Timestamp'], format='%H:%M:%S').dt.hour
    melted_df['Minute'] = pd.to_datetime(melted_df['Timestamp'], format='%H:%M:%S').dt.minute

    print("Data after reading and melting:")
    print(melted_df[melted_df['Power Output'] != 0].head(50))  # Print the first 50 rows with non-zero values

    return melted_df[['Timestamp', 'Power Output', 'Year', 'Month', 'Day', 'Hour', 'Minute']]


# Function to filter data based on given range
def filter_data(df, year_from, year_to, month_from, month_to, day_from, day_to, hour_from, hour_to):
    filtered_df = df[
        (df['Year'] >= year_from) & (df['Year'] <= year_to) &
        (df['Month'] >= month_from) & (df['Month'] <= month_to) &
        (df['Day'] >= day_from) & (df['Day'] <= day_to) &
        (df['Hour'] >= hour_from) & (df['Hour'] <= hour_to)
    ]
    
    print("Filtered data:")
    print(filtered_df[filtered_df['Power Output'] != 0].head(50))  # Print the first 50 rows with non-zero values

    return filtered_df

def mean_filter_for_values(matrix_out):
    mean_out = []
    for h in range(24):
        for m in range(0, 60, 15):
            # Get the array of power outputs for the current time slot
            a = matrix_out[(matrix_out['Hour'] == h) & (matrix_out['Minute'] == m)]['Power Output'].values
            
            # Convert to numeric, setting errors='coerce' will turn problematic entries to NaN
            a = pd.to_numeric(a, errors='coerce')
            
            # Filter out NaN values
            a = a[~np.isnan(a)]
            
            # Only proceed if the array is not empty
            if len(a) > 0:
                pOutMean = np.mean(a)
                pOutSTD = np.std(a)
                
                # Apply the filtering based on standard deviation
                pOutFiltered = a[(a > pOutMean - 1.5 * pOutSTD) & (a < pOutMean + 1.5 * pOutSTD)]
                
                # Calculate the mean of the filtered values, ignoring NaNs
                pOutFilteredMean = np.mean(pOutFiltered) if len(pOutFiltered) > 0 else 0
                
                # Further filter to only include values greater than the mean
                pOutFilter2 = pOutFiltered[pOutFiltered > pOutFilteredMean]
                
                # Calculate the final mean, setting to 0 if no valid values are left
                temp_var = np.mean(pOutFilter2) if len(pOutFilter2) > 0 else 0
                mean_out.append(temp_var)
            else:
                mean_out.append(0)  # Append 0 if no valid data is available for this slot
    
    mean_out = np.array(mean_out)
    
    print("Mean filtered values:")
    print(mean_out)
    
    return mean_out


# Function to apply Savitzky-Golay filter
def apply_sg_filter(values, window_size, poly_order):
    sgf_values = savgol_filter(values, window_size, poly_order)
    
    print("Savitzky-Golay filtered values:")
    print(sgf_values)
    
    return sgf_values

# Function to plot the results
def plot_results(mean_values, sgf_values):
    plt.figure(figsize=(10, 6))
    plt.plot(mean_values, label='Mean Filtered Values', linestyle='--')
    plt.plot(sgf_values, label='Savitzky-Golay Filtered Values', linestyle='-')
    plt.xlabel('Time (15-minute intervals)')
    plt.ylabel('Power Output (Watts)')
    plt.title('Estimated Clear Sky PV Power Production')
    plt.legend()
    plt.grid(True)
    plt.show()

# Main function to execute the process
def process_pv_data(filepath, export_filepath, year_from, year_to, month_from, month_to, day_from, day_to, hour_from, hour_to):
    # Check if the file is an Excel file and convert to CSV if necessary
    if filepath.suffix in ['.xls', '.xlsx']:
        csv_filepath = filepath.with_suffix('.csv')
        convert_excel_to_csv(filepath, csv_filepath)
    else:
        csv_filepath = filepath

    df = read_data(csv_filepath, skiprows=8)
    filtered_df = filter_data(df, year_from, year_to, month_from, month_to, day_from, day_to, hour_from, hour_to)
    mean_values = mean_filter_for_values(filtered_df)
    #you can change the window size according to your needs by changing the window_size variable on the next line.
    sgf_values = apply_sg_filter(mean_values, window_size=15, poly_order=2)
    plot_results(mean_values, sgf_values)
    
    results_df = pd.DataFrame({
        'Hour': list(range(24)) * 4,
        'Minute': [0, 15, 30, 45] * 24,
        'Mean Value': mean_values,
        'SGF Value': sgf_values
    })
    '''
    # Write to a temporary file first
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    try:
        results_df.to_excel(temp_file.name, index=False)
        temp_file.close()  # Close the file to flush it to disk
        os.replace(temp_file.name, export_filepath)  # Rename the temporary file to the final name
    except Exception as e:
        print(f"Error writing to Excel: {e}")
        os.remove(temp_file.name)  # Clean up the temporary file if something goes wrong'''

def open_file_dialog():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx"), ("CSV files", "*.csv")])
    if file_path:
        export_filepath = Path(file_path).with_suffix('.xlsx')
        process_pv_data(
            filepath=Path(file_path),
            export_filepath=export_filepath,
            year_from=2018, year_to=2018,
            month_from=3, month_to=3,
            day_from=1, day_to=31,
            hour_from=0, hour_to=23
        )

def create_gui():
    root = Tk()
    root.title("PV Energy Prediction")
    root.geometry("300x100")

    label = Label(root, text="Select a CSV or Excel file for processing")
    label.pack(pady=10)

    button = Button(root, text="Load File", command=open_file_dialog)
    button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
