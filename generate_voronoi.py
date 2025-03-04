import numpy as np
import pandas as pd
import geopandas as gpd

from shapely.ops import unary_union
from shapely.geometry import Polygon
from geovoronoi import voronoi_regions_from_coords, points_to_coords

# Import tkinter for GUI message boxes
import tkinter as tk
from tkinter import messagebox

# Import matplotlib for plotting
import matplotlib.pyplot as plt

# Initialize tkinter (needed to display message boxes)
root = tk.Tk()
root.attributes('-topmost', True)  # Bring the root window to the front
root.withdraw()  # Hide the root window (after setting attributes)

# Step 1: Reading energy production data and assigning proper column names
df_data = pd.read_excel(
    '11.2021.xlsx',
    sheet_name='All Data',
    header=3
)
df_data.drop(index=0, inplace=True)
df_data.rename(columns={'Unnamed: 0': 'DateTime'}, inplace=True)
df_data.fillna(0, inplace=True)

# Step 2: Assigning date and time stamps at 15-minute intervals to the 'DateTime' column
initial_timestamp = pd.to_datetime("1-11-2021 00:00", format="%d-%m-%Y %H:%M")
df_data['DateTime'] = pd.date_range(initial_timestamp, freq='15T', periods=len(df_data))
df_data.set_index('DateTime', inplace=True)

# Step 3: Reading the table with PV plant coordinates
df_coord = pd.read_excel(
    'PV Plants (address).xlsx',
    header=2
)

# Step 4: Adding proper column names to the coordinate DataFrame
coord_rename = {
    'Plant ': 'Plant',
    'Coordinates (dd)': 'lat',
    'Unnamed: 5': 'lon',
}
df_coord.rename(columns=coord_rename, inplace=True)
df_coord = df_coord[['Plant', 'lat', 'lon']]

# Step 5: Transforming the energy data to a long format and handling errors
df_data = df_data.melt(var_name='Plant', value_name='energy', ignore_index=False).reset_index()
error_count = df_data[pd.to_numeric(df_data.energy, errors='coerce').isnull()].shape[0]
df_data.loc[pd.to_numeric(df_data.energy, errors='coerce').isnull(), 'energy'] = 0

# Displaying the error count in a message box if there are any errors
if error_count > 0:
    messagebox.showwarning("Data Entry Correction", f"{error_count} records fixed due to wrong data entry", parent=root)
else:
    print(f'{error_count} records fixed due to wrong data entry')

# Step 6: Merging the energy data with coordinates (Modified)
df_plants = pd.merge(df_data, df_coord, on='Plant', how='outer', indicator=True)

# Convert '_merge' column to string to avoid TypeError during file write
df_plants['_merge'] = df_plants['_merge'].astype(str)

# Step 6a: Identifying plants missing in one of the Excel files
# Plants missing in coordinate data
missing_in_coord = df_plants[df_plants['_merge'] == 'left_only']['Plant'].unique()
# Plants missing in energy data
missing_in_data = df_plants[df_plants['_merge'] == 'right_only']['Plant'].unique()

# Step 6b: Issuing warning messages for missing plants using message boxes
if len(missing_in_coord) > 0:
    warning_msg_coord = "The following plants are missing in the coordinates Excel and will not appear on the map:\n"
    for plant in missing_in_coord:
        warning_msg_coord += f" - {plant}\n"
    messagebox.showwarning("Missing Coordinates", warning_msg_coord.strip(), parent=root)

if len(missing_in_data) > 0:
    warning_msg_data = "The following plants are missing in the energy data Excel but will appear on the map:\n"
    for plant in missing_in_data:
        warning_msg_data += f" - {plant}\n"
    messagebox.showwarning("Missing Energy Data", warning_msg_data.strip(), parent=root)

# Proceed with the merged DataFrame
# For plants missing coordinates, 'lat' and 'lon' will be NaN

# Step 7: Sorting and resetting the index of the merged DataFrame
df_plants.sort_values('DateTime', inplace=True)
df_plants.reset_index(drop=True, inplace=True)

# Step 8: Including all plants with coordinates
df_plants_i = df_plants.copy()
date_i = pd.to_datetime("1-11-2021 00:00", format="%d-%m-%Y %H:%M")
df_plants_i['DateTime'] = df_plants_i['DateTime'].fillna(date_i)
df_plants_i['energy'] = df_plants_i['energy'].fillna(0)

# Step 9: Handling missing coordinates before creating the GeoDataFrame
# Drop rows where 'lat' or 'lon' is NaN because we cannot create geometries without coordinates
df_plants_i_nonan = df_plants_i.dropna(subset=['lat', 'lon'])

# Create a separate DataFrame for plants missing coordinates (for reporting)
df_plants_i_nan = df_plants_i[df_plants_i['lat'].isna() | df_plants_i['lon'].isna()]

if not df_plants_i_nan.empty:
    warning_msg_nan = f"The following plants lack coordinate data and will not be mapped:\n"
    for plant in df_plants_i_nan['Plant'].unique():
        warning_msg_nan += f" - {plant}\n"
    messagebox.showwarning("Missing Coordinates at Timestamp", warning_msg_nan.strip(), parent=root)

# Step 10: Creating a GeoDataFrame for the selected plants with coordinates
gdf_plants_i = gpd.GeoDataFrame(
    df_plants_i_nonan,
    geometry=gpd.points_from_xy(df_plants_i_nonan.lon, df_plants_i_nonan.lat)
)

# Step 11: Defining the bounding box for the area of interest
bbox = [33.2, 34.9, 33.5, 35.2]
margin = 0.005

bbox_x = [bbox[0]-margin, bbox[0]-margin, bbox[2]+margin, bbox[2]+margin]
bbox_y = [bbox[1]-margin, bbox[3]+margin, bbox[3]+margin, bbox[1]-margin]

# Step 12: Creating the bounding box polygon
polygon_geom = Polygon(zip(bbox_x, bbox_y))
bbox_polygon = gpd.GeoDataFrame(index=[0], crs='epsg:4326', geometry=[polygon_geom])

# Step 13: Filtering the GeoDataFrame of plants within the bounding box
gdf_plants_i = gdf_plants_i.set_crs('EPSG:4326')
gdf_plants_i = gdf_plants_i[
    (gdf_plants_i['lat'] >= bbox[1]) & (gdf_plants_i['lat'] <= bbox[3]) &
    (gdf_plants_i['lon'] >= bbox[0]) & (gdf_plants_i['lon'] <= bbox[2])
]

# Step 14: Creating the boundary shape for Voronoi diagram generation
boundary_shape = unary_union(bbox_polygon.geometry)

# Step 15: Generating Voronoi regions from the plant coordinates
coords = points_to_coords(gdf_plants_i.geometry)
poly_shapes, pts = voronoi_regions_from_coords(coords, boundary_shape)

# Remove unnecessary print to improve runtime
# print(coords)

# Step 16: Creating a GeoDataFrame of Voronoi polygons
poly_list = [v for k, v in poly_shapes.items()]
voronoi = gpd.GeoDataFrame(crs='epsg:4326', geometry=poly_list)

# Step 17: Plotting the Voronoi regions and plant locations
# Create a figure and axis with a specified size
fig1, ax1 = plt.subplots(figsize=(10, 10))

# Plot the Voronoi polygons with thicker edges
voronoi.plot(ax=ax1, color='white', edgecolor='black', linewidth=1)

# Plot the PV plant locations with larger markers
gdf_plants_i.plot(ax=ax1, color='red', markersize=50)

# Add titles and labels for clarity
ax1.set_title('Voronoi Diagram of PV Systems')
ax1.set_xlabel('Longitude')
ax1.set_ylabel('Latitude')

# Display the plot without blocking
plt.show(block=False)

# Step 18: Spatial join between Voronoi polygons and plant data
gdf_voronoi = voronoi.sjoin(gdf_plants_i, how='left')

# Drop unnecessary or problematic columns before writing to file
columns_to_drop = ['DateTime', '_merge', 'index_right']
for col in columns_to_drop:
    if col in gdf_voronoi.columns:
        gdf_voronoi.drop(columns=col, inplace=True)

# Aggregate the data to remove duplicates
gdf_voronoi = gdf_voronoi.dissolve(by='Plant', as_index=False)

# Step 19: Saving the Voronoi regions to a GeoJSON file
output_path = 'solar-voroni_simplified.geojson'
gdf_voronoi.to_file(output_path, driver='GeoJSON')

# Additional Step: Plotting the GeoJSON file
# Load the GeoJSON file
gdf = gpd.read_file(output_path)

# Ensure 'Plant' is a categorical variable
gdf['Plant'] = gdf['Plant'].astype('category')

# Plot the GeoDataFrame, coloring by 'Plant'
fig2, ax2 = plt.subplots(figsize=(10, 10))
gdf.plot(column='Plant', legend=True, ax=ax2, cmap='Set3')
ax2.set_title('Voronoi Regions Colored by Plant')
ax2.set_xlabel('Longitude')
ax2.set_ylabel('Latitude')

# Display the second plot without blocking
plt.show(block=False)

# Optional: Keep the script running to prevent figures from closing
input("Press Enter to exit and close figures...")

# Close the tkinter root window after use
root.destroy()
