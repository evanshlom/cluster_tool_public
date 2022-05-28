# cluster_tool.py
# Cluster Tool 0.0.1
# Evan Shlom

from statistics import median
import pandas as pd
import openpyxl
import streamlit as st



# Webpage Header
st.title('Cluster Tool')

# Upload Button
# Prompt user to upload a file
# Create dataset object using uploaded table
# Uploaded file MUST use the Cluster Template excel file which is provided to the user
uploaded_file = st.file_uploader(
  "Upload Cluster Tool Template",
  help="To download a copy of the Cluster Tool Template, go to Facilities Planning SharePoint > SL Solutions Tools > AI Solutions Tools"
)

# Slider
# User can drag slider to set number of clusters
# The model will take num_clusters as the parameter for number of clusters in a KMeans model
num_clusters = st.slider(
  "Choose the number of clusters to calculate for your dataset (Most common is 2 to 4)",
  min_value = 2,
  max_value = 20,
  value = 2,
  help = "The program will use whichever number the slider is set to and you can change the slider whenever you want")
st.caption("Consider adding clusters to account for outliers")

# Check User Uploaded File
# Program waits here until user uploads their file
if uploaded_file is None:
  st.stop()



# Read Dataset
# Uploaded file MUST use the Cluster Template
df = pd.read_excel(uploaded_file, sheet_name='Inputs', skiprows=3)

# Show Uploaded Data
# User can view their inputs in a dataframe on the webpage 
st.subheader("Inputs")
# Restrict table height because this is less important to the user and mainly exists for the user to qualify a successful upload
st.dataframe(df, height=100)

# Remove 2nd variable column from the dataframe if the user did not provide data for a second variable
# Cluster Template instructed user to leave the 2nd variable column blank if they do not need a 2nd variable, in which case this column should show nans when uploaded
# Check if column 2 has ANY nans
if df.iloc[:,2].isnull().values.any():
  # Drop 2nd variable column if contains any nans
  # User needed to populate every cell in the 2nd variable column if they want to use a 2nd variable
  df.drop(df.columns[2],axis=1,inplace=True)

# Prepare NumPy array for the KMeans clustering model
import numpy as np
# Cluster Tool clusters rows based on the 1st Variable column, and optionally the 2nd Variable column if user included a 2nd variable in the Cluster Template
array = df.iloc[:, 1:2]

from sklearn.cluster import KMeans
# n_clusters takes the value decided by the user from the Slider at the top of the webpage
kmeans = KMeans(n_clusters=num_clusters)
kmeans.fit(array)
centroids = kmeans.cluster_centers_
labels = kmeans.labels_



# Show Results
st.header("Results")
# Show Cluster Centroids
centroids_df = pd.DataFrame(
  centroids,
  columns=["Cluster Average"]).sort_values(by=['Cluster Average']
)
st.subheader("Centroids")
st.table(centroids_df.style.set_precision(2)) #round to hundredths

# Calculate Cluster Boundaries
# Any KMeans cluster boundary is the median between two neighboring centroids
# Define pairwise function (it was not working out of the box from itertools so manually define it for now)
# Enables iteration between neighboring centroids
import itertools as it
def pairwise(iterable):
  # pairwise('ABCDEFG') --> AB BC CD DE EF FG
  a, b = it.tee(iterable)
  next(b, None)
  return zip(a, b)
# Sort centroids array to iterate pairwise for medians between neighboring centroids
centroids_sorted = np.sort(centroids, axis=0)
# Iterate for boundaries
boundaries = []
from statistics import median
for i,j in pairwise(centroids_sorted):
  boundaries.append(median([i,j]))
boundaries_df = pd.DataFrame(boundaries, columns=["Cluster Boundaries"])

# Show Cluster Boundaries
st.subheader("Cluster Boundaries")
st.caption('Cluster boundary values are usually more useful than centroid values')
st.table(boundaries_df.style.set_precision(2))
st.caption(
  '''You can alternatively think of the
  cluster boundary as a "threshold" from
  one cluster to an adjacent cluster'''
)



# Process DataFrame for User Download
# Prepare New Excel Workbook
# User will be able to download this file
import xlsxwriter
from pandas.io.formats.excel import ExcelFormatter
from io import BytesIO
# Write file to in-memory string using BytesIO
# See: https://xlsxwriter.readthedocs.io/workbook.html?highlight=BytesIO#constructor
output = BytesIO()
workbook = xlsxwriter.Workbook(output, {'in_memory': True})
# Name worksheet
worksheet = workbook.add_worksheet(name="Outputs")
# Add cluster labels to dataframe
df['Cluster'] = labels
# Copy dataframe to worksheet 
#https://stackoverflow.com/questions/50864881/write-pandas-dataframe-to-excel-with-xlsxwriter-and-include-write-rich-string:
worksheet.add_write_handler(list, lambda worksheet, row, col, args: worksheet.write_rich_string(row, col, *args))
cells = ExcelFormatter(df, index=False).get_formatted_cells()
for cell in cells:
    worksheet.write(cell.row, cell.col ,cell.val)
workbook.close()

# Show Final DataFrame
st.subheader("Cluster Labels")
# Show Download Button for DataFrame
st.download_button(
    label="Click Here to Download processed dataset with cluster labels",
    data=output.getvalue(),
    file_name="Clustered Dataset.xlsx",
    mime="application/vnd.ms-excel"
)
# Show DataFrame
# User can highlight-copy-paste from this table instead of downloading a file
st.caption("Use download button or use table below to highlight, copy and paste results")
st.table(df.style.set_precision(2))



# Show Ballons Animation
# Signals that this program successfully ran
st.balloons()



# cluster_tool.py
# Evan Shlom