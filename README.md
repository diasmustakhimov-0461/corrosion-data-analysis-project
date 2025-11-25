# Corrosion Data Analysis Project

This project provides tools and notebooks for analyzing Electrochemical Impedance Spectroscopy (EIS) data used in corrosion research.  
The workflow supports:

- Loading multiple Excel files containing EIS measurements
- Extracting multiple tables from each sheet using row/column ranges
- Generating Nyquist, Bode Magnitude, and Bode Phase plots
- Organizing data, scripts, and plots in a reproducible structure

# Project Structure
corrosion-data-analysis-project
 -data/ 	#Raw EIS Excel files
 -notebooks/	 #Jupyter notebooks for analysis
 -plots/	 #Saved plots
 -src/ 		#Python modules
 -README.md	 #Project documentation
 -requirements.txt #Python dependencies
 -gitignore

Future Improvements:
-Automated detection of table boundaries
-Batch plotting for entire datasets
-Parameter fitting (equivalent circuits)
-Exporting results to CSV/Excel
