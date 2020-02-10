''' Open GUM's I-V files and save to Excel.
Jeremy Gillbanks, 10 Feb 2020.
Takes a directory location and returns all results to a single workbook for plotting.
It IGNORES the status file.
Reduces time manually importing files, using text to columns, etc.'''


# List the input folder:
file_location = "/Volumes/GoogleDrive/My Drive/PhD/20200201 SiNx water ingress raw results/20200206 I-V scans for SiNx (initial)/jeremy/F1_lowfield/"
# Write the intended output file directory:
output_file_directory = '/'.join(file_location.split("/")[:-2]) # for my convenience it moves up one file level
output_filename = "20200206 Initial SiNx results.xlsx"

# Nothing should need to be touched beyond this point (hopefully!)

import pandas as pd
import os

def folder_check(file_location):
	''' Ensure there is a '/' at the end of the file location. '''
	if file_location[-1] != "/":
		file_location += "/"
	return file_location

def convert_txt_to_dataframe(file_location, filename):
	''' Convert output txt file to dataframe for manipulation. '''
	file_location = folder_check(file_location)
	data = pd.read_fwf(file_location + filename, delimiter="\t")
	# Get date & time before dataframe manipulation
	try:
		date = data.columns[0].split("[")[1].split("]")[0]
		time = data.columns[0].split("[")[2].split("]")[0]
	except:
		print(data.columns[0].split("["))
		# raise e

	data.dropna(inplace = True) # Ignore N/A values
	data = data.iloc[:,0].str.split("\t", n = 1, expand = True) # Remove tab-delimiter from first column
	data.iloc[0][1] = filename # Insert the filename (for posterity) into the 1st row, 2nd column
	data.columns = [date, time] # Rename the column names to keep date & time data
	return data

def main():
	output_dataframe = pd.DataFrame()
	output_directory = folder_check(file_location = output_file_directory)
	for file in os.listdir(file_location):
		if "Status" not in file:
			print("Converting", file)
			data = convert_txt_to_dataframe(file_location = file_location, filename = file)
			print("Concatenating", file, "to dataframe.")
			output_dataframe = pd.concat([output_dataframe, data], sort = False, axis = 1)
			# print(output_dataframe.head())
	print(output_dataframe.head())
	
	output_dataframe.to_excel(output_directory + output_filename, sheet_name = "All results")

if __name__ == '__main__':
	main()