import pandas as pd
import numpy as np
import os
import openpyxl

# For Radiant at C209 with a 700-volt external amplifier

def importdata(path, header, lines, name = None):
	table = pd.read_csv(path, sep = '\t', header = header, nrows = lines, names = name, encoding = 'unicode_escape')
	return table

def discrete_integral(x, y):
	z = 0
	for index in range(len(x) - 1):
		z = z + 0.5 * (y[index] + y[index + 1]) * (x[index + 1] - x[index])
	return z

#Deal with inputs with Input.txt
input = pd.read_table('Input_Cal.txt', delim_whitespace = True, header = None, nrows = 2)
correctinput = pd.read_table('Input_Cal.txt', delim_whitespace = True, header = 1, names = [1, 2, 3])

path_root = input.values[0, 0] #Main directory which doesn't change very often
path_sample = input.values[1, 0] #Such a file structure is highly preferred here for the convenience of further explict naming of output files
path = path_root + path_sample + '\\'
minimum = correctinput.values[1,0]

ifcorrectsample = correctinput.values[0, 0]   #If to correct geometric parameters? 1-Yes, 0-No
thick_actual = correctinput.values[0, 1]   #Actual thickness as indicated by SEM
elec_actual = correctinput.values[0, 2]   #Actual electrode area derived from optical microscopes

filename1 = 'UCal_' + path_sample.replace('\\', '_') + '.xlsx' #This file appears to be "UCal_date_samplename.xlsx" which contains energy storage results, with a Criteria sheet manifesting some key maximum values
filename2 = 'PE_' + path_sample.replace('\\', '_') + '.xlsx' #This file appears to be "PE_date_samplename.x     lsx" which contains PE loop data in the dictionary order of filenames from left to right
filename3 = 'Info_' + path_sample.replace('\\', '_') + '.xlsx' #This file appears to be "Info_date_samplename.xlsx" which contains some parameters of the sample and measurements. You can check here if the UCal file is geometrically corrected

writer1 = pd.ExcelWriter(path + filename1)
writer2 = pd.ExcelWriter(path + filename2)
writer3 = pd.ExcelWriter(path + filename3)

key_criteria = pd.DataFrame(columns = ["FoldNum", "MaxENominal", "MaxEActual", "MaxPm", "MaxUe", "MaxFoM"])

dot_list = os.listdir(path)
for dot in dot_list:
	if os.path.isfile(os.path.join(path, dot)):
		dot_list.remove(dot)
		continue
	else:
		local_path = path + dot + '\\'
		efield_list = os.listdir(local_path)
		result = pd.DataFrame(columns = ["FileName", "NominalElectricField", "ActualNominalElectricField", "MaxElectricField", "MaxPolarization", "RemanentPolarization", "ChargedEnergy", "DischargedEnergy", "Efficiency", "FigureOfMerit"])
		sample_info = pd.DataFrame(columns = ["FileName", "NominalElectricField", "InputElectrode", "InputThickness", "IfCorrectSample", "ActualElectrode", "ActualThickness", "MeasureFieldOrVoltage", "PlotFieldOrVoltage", "Frequency", "LoopType", "NumberOfPoints"])
		pe = pd.DataFrame(columns = ['Initiate'])

		for efield in efield_list:
			if efield[-3] == 'd':
				continue
			final_path = local_path + efield
			efield_name = efield.replace('.txt', '')
			
			#import data
			geo_para = importdata(final_path, 21, 2) #Input geometric parameters
			amplifier = importdata(final_path, 24, 1)
			ifinternal = amplifier.values[0, 1] == 'Internal'
			mea_mode = importdata(final_path, 34 - 3 * ifinternal, 6) #Measurement modes (e.g. Voltage or Field specified, and P-E loop profile)
			num_pts = importdata(final_path, 41 - 3 * ifinternal, 1) #Number of points
			plot_info = importdata(final_path, 43 - 3 * ifinternal, 1, name = [1, 2, 3, 4])

			#sample info
			elec_set = geo_para.values[0,1]
			thick_set = geo_para.values[1,1] * 1000 #Convert to nm as unit
			mea_field = mea_mode.values[0,0] != 'Volts:'
			mea_field_out = mea_field * 'Field' + (1 - mea_field) * 'Volt'
			applied_field = float(mea_mode.values[1, 1].replace(' (kV/cm)', ''))
			plot_field = plot_info.values[0,2] != 'Drive Voltage'
			plot_field_out = plot_field * 'Field' + (1 - plot_field) * 'Volt'
			perd = mea_mode.values[2,1]
			freq = 1000 / float(perd)
			pe_type = mea_mode.values[5,1]
			points = num_pts.values[0,0]
			sample_info.loc[efield_list.index(efield)] = [efield_name, applied_field, elec_set, thick_set, ifcorrectsample, elec_actual, thick_actual, mea_field_out, plot_field_out, freq, pe_type, points]
			
			#main data
			main_data = importdata(final_path, 44 - 3 * ifinternal, points)
			thick = thick_set * ifcorrectsample + thick_actual * (1 - ifcorrectsample)
			elec = elec_set * ifcorrectsample + elec_actual * (1 - ifcorrectsample)
			actual_applied_field = applied_field * thick / thick_actual
			if plot_field:
				c = main_data.values[:,2] * thick / thick_actual
			else:
				c = main_data.values[:,2] / thick_set * 10000 * thick / thick_actual
			d = main_data.values[:,3] * elec / elec_actual
			c = c.tolist()
			d = d.tolist()
			emax = max(c)
			pm = d[c.index(emax)]
			if pm == 0 or d[0] < -200:
				result.loc[efield_list.index(efield)] = [efield_name, applied_field, actual_applied_field, 0, 0, 0, 0, 0, 0, 0]
				continue
			if pe_type == 'Standard Bipolar':
				c_abs = list(map(abs, c))
				c0 = c_abs.index(min(c_abs[c.index(emax):(len(c) - c.index(emax))]))
				while c0 < c.index(emax):
					c_abs[c0] = -121
					c0 = c_abs.index(min(c_abs[c.index(emax):(len(c) - c.index(emax))]))
				pr = importdata(final_path, 45 + points - 3 * ifinternal, 2)
				if isinstance(pr.values[0, 0], str) == 0:
					pr = importdata(final_path, 46 + points - 3 * ifinternal, 2)
				p2r = (pr.values[0, 1] - pr.values[1, 1]) / 2
			if pe_type == 'Standard Monopolar':
				p0r = d[0]
				d[:] = [x - p0r for x in d]
				p2r = d[-1]
				c0 = len(c) - 1
				pm = d[c.index(emax)]
			if abs(pm) < minimum:   
				d = [i * 10 for i in d]
				pm = 10 * pm
				p2r = 10 * p2r
			c1 = c[0:c.index(emax) + 1]
			d1 = d[0:c.index(emax) + 1]
			c2 = c[c.index(emax):c0 + 1]
			d2 = d[c.index(emax):c0 + 1]

			u = discrete_integral(d1, c1) / 1000
			ue = discrete_integral(d2, c2) / -1000
			if ue != 0:
				yita = ue / u
				fom = ue / (1 - yita)
			result.loc[efield_list.index(efield)] = [efield_name, applied_field, actual_applied_field, emax, pm, p2r, u, ue, yita, fom]
			
			#PE
			pe = pd.concat([pe, pd.DataFrame({'': c}), pd.DataFrame({efield_name: d})], axis = 1)
			pe.fillna(0)

	result = result.sort_values("NominalElectricField")
	sample_info = sample_info.sort_values("NominalElectricField")
	key_criteria.loc[dot_list.index(dot)] = [dot, max(result.values[:, 1]), max(result.values[:, 2]), max(result.values[:, 4]), max(result.values[:, 7]), max(result.values[:, 9])]
	pe = pe.drop(columns = ['Initiate'])

	#Excel writing
	result.to_excel(writer1, sheet_name = dot, index = False)
	sample_info.to_excel(writer3, sheet_name = dot, index = False)
	pe.to_excel(writer2, sheet_name = dot, index = False)

key_criteria.to_excel(writer1, sheet_name = 'Criteria', index = False)

writer1.save()
writer2.save()
writer3.save()
writer1.close()
writer2.close()
writer3.close()
print('Done!')