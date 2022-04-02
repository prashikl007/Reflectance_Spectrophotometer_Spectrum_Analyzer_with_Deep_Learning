# Spectrum_analyzer
script to analyze the spectrum of element from camera to corresponing wavelength

Color_pxel_to_wavelength.py ---> script to map colors to respective wavelength and sotr it to excel file ---> pixel_wavelength.xlsx

pixel_wavelength.xlsx ---> contains first three column as r,g,b values of color, fourth column is wavelength for respective color ---> calculation formula is present in Color_pxel_to_wavelength.py

detect_wavelength.py  ---> image of spectrum is provided to this script, which generates the respective spectrum wavelength and plots is amplitude / intensity,  --> this file uses data from pixel_wavelength.xlsx to map wavelengths

spectrum_2.jpg is a standard visible spectrum used for manupulation of data.
