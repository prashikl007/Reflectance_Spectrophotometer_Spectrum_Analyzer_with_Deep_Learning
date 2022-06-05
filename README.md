# Relfection Spectrophotometer Spectrum Analyzer with Deep Learning
script to analyze the visible spectrum of element from camera to corresponding wavelengths.
![This is an image](https://github.com/prashikl007/Visible_Spectrum_Analyzer_with_Deep_Learning/blob/main/al%20foil.jpg)

Horizontal axis in plot spectrum graph is wavelength in 'nm', vertical axis is count for corresponding color pixels or intensity of the wavelength of that color pixel.

GUI_spectrometer.py   ------>  GUI script with complete functionality. This requires  ----> pixel_wavelength.xlsx, spectrum_model.h5

    1> pixel_wavelength.xlsx ---> contains first three column as r,g,b values of color, fourth column is wavelength for respective color ---> for this mapping script is present in Color_pxel_to_wavelength.py
    2> spectrum_model.h5 ------> neural network model for prediction.
    
    Press Analyse button to get plot
    Press store to Dataset button to store detected wavelengths shown in plot in optical_spectrum_dataset.xlsx
    Press Predict material for detecting material


Codes in following script are embedded in GUI script.

    1> Color_pxel_to_wavelength.py ---> script to map colors to respective wavelength and store it to excel file ---> pixel_wavelength.xlsx
    2> detect_wavelength.py  ---> image of spectrum is provided to this script, which generates the respective spectrum wavelength and plots it's amplitude / intensity,  --> this file uses data from pixel_wavelength.xlsx to map wavelengths of spectrum from new material or spectrum image.
    3> spectrum_classifier.py   -----> this script was used for training NN modelwhich requires --->  optical_spectrum_dataset.xlsx

optical_spectrum_dataset.xlsx  -------> wavelengths for different materials like salt(NaCl), chalk(CaCO3), table sugar. This sheet is the dataset for training NN model and can be updated by adding new spectrum wavelengths of new material with "add to dataset" button in GUI so that model can be trained for new materials.

spectrum_300pixels.jpg   is a visible spectrum used for manupulation of data. (300 pixels or row in image = 700nm to 400nm wavelength)

Training and validation curve of the NN model
![This is an image](https://github.com/prashikl007/Relfection_Spectrophotometer_Spectrum_Analyzer_with_Deep_Learning/blob/main/training_and_validation_curve.jpeg)
