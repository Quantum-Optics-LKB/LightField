# Import the .NET class library
import clr

# Import python sys module
import sys

# Import other libraries
import os
import pandas as pd
import matplotlib.pyplot as plt
import subprocess

# Import System.IO for saving and opening files
from System.IO import *

# Import C compatible List and String
from System import String
from System.Collections.Generic import List

# Add needed dll references
sys.path.append(os.environ['LIGHTFIELD_ROOT'])
sys.path.append(os.environ['LIGHTFIELD_ROOT']+"\\AddInViews")
clr.AddReference('PrincetonInstruments.LightFieldViewV5')
clr.AddReference('PrincetonInstruments.LightField.AutomationV5')
clr.AddReference('PrincetonInstruments.LightFieldAddInSupportServices')

# PI imports
from PrincetonInstruments.LightField.Automation import Automation
from PrincetonInstruments.LightField.AddIns import SpectrometerSettings, CameraSettings, DeviceType, ExperimentSettings, RegionOfInterest


class Spectrometer:
    """
    Class to control PrincetonIntruments Acton SP500i spectrometer and Pixis camera using Lightfield.
    This API enables to control LightField software with or without the GUI displayed the GUI option. 
    To work administrator permissions are needed.
    """
    def __init__(self):
        self.auto = None
        self.experiment = None
        self.experiment_name = None
        self.ROI = None
    
    def launch_experiment(self, experiment: str = None, interface: bool = True):
        """Launch the connection with Lightfield.
        Args:
            experiment (str, optional): Eperiment name (the experiment files are in the experiment folder). Defaults to None, the last experiment will be launched.
            interface (bool, optional): If interface is True, the GUI should be displayed. Defaults to True.
        """
        try:
            self.auto = Automation(interface, List[String]())
            self.experiment = self.auto.LightFieldApplication.Experiment
            if experiment is not None:
                self.experiment.Load(experiment)
                self.experiment_name = self.experiment.get_Name()
            print(f'Success, experiment {self.experiment_name} set up.')
        except Exception as err:
            print(err)
        
    def set_center_wavelength(self, center_wave_length: float): 
        """Set the center wavelength of the spectrometer.
        Args:
            center_wave_length (float): Center wavelength.
        """ 
        self.set_value(
            SpectrometerSettings.GratingCenterWavelength,
            center_wave_length)
        
    def set_grating(self, grating: int = 2):
        """Sets spectrometer grating.
        Args:
            grating (int, optional): 0 for 'Mirror', 1 for 'Density: 900g/mm, Blaze: 550 nm', 2 for 'Density: 300g/mm, Blaze: 750 nm'. Defaults to 2.
        """
        
        if grating == 0:
            self.set_value(
                SpectrometerSettings.Grating,
                '[Mirror,1200][0][0]'
            )
        elif grating == 0:
            self.set_value(
                SpectrometerSettings.Grating,
                '[550nm,900][1][0]'
            )
        elif grating == 2:
            self.set_value(
                SpectrometerSettings.Grating,
                '[750nm,300][2][0]'
            )

    def get_spectrometer_info(self):
        """Give center wavelength and grating currently used.
        """
        print(String.Format("{0} {1}","Center Wave Length:" ,
                    str(self.experiment.GetValue(
                        SpectrometerSettings.GratingCenterWavelength))))     
        
        print(String.Format("{0} {1}","Grating:" ,
                    str(self.experiment.GetValue(
                        SpectrometerSettings.Grating))))
       
    def set_value(self, setting, value):
        """Check for existence before setting and set provided value to setting. 
        Should be used mostly as an internal function as the setting parameter is not very readable.
        Args:
            setting (Settings): camera, spectrometer, experiment settings you want to set up.
            value (_type_): the new value to provide.
        """
       
        if self.experiment.Exists(setting):
            self.experiment.SetValue(setting, value)
            print(f'{setting} set to {value}.')
        else:
            print(f'Error in setting {setting} to value: {value}.')

    def set_exposure_time(self, time: float):
        """Function to set exposure time.

        Args:
            time (float): Exposure time (in ms).
        """
        self.set_value(CameraSettings.ShutterTimingExposureTime, float(time))

    def set_saved_filename(self, filename: str, increment: bool = False, add_date: bool = False, add_time: bool = False):    
        """Sets the name of the saved data file. (See the "Save Data file" tab on the GUI).
        Args:
            filename (_type_): File name.
            increment (bool, optional): Add an increment number for each new data file. Defaults to False.
            add_date (bool, optional): Add current date to the name. Defaults to False.
            add_time (bool, optional): Add current time to the name. Defaults to False.
        """
        self.set_value(
            ExperimentSettings.FileNameGenerationBaseFileName,
            Path.GetFileName(filename))
        
        # Option to Increment, set to false will not increment
        self.set_value(
            ExperimentSettings.FileNameGenerationAttachIncrement,
            increment)

        # Option to add date
        self.set_value(
            ExperimentSettings.FileNameGenerationAttachDate,
            add_date)

        # Option to add time
        self.set_value(
            ExperimentSettings.FileNameGenerationAttachTime,
            add_time)
    
    def get_full_sensor_size(self):
        """Gives the full sensor dimensions (number of pixels).
        Returns:
            int: X position of the left border.
            int: Y position of the up border.
            int: Width of the region.
            int: Heigth of the region.
            int: XBinning, binning of the x axis.
            int: YBinning, binning of the y axis. 
        """
        dimensions = self.experiment.FullSensorRegion
        return dimensions.X, dimensions.Y, dimensions.Width, dimensions.Height, dimensions.XBinning, dimensions.YBinning 
    
    def set_custom_ROI(self, x: int = None, y: int = None, width: int = None, height: int = None, xbinning: int = None, ybinning: int = None):
        """Set a custom region as ROI. 

        Args:
            x (int): position of the left border.
            y (int): position of the up border.
            width (int): width of the region.
            height (int): heigth of the region.
            xbinning (int, optional): binning of the x axis. Defaults to None, will be set equal to width.
            ybinning (int, optional): binning of the x axis. Defaults to None,  will be set equal to height.
        Returns:
        RegionOfInterest: the custum ROI set.
        """
        if xbinning is None:
            xbinning = width
        if ybinning  is None:
            ybinning = height

        regions = RegionOfInterest(x, y, width, height, xbinning, ybinning)
        self.Experiment.SetCustomRegions(regions)
        return regions
    
    def set_sensor_mode(self, mode: int = 4.0):
        """Set the sensor mode.
        Args:
            mode (int, optional): 1 for the entire area of the sensor, 2 rows binned , 3 only one row, 4 custom ROI. Defaults to 4.0.
        """
        self.set_value(CameraSettings.ReadoutControlRegionsOfInterestSelection, float(mode))

    def acquire(self, filename: str):
        """Acquire spectrum data. 
        Args:
            filename (str): data file name.
        """
        self.set_saved_filename(filename=filename)
        self.experiment.Acquire()

    def save_experiment(self, name: str = None):
        """Save experiment.

        Args:
            name (str, optional): experiment name. Defaults to None.
        """
        if name is not None:
            self.experiment_name = name
        self.experiment.SaveAs(self.experiment_name)

    def plot_spectrum(self, file):
        """Plot a data file.
        Args:
            file (str): plot
        """
        data = pd.read_csv(file)
        plt.plot(data[data.columns[2]][:-1], data.columns[5][:-1])
        plt.xlabel('Wavelength [nm]')
        plt.ylabel('Intensity (counts)')

    def disconnect(self):
        """This function enables to close properly the lightfield app running. Always use this function, otherwise it can break lightfield and make it crash later on.
        """
        subprocess.run(["taskkill", "/IM", "AddInProcess.exe"])
        print("Lightfield disconnected.")
