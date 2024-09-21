"""the aim of this script is to automate screen captures from Teledyne Lecroy Oscilloscopes running Maui
Created 2024-09-19 by Elijah Berger (@nosnowfall)
License: use as you wish, but please credit me for derivations, forks, etc.
Documentation: see the docs for tkinter, pyvisa, and Teledyne Lecroy's Remote/Automation manuals

Technical disclaimers:
You must first install pyvisa (from pip), and then the appropriate NI VISA drivers from National Instruments
This script relies upon configurations in the same directory, scopecaptureconfig.ini
It assumes a Windows x64 machine, but can be easily modified for x32 or Unix"""
import logging
import configparser
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from pathlib import Path
import pyvisa

def initial_config() -> tuple[configparser.ConfigParser, Path]:
    """load or create if not found settings for the script
    config settings are:
    background | background color for screen capture <BLACK/WHITE>
    imagepath | default save directory, does NOT check validity at runtime
    imagename | default filename, to be replaced with autogeneration
    instrumentaddr | for faster connections to the same machine"""
    logging.info('loading configuration files...')
    config = configparser.ConfigParser()
    configfilepath = Path(__file__).parent / 'scopecaptureconfig.ini'
    logging.debug(f'looking for: {configfilepath}')
    if not config.read(configfilepath): # returns false if the file is nonexistant or empty
        logging.debug('could not find scopecaptureconfig.ini; creating it now...')
        config['config'] = {'background': 'WHITE', 'imagepath': 'C:\\Users\\Public\\Pictures', 'imagename': 'screencapture.jpeg', 'instrumentaddr': 'USB0::TEMPLATE'}
        save_config(config, configfilepath)
    else:
        logging.debug('found scopecaptureconfig.ini...')
    for key in config['config']:
        logging.info(f'set {key}: {config['config'][key]}')
    return config, configfilepath

def save_config(config: configparser.ConfigParser, filepath: Path) -> None:
    """helper function so users can change configs later"""
    logging.info('saving updated configuration')
    config.write(open(filepath,'w'))
    return None

def change_config(config: configparser.ConfigParser, key: str, val: str) -> None:
    logging.debug(f'changing config of {key} to {val}')
    config['config'][key] = val
    return None

def main():
    cfg, cfgpath = initial_config() # tkinter is often used with global scope, but we are trying to avoid that

    # main window for tkinter
    root = Tk()
    root.title("Oscilloscope Screen Capture")
    main = ttk.Frame(root, padding="3 3 12 12")
    main.grid(column=0, row=0, sticky=(N, S, E, W))
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    # connection settings
    visastatustext = StringVar() # displays status of visa resourcemanager
    visastatustext.set('VISA: DOWN')
    visastatus = BooleanVar() # todo: use for button enable/disable
    visastatus.set(False)

    connstatustext = StringVar() # displays status of instrument connection
    connstatustext.set('LINK: DOWN')
    connstatus = BooleanVar() # todo: use for button enable/disable
    connstatus.set(False)
    
    oscope = pyvisa.Resource # i think this hack allows callbacks to edit the variable, circumventing stupid macro chains

    def loadvisa() -> tuple[pyvisa.ResourceManager, list]:
        """should not fail macro to create a Resource Manager, called on startup"""
        logging.debug('loading VISA resource manager')
        try:
            rm = pyvisa.ResourceManager("C:\\Windows\\System32\\visa64.dll") # you should specify the dll, see pyvisa docs for more info
            resources = rm.list_resources()
        except Exception as e: # prevent crashes, just let us try again later
            logging.warning(f'Visa resource manager crashed: {repr(e)}')
            visastatus.set(False)
            visastatustext.set('VISA: DOWN')
            rm = None
            resources = []
        else:
            visastatus.set(True)
            visastatustext.set("VISA: UP")
            logging.debug('VISA RM loaded successfully')
        finally:
            return rm, resources
    
    def tryconnect() -> None:
        """try to open visa comms with instrument, fails quite often for I think backend bug reasons"""
        logging.debug(f'trying connection to {cfg['config']['instrumentaddr']}')
        for addr in rm.list_opened_resources(): # prevent duplicate open resources
            addr.close()
            connstatustext.set('LINK: DOWN')
        try:
            oscope = rm.open_resource(cfg['config']['instrumentaddr']) # pull from cfg for callback ability
        except Exception as e:
            logging.warning(f'Instrument connection failed: {repr(e)}')
            oscope = pyvisa.Resource # same hack again, fingers crossed
            connstatus.set(False)
            connstatustext.set('LINK: DOWN')
        else:
            connstatus.set(True)
            connstatustext.set('LINK: UP')
        finally:
            return None
    
    def settarget() -> None:
        """callback macro to change instrument address in cfg
        could probably be a lambda at button definition"""
        change_config(cfg, 'instrumentaddr', target.get())

    rm, resources = loadvisa() # i couldnt find a way to initialize these after the window is created

    visaframe = ttk.Labelframe(main, text='NI Visa Status')
    visaframe.grid(column=0,row=0)
    visastatuslabel = ttk.Label(visaframe, textvariable=visastatustext)
    visastatuslabel.grid(column=0,row=0)
    visabutton = ttk.Button(visaframe, text='Try VISA', command=loadvisa) # this won't work as is, because we need to return the RM and resources
    visabutton.grid(column=0,row=1)

    connframe = ttk.Labelframe(main, text='Instrument Status')
    connframe.grid(column=1, row=0)
    connstatuslabel = ttk.Label(connframe, textvariable=connstatustext)
    connstatuslabel.grid(column=0,row=0)
    connbutton = ttk.Button(connframe, text='Connect Instrument', command=tryconnect)
    connbutton.grid(column=1, row=0)
    
    target = StringVar()
    target.set(cfg['config']['instrumentaddr'])
    connentry = ttk.OptionMenu(connframe, variable=target, *resources, command=settarget) # autopopulates from resources i think
    connentry.grid(column=0,row=1,columnspan=2)

    # background color, radiobutton choice and saves to cfg
    bckgframe = ttk.LabelFrame(main, text='Background color')
    bckgframe.grid(column=2, row=3, sticky=EW)
    bckg = StringVar()
    bckg.set(cfg['config']['background'])
    black = ttk.Radiobutton(bckgframe, text='Black', variable=bckg, value='BLACK', command=lambda: change_config(cfg, 'background', 'BLACK'))
    white = ttk.Radiobutton(bckgframe, text='White', variable=bckg, value='WHITE', command=lambda: change_config(cfg, 'background', 'WHITE'))
    black.pack(side=LEFT)
    white.pack(side=RIGHT)

    # image save directory, using file picker dialog box
    def choose_savedir() -> None:
        newdir = filedialog.askdirectory()
        imagepath.set(newdir)
        change_config(cfg, 'imagepath', newdir)
    ttk.Label(main, text='Save to:').grid(column=0, row=1, sticky=E)

    imagepath = StringVar()
    imagepath.set(cfg['config']['imagepath'])
    imagepath_entry = ttk.Label(main, textvariable=imagepath, background='#d3d3d3') # need to update functionality - currently doesnt change cfg
    imagepath_entry.grid(column=1, row=1, sticky=EW)
    browsebutton = ttk.Button(main, text='Browse', command=choose_savedir)
    browsebutton.grid(column=2, row=1, sticky=W)

    # screencap
    def prtscrmacro() -> None:
        hcsucmd = f"HCSU DEV, JPEG, BCKG, {cfg['config']['background']}, AREA, GRIDAREAONLY, PORT, NET" # setup screen capture params
        oscope.write(hcsucmd)
        oscope.write('SCDP') # ask scope to make a screen capture, which according to our previous command will send over VISA
        capture = oscope.read_raw() # read image file from instrument
        savemacro(capture)

    def savemacro(photo: bytes) -> None:
        savedir = Path(cfg['config']['imagepath'])
        f = open(savedir / imagename.get(), 'wb+') # wb+ for binary writing. todo: autogenerate names
        f.write(photo)
        f.close()

    ttk.Label(main, text='Save as:').grid(column=0, row=3, sticky=E)
    imagename = StringVar()
    imagename.set(cfg['config']['imagename'])
    imagename_entry  = ttk.Entry(main, textvariable=imagename, background='#d3d3d3')
    imagename_entry.grid(column=1, row=3, sticky=EW)
    screengrabber = ttk.Button(main, text='Print Screen', command=prtscrmacro)
    screengrabber.grid(column=1, row=4, sticky=EW)

    for child in main.winfo_children():
        child.grid_configure(padx=5, pady=5)

    root.mainloop()

    # cleanup
    for addr in rm.list_opened_resources():
        addr.close()
    rm.close()
    save_config(cfg, cfgpath)
    return None

if __name__ == "__main__":
    logging.basicConfig(
        level = logging.DEBUG,
        format = '%(asctime)s.%(msecs)03d %(levelname)s: %(message)s',
        datefmt = '%Y-%m-%d %H:%M:%S',
        encoding = 'utf-8'
    )
    logging.debug('Logging initialized')
    try:
        main()
    except Exception as e:
        logging.warning(f'Execution failed with error: {repr(e)}')
        exit()