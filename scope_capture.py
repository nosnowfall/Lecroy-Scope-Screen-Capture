"""the aim of this script is to automate screen captures from Teledyne Lecroy Oscilloscopes running Maui
Created 2024-09-19 (@nosnowfall)
Modified 2025-04-10 (@nosnowfall)
License: use as you wish, but please credit me for derivations, forks, etc.
Documentation: see the docs for tkinter, pyvisa, and Teledyne Lecroy's Remote/Automation manuals

Technical disclaimers:
You must first install pyvisa (from pip), and then the appropriate NI VISA drivers from National Instruments
This script relies upon configurations in the same directory, scopecaptureconfig.ini
It assumes a Windows x64 machine, but can be easily modified for x32 or Unix
USB based connections are wayyy more stable than TCP/IP"""

import logging
import configparser
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from pathlib import Path
import pyvisa
from datetime import datetime

class Oscope():
    """ Wrapping the oscilloscope connection in a class solves some tkinter issues """
    def __init__(self) -> None:
        self.visaobj :pyvisa.resources.Resource = 0 # dummy val

    def update_oscope(self, addr: pyvisa.resources.Resource) -> None:
        """ change the target VISA address """
        self.visaobj = addr
        logger.info(f'changed addr to: {addr}')

class RM_Manager():
    def __init__(self) -> None:
        self.resman = 0 # dummy val
        self.reslist :list = [] # fill after opening connection
        self.start_resman() # true init function

    def start_resman(self) -> None:
        """ Starts the PyVISA resource manager instance and grabs a list of available VISA resources """
        logger.info('starting VISA manager')
        self.resman :pyvisa.ResourceManager = pyvisa.ResourceManager("C:\\Windows\\System32\\visa64.dll")
        self.reslist = self.resman.list_resources() # this will find any saved resources in NI Max unfortunately
        logger.debug(f'RM found VISA resources: {self.reslist}')

    def reset_resman(self) -> None:
        """ Reset the resource manager in case of error """
        logger.info('resetting VISA manager')
        self.reslist.clear()
        self.close_links()
        self.close_resman()
        self.start_resman()
    
    def close_links(self) -> None:
        """ Close any active VISA resource """
        logger.info('clearing VISA links')
        for object in self.resman.list_opened_resources():
            object.close()

    def close_resman(self) -> None:
        """ Ensure the PyVISA backend closes """
        logger.info('closing resource manager')
        self.resman.close()

def initial_config() -> tuple[configparser.ConfigParser, Path]:
    """load or create if not found settings for the script
    config settings are:
    background | background color for screen capture <BLACK/WHITE>
    imagepath | default save directory, does NOT check validity at runtime
    imagename | default filename, to be replaced with autogeneration
    instrumentaddr | for faster connections to the same machine"""
    logger.info('loading configuration files...')
    config :configparser.ConfigParser = configparser.ConfigParser()
    configfilepath :Path = Path(__file__).parent / 'scopecaptureconfig.ini'
    logger.debug(f'looking for: {configfilepath}')
    if not config.read(configfilepath): # returns false if the file is nonexistant or empty
        logger.warning('could not find scopecaptureconfig.ini; creating it now...')
        config['config'] = {'background': 'WHITE', 'imagepath': 'C:\\Users\\Public\\Pictures'}
        save_config(config, configfilepath)
    else:
        logger.debug('found scopecaptureconfig.ini...')
    for key in config['config']:
        logger.info(f'set {key}: {config['config'][key]}')
    return config, configfilepath

def save_config(config: configparser.ConfigParser, filepath: Path) -> None:
    """ helper function for saving user-changed config """
    logger.info('saving updated configuration')
    config.write(open(filepath,'w'))
    return None

def change_config(config: configparser.ConfigParser, key: str, val: str) -> None:
    """ helper function for changing config without saving """
    logger.debug(f'changing config of {key} to {val}')
    config['config'][key] = val
    return None

def main() -> None:
    cfg, cfgpath = initial_config() # tkinter is often used with global scope, but we are trying to avoid that

    # main window for tkinter, todo make resizeable
    root = Tk()
    root.title("Oscilloscope Screen Capture")
    main = ttk.Frame(root, padding="3 3 12 12")
    main.grid(column=0, row=0, sticky=(N, S, E, W))
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    # connection settings
    connstatustext = StringVar() # displays status of instrument connection
    connstatustext.set('LINK: DOWN')
    connstatus = BooleanVar() # todo: use for button enable/disable
    connstatus.set(False)
    
    resman = RM_Manager() # object wrapper hack part two

    oscope = Oscope() # trying a new object wrapper hack to fool tkinter
    
    def tryconnect() -> None:
        """ try to open visa comms with instrument, fails quite often for I think backend bug reasons """
        logger.debug(f'trying connection to {target.get()}')
        resman.close_links() # make sure we don't duplicate links
        connstatustext.set('LINK: DOWN')
        try:
            oscope.update_oscope(resman.resman.open_resource(target.get())) # pull from cfg for callback ability
        except Exception as e:
            logger.warning(f'Instrument connection failed: {repr(e)}')
            oscope.update_oscope(0)
            connstatus.set(False)
            connstatustext.set('LINK: DOWN')
        else:
            connstatus.set(True)
            connstatustext.set('LINK: UP')
        finally:
            return None

    def retryvisa() -> None:
        """ reset the Resource Manager and refresh menu of VISA devices """
        resman.reset_resman()
        connentry['values'] = resman.reslist
    # retry button
    visabutton = ttk.Button(main, text='Retry VISA', command=retryvisa) # this won't work as is, because we need to return the RM and resources - 2025-04-10 maybe not an issue?
    visabutton.grid(column=0,row=0)
    # status and connect box
    connframe = ttk.Labelframe(main, text='Instrument Status')
    connframe.grid(column=1, row=0, columnspan=2, sticky=EW)
    connstatuslabel = ttk.Label(connframe, textvariable=connstatustext)
    connstatuslabel.grid(column=0,row=0)
    connbutton = ttk.Button(connframe, text='Connect Instrument', command=tryconnect)
    connbutton.grid(column=1, row=0)
    # let user choose target
    target = StringVar()
    connentry = ttk.Combobox(connframe, width=45, textvariable=target) # can we make size variable?
    connentry['values'] = resman.reslist # autopopulates from resources i think
    connentry.grid(column=0,row=1,columnspan=2)

    # background color, radiobutton choice and saves to cfg
    bckgframe = ttk.LabelFrame(main, text='Background color')
    bckgframe.grid(column=0, row=4, sticky=EW)
    bckg = StringVar()
    bckg.set(cfg['config']['background'])
    black = ttk.Radiobutton(bckgframe, text='Black', variable=bckg, value='BLACK', command=lambda: change_config(cfg, 'background', 'BLACK'))
    white = ttk.Radiobutton(bckgframe, text='White', variable=bckg, value='WHITE', command=lambda: change_config(cfg, 'background', 'WHITE'))
    black.pack(side=LEFT)
    white.pack(side=RIGHT)

    # image save directory, using file picker dialog box
    def choose_savedir() -> None:
        """ open file dialog to pick a savefile directory - remembers from last use """
        newdir = filedialog.askdirectory() # there is a bug with root not closing filedialogs but only with many many opens.
        # ^^ beware of unavailable network drives, Windows 11 crashes when some programs try to read them
        imagepath.set(newdir)
        change_config(cfg, 'imagepath', newdir)
    ttk.Label(main, text='Save to:').grid(column=0, row=1, sticky=E)

    imagepath = StringVar()
    imagepath.set(cfg['config']['imagepath'])
    imagepath_entry = ttk.Label(main, textvariable=imagepath, background='#d3d3d3') # need to update functionality - currently doesnt change cfg - 2025-04-10 maybe it does now
    imagepath_entry.grid(column=1, row=1, sticky=EW)
    browsebutton = ttk.Button(main, text='Browse', command=choose_savedir)
    browsebutton.grid(column=2, row=1, sticky=W)

    # screencap
    def prtscrmacro() -> None:
        """ sends VISA command to print screen, reads response, calls savemacro()
        
        If a new kind of scope needs to be added, this is the place to start tweaking """
        hcsucmd = f"HCSU DEV, JPEG, BCKG, {cfg['config']['background']}, AREA, GRIDAREAONLY, PORT, NET" # setup screen capture params
        oscope.visaobj.write(hcsucmd)
        oscope.visaobj.write('SCDP') # ask scope to make a screen capture, which according to our previous command will send over VISA
        capture = oscope.visaobj.read_raw() # read image file from instrument
        savemacro(capture)

    def savemacro(photo: bytes) -> None:
        """ saves data to file as jpeg """
        if autonaming.get():
            filename = datetime.now().strftime('scopecapture_%Y%m%d_%H%M%S%f')[:-3] + '.jpeg'
        else:
            filename = imagename.get()
        savedir = Path(cfg['config']['imagepath'])
        logger.info(f'writing to file now')
        f = open(savedir / filename, 'wb+') # wb+ for binary writing. todo: auto increment names
        f.write(photo)
        f.close()

    # choosing filename
    def namebuttonfunc():
        """ callback to enable or disable automatic file naming """
        if autonaming.get():
            imagename_entry.config(state='disabled')
        else:
            imagename_entry.config(state='enabled')

    nameframe = ttk.LabelFrame(main, text='Save as:')
    nameframe.grid(column=1, row=3, columnspan=2, sticky=EW)
    imagename = StringVar()
    autonaming = BooleanVar()
    autonaming.set(True)
    imagename.set(cfg['config']['imagename'])

    autonameframe = ttk.LabelFrame(main, text='Naming')
    autonameframe.grid(column=0, row=3, sticky=EW)
    auto = ttk.Radiobutton(autonameframe, text='Auto', variable=autonaming, value=True, command=namebuttonfunc)
    manual = ttk.Radiobutton(autonameframe, text='Manual', variable=autonaming, value=False, command=namebuttonfunc)
    auto.pack(side=LEFT)
    manual.pack(side=RIGHT)

    imagename_entry  = ttk.Entry(nameframe, textvariable=imagename, background='#d3d3d3', state='disabled')
    imagename_entry.pack()
    
    screengrabber = ttk.Button(main, text='Print Screen', command=prtscrmacro)
    screengrabber.grid(column=1, row=4, sticky=EW)

    for child in main.winfo_children():
        child.grid_configure(padx=5, pady=5)

    root.mainloop() #  party starts now

    # cleanup
    resman.close_links()
    resman.close_resman()

    cfg['config']['instrumentaddr'] = 'No instrument found' # prevent false information on next start
    save_config(cfg, cfgpath)
    return None

if __name__ == "__main__":
    logging.basicConfig(
        level = logging.INFO,
        format = '%(asctime)s.%(msecs)03d %(levelname)s: %(message)s',
        datefmt = '%Y-%m-%d %H:%M:%S',
        encoding = 'utf-8'
    )
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    logging.debug('Logging initialized')
    try:
        main()
    except Exception as e:
        logging.warning(f'Execution failed with error: {repr(e)}')
        exit()
