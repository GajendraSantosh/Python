'''
Special Thanks to https://stackoverflow.com

This program will open the Googel Chrome and download files to the default downloaded folder
and it will rename all downloaded file.

Note: -> Make sure to download the nessesary compatibility Google Chrome Web drivers,
          and paste the chromedriver.exe in the root directory of this .py file.
          http://chromedriver.chromium.org/downloads
          https://chromedriver.storage.googleapis.com/index.html
      -> Disable 'ask where to save each file before downloading' in the advance settings
'''

# Imports Modules
from selenium import webdriver  # Chrome
from sys import exit            # Exit code
import datetime                 # Date
import time                     # Sleep
import os                       # Rename files, Delete files, find latest file
import platform                 # Check the OS
import glob                     # Find latest file in path
import win32gui                 # Native OS dialog/Window
import re                       # reg ex


'''Finds the os and set the Download folder according to os'''
if platform.release()=='XP':
    dPath = "C:\\Documents and Settings\\"+os.environ.get('USERNAME')+"\\My Documents\\Downloads\\"
else:
    dPath = "C:\\Users\\"+os.environ.get('USERNAME')+"\\Downloads\\"

    
'''Wait until download to be complete'''
def download_wait(path_to_downloads):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < 20:
        time.sleep(1)
        dl_wait = False
        for fname in os.listdir(path_to_downloads):
            if fname.endswith('.crdownload'):
                dl_wait = True
        seconds += 1
    return seconds


'''Finds 'Save as' dialog box while downloading the file'''
class WindowFinder:
    """Class to find and make focus on a particular Native OS dialog/Window """
    def __init__ (self):
        self._handle = None

    def find_window(self, class_name, window_name = None):
        """Pass a window class name & window name directly if known to get the window """
        self._handle = win32gui.FindWindow(class_name, window_name)

    def _window_enum_callback(self, hwnd, wildcard):
        """Call back func which checks each open window and matches the name of window using reg ex"""
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) != None:
            self._handle = hwnd

    def find_window_wildcard(self, wildcard):
        """ This function takes a string as input and calls EnumWindows to enumerate through all open windows """
        self._handle = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)
        return self._handle

    def set_foreground(self):
        """Get the focus on the desired open window"""
        win32gui.SetForegroundWindow(self._handle)


'''Finds the variable contain Number or not'''
def is_number(s):
    try:
        float(s)
        return True
    except:
        return False

    
browser = webdriver.Chrome() # Open Crome Application
browser.get('http://prod01/projects/loginform.aspx') # goto Projects Export

browser.find_element_by_id('txtusername').send_keys('UserID') # textbox User ID
browser.find_element_by_id('txtpassword').send_keys('UserPassword') # textbox Password
browser.execute_script("""javascript:WebForm_DoPostBackWithOptions(new WebForm_PostBackOptions("LoginButton", "", true, "Login1", "", false, true))""")
browser.get('http://prod01/projects/ProjectsExport.aspx') # goto Projects Export

#Download DIS Export ---
while True:
    if browser.current_url=='http://prod01/projects/ProjectsExport.aspx':
        browser.find_element_by_id('ctl00_ContentPlaceHolder1_rbtnDatainputStage').click() # optionbox Data input Stage
        browser.execute_script("""javascript:WebForm_DoPostBackWithOptions(new WebForm_PostBackOptions("ctl00$ContentPlaceHolder1$lnkExportToExcel1", "", true, "", "", false, true))""") # Click Export to Excel
        time.sleep(0.5)
        if is_number(WindowFinder().find_window_wildcard(".*Save As.*")):
            print("Error : Ask where to save each file before downloading = True")
            exit(0)
        else:
            download_wait(dPath)
            DIS = max(glob.glob(dPath+"*"), key=os.path.getctime)
            browser.get('http://prod01/projects/ProjectsExport.aspx') # goto Projects Export
        break
        
#Download Aproved Export ---
while True:
    if browser.current_url=='http://prod01/projects/ProjectsExport.aspx':    
        browser.find_element_by_id('ctl00_ContentPlaceHolder1_rptrProjectFields_ctl17_chkProjectFields').click() # checkbox QC/PublishedBy
        browser.find_element_by_id('ctl00_ContentPlaceHolder1_rbtnAproved').click() # ob Aproved
        browser.execute_script("""javascript:WebForm_DoPostBackWithOptions(new WebForm_PostBackOptions("ctl00$ContentPlaceHolder1$lnkExportToExcel1", "", true, "", "", false, true))""") # Click Export to Excel
        time.sleep(0.5)
        if is_number(WindowFinder().find_window_wildcard(".*Save As.*")):
            print("Error : Ask where to save each file before downloading = True")
            exit(0)
        else:
            download_wait(dPath)
            Apr = max(glob.glob(dPath+"*"), key=os.path.getctime)
            browser.get('http://prod01/projects/ProjectsExport.aspx') # goto Projects Export
        break

#Download Published Export ---
while True:
    if browser.current_url=='http://prod01/projects/ProjectsExport.aspx':
        browser.find_element_by_id('ctl00_ContentPlaceHolder1_rptrProjectFields_ctl17_chkProjectFields').click() # checkbox QC/PublishedBy
        browser.find_element_by_id('ctl00_ContentPlaceHolder1_rbtnPublished').click() # optionbox Published
        browser.find_element_by_id('ctl00_ContentPlaceHolder1_txtFromDate').send_keys(datetime.date.today().strftime ("%d-%b-%Y")) # textbox From Date
        browser.execute_script("""javascript:WebForm_DoPostBackWithOptions(new WebForm_PostBackOptions("ctl00$ContentPlaceHolder1$lnkExportToExcel1", "", true, "", "", false, true))""") # Click Export to Excel
        time.sleep(0.5)
        if is_number(WindowFinder().find_window_wildcard(".*Save As.*")):
            print("Error : Ask where to save each file before downloading = True")
            exit(0)
        else: 
            download_wait(dPath)
            Pub = max(glob.glob(dPath+"*"), key=os.path.getctime)
            break
        

'''Rename the downloaded file'''
try:
    os.rename(DIS,dPath+"DIS_Export.xlsx")
except:
    os.remove(dPath+"DIS_Export.xlsx")
    os.rename(DIS,dPath+"DIS_Export.xlsx")

try:
    os.rename(Apr,dPath+"Aproved_Export.xlsx")
except:
    os.remove(dPath+"Aproved_Export.xlsx")
    os.rename(Apr,dPath+"Aproved_Export.xlsx")

try:
    os.rename(Pub,dPath+"Published_Export.xlsx")
except:
    os.remove(dPath+"Published_Export.xlsx")
    os.rename(Pub,dPath+"Published_Export.xlsx")    
            
browser.execute_script("""javascript:__doPostBack('ctl00$hlnkLogout','')""") # Click Logout
browser.quit() # Close Crome Application
