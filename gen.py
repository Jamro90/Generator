    # -----|MODULES & LIBERES|----- #
import sys                                                                                                  # module making sys calls
sys.path.append("/home/chris/python_scripts/terminal_installer/GUI/apllicant/generator.ver1.1alpha/lib")    # importing a modules in lib dir
from date_lib import date_read, date_show, date_change
import docx_lib
from gui_lib import window_start
import logic_n_alert
from tool_lib import toolbar_settings



print("Hello, World!")
print(date_read())
window_start()
