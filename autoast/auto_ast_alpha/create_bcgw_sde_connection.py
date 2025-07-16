import arcpy
import os
import tempfile
from tkinter import *
import arcpy
import keyring
import shutil
import logging
import sys

import config


class BCGWLoginWindow:
    DEFAULT_WINDOW_SIZE = "350x150"
    BUTTON_COLOR = "#fff"  # white


    def __init__(self, user, msg):
        # Initialize the main Tkinter window
        self.main = Tk()
        self.user = user
        self.msg = msg
        self.login_successful = False  # Flag to track login status
        self.main.title('Enter Credentials to BCGW')
        self.main.geometry(self.DEFAULT_WINDOW_SIZE)

        # Configure rows and columns for grid layout
        for i in range(11):
            self.main.rowconfigure(i, weight=1)
            self.main.columnconfigure(i, weight=1)

        # Username label and entry
        self.label_text = StringVar(value="Username: ")
        self.label_dir = Label(self.main, textvariable=self.label_text, height=1)
        self.label_dir.grid(row=2, column=4)

        self.username_box = Entry(self.main)
        self.username_box.insert(0, self.user)
        self.username_box.bind("<FocusIn>", self.clear_widget)
        self.username_box.bind('<FocusOut>', self.repopulate_defaults)
        self.username_box.grid(row=2, column=5, sticky='NS')

        # Password label and entry
        self.pw_text = StringVar(value="Password: ")
        self.label_pw = Label(self.main, textvariable=self.pw_text, height=1)
        self.label_pw.grid(row=4, column=4)

        self.password_box = Entry(self.main, show='*')
        self.password_box.insert(0, '')
        self.password_box.bind("<FocusIn>", self.clear_widget)
        self.password_box.bind('<FocusOut>', self.repopulate_defaults)
        self.password_box.bind('<Return>', lambda event: self.login())
        self.password_box.grid(row=4, column=5, sticky='NS')

        # Message label
        self.msg_text = StringVar(value=msg)
        self.msg_dir = Label(self.main, textvariable=self.msg_text, height=2, wraplength=250, justify='left')
        self.msg_dir.grid(row=5, column=5, columnspan=3)

        # Login button
        self.login_btn = Button(self.main, text='Create BCGW Connection', command=self.login, bg=self.BUTTON_COLOR)
        self.login_btn.bind('<Return>', lambda event: self.login())
        self.login_btn.grid(row=7, column=5, sticky='E')

        # Quit button
        self.quit_btn = Button(self.main, text='Quit', command=self.quit_application, bg=self.BUTTON_COLOR)
        self.quit_btn.grid(row=7, column=6, sticky='W')

        # Start the Tkinter event loop
        self.main.mainloop()

    def clear_widget(self, event):
        # Clears the username or password entry box when focused
        widget = event.widget
        if widget == self.username_box and widget.get() == '':
            widget.delete(0, 'end')
        elif widget == self.password_box and widget.get() == '':
            widget.delete(0, 'end')

    def repopulate_defaults(self, event):
        # Repopulates default text if entry box is empty and focus is lost
        widget = event.widget
        if widget != self.main.focus_get() and widget.get() == '':
            widget.insert(0, '')

    def login(self):
        # Logs in and closes the window
        self.username = self.username_box.get()
        self.password = self.password_box.get()
        self.login_successful = True
        self.main.destroy()

    def quit_application(self):
        # Quits the application
        self.login_successful = False
        self.main.destroy()

    def entered_creds(self):
        # Returns entered username and password
        return self.username, self.password


class ManageCredentials:
    def __init__(self, key, proj_location=None):
        # Set up stored key and initialize username and password
        os.environ["ORACLE_CONN"] = key
        self.stored_key = os.environ["ORACLE_CONN"]
        self.proj_location = proj_location
        self.username = None
        self.password = None
        self.max_tries = 3
        self.try_count = 0

        # Create a temporary directory to store SDE connection file
        self.out_folder_path = tempfile.mkdtemp()

        # Configure logging
        self.setup_logging()


    def setup_logging(self):
        log = '!ManageCredentials'
        # Create a logger
        self.logger = logging.getLogger(f'{log}')
        self.logger.setLevel(logging.INFO)

        # Create a logging format
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M.%S')

        if self.proj_location:
            # Use the provided location for the logger file
            log_file = os.path.join(self.proj_location, f'{log}.log')

            # Check if the logger file already exists
            if os.path.exists(log_file):
                # If the logger file exists, add a file handler to append to it
                file_handler = logging.FileHandler(log_file, mode='a')
            else:
                # If the logger file doesn't exist, create a new file handler
                file_handler = logging.FileHandler(log_file, mode='w')
        else:
            # Use the provided location for the logger file
            log_file = os.path.join(self.out_folder_path, f'{log}.log')

            # Use current directory for the logger file if location is not provided
            file_handler = logging.FileHandler(f'{log}.log', mode='w')

        # Set the level and formatter for the file handler
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(formatter)

        # Add the file handler to the logger
        self.logger.addHandler(file_handler)


    def create_sde(self, user, pw):
        """
        Function to create SDE connection and test it
        """
        def test_sde_connection():
            """
            Function to test the SDE connection
            """
            try:
                arcpy.env.workspace = os.environ["SDE_FILE_PATH"]
                # Check if any datasets exist
                return bool(arcpy.ListDatasets())
            except (IndexError, arcpy.ExecuteError):
                self.logger.error("Failed SDE test using credentials")
                return False

        #set oracle connection variables
        out_name = "bcgw.sde"
        database_platform = config.CONNPLATFORM
        instance = config.CONNINSTANCE
        # instance = "bcgw-i.bcgov/idwdlvr1.bcgov"   #Test oracle environment.
        account_authentication = "DATABASE_AUTH"
        username = str(user)
        password = str(pw)
        save_user_pass = "SAVE_USERNAME"

        # Create SDE Connection
        try:
            arcpy.CreateDatabaseConnection_management(self.out_folder_path, out_name, database_platform, instance,
                                                      account_authentication, username, password, save_user_pass)
            sde_full = os.path.join(self.out_folder_path, out_name)
            os.environ["SDE_FILE_PATH"]=sde_full
            self.logger.info(f"SDE created in: {sde_full}")

            # Test the connection
            return test_sde_connection()
        
        except Exception as e:
            # Log the error or handle it appropriately
            self.logger.error(f"Error creating SDE connection: {e}")
            # print(f"Error creating SDE connection: {e}")
            return False

    def check_credentials(self):
        """
        Function to check stored credentials and prompt for new ones if needed
        """
        # Try to use stored credentials first
        creds = keyring.get_credential(self.stored_key, "")
        if creds:
            self.username = creds.username
            self.password = creds.password
            if self.create_sde(self.username, self.password):
                self.logger.info(f"Existing credentials verified. SDE connection established successfully for '{self.stored_key}'.")
                return True

        # Prompt the user to enter credentials if no stored credentials or
        #current credentials are not valid
        return self.get_and_check_credentials()

    def get_and_check_credentials(self, remaining_tries=3):
        """
        Function to prompt the user for new credentials and check them
        """
        user = os.getenv('USERNAME') or ''
        message = f"{self.stored_key} credentials are unavailable or invalid. Enter your credentials (Remaining Attempts: {remaining_tries})"

        # Check if the user has exceeded the maximum number of tries
        if self.try_count >= self.max_tries:
            self.logger.error("User has exceeded the maximum number of tries.")
            # print("You have exceeded the maximum number of tries.")
            return False

        # Instantiate BCGWLoginWindow class to get new credentials
        creds = BCGWLoginWindow(user, message)
        if creds.login_successful:
            # Get new credentials from user
            self.username, self.password = creds.entered_creds()
            # Increment the try count
            self.try_count += 1
        else:
            self.logger.warning("Application aborted at login window")
            return False

        # Check if the new credentials are valid
        if self.create_sde(self.username, self.password):
            # Store the credentials into keyring if valid
            keyring.set_password(self.stored_key, self.username, self.password)
            self.logger.info(f"Entered credentials verified. SDE connection established successfully for '{self.stored_key}'.")
            return True
        else:
            # Calculate remaining tries
            remaining_tries = self.max_tries - self.try_count
            if remaining_tries == 0:
                self.logger.error("User has exceeded the maximum number of tries.")
                try:
                    shutil.rmtree(os.path.dirname(os.path.abspath(os.getenv("SDE_FILE_PATH"))))
                except Exception as e:
                    self.logger.error("Could not remove temporary directory.")
                finally:
                    return False

            # If not valid, recursively call get_and_check_credentials to prompt for credentials again
            self.logger.warning(f"Invalid credentials. {remaining_tries} tries remaining.")
            return self.get_and_check_credentials(remaining_tries)

if __name__ == '__main__':
    # Create an instance of ManageCredentials
    key_name = config.CONNNAME
    out = r"\\spatialfiles.bcgov\work\srm\wml\Workarea\arcproj\!Williams_Lake_Toolbox_Development\automated_status_ARCPRO\steve\logging"

    try:
        oracleCreds = ManageCredentials(key_name, out)
        if not oracleCreds.check_credentials():
            arcpy.AddError("BCGW credentials could not be established.")
            print("BCGW credentials could not be established.")
            sys.exit()
    except Exception as e:
        arcpy.AddError("Failure occurred when establishing BCGW connection.")
        print(f"Failure occurred when establishing BCGW connection - {e}")
        sys.exit()
    print(os.getenv("SDE_FILE_PATH"))
    print(keyring.get_credential(key_name, ""))

    #cleanup credentials and sde connections
    keyring.delete_password(key_name, "sdrichar")
    shutil.rmtree(os.path.dirname(os.path.abspath(os.environ["SDE_FILE_PATH"])))
