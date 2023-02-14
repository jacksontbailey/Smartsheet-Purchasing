import sys
import time
import pyinputplus as pyip

class InputQuestions:

    """This class contains methods to prompt the user with questions to retrieve data and options.

    Attributes:
        start_time (int): The time when the program started executing.
        option_selected (str): The option selected by the user from the main menu.
        new_sheet_name (str): The name of the new smartsheet.
        excel_sheet_name (str): The name of the excel sheet to import data from.
        smartsheet_name (str): The name of the smartsheet to import data into.
        end (bool): A flag to determine if the user wants to end the program.
    """

    start_time = None

    def __init__(self):
        """The constructor for the InputQuestions class."""
        self.option_selected = None
        self.new_sheet_name = None
        self.excel_sheet_name = None
        self.smartsheet_name = None
        self.end = False
        InputQuestions.instance = self
    

    @classmethod
    def start_timer(cls):
        """Start the timer to measure the runtime of the program."""
        cls.start_time = time.time()


    @classmethod
    def print_run_time(cls):
        """Print the runtime of the program."""
        if cls.start_time is not None:
            print("Script run time: {:.2f} seconds".format(time.time() - cls.start_time))


    @classmethod
    def end_script(cls):
        """End the program and print the runtime if the user has indicated they want to end the program."""
        if cls.instance.end:
            cls.print_run_time()
            sys.exit("User has ended the program...")


    def main_script_options(self):
        """Prompt the user with a menu of options to create a new smartsheet or import data into an existing one.

        Returns:
            str: The option selected by the user.
        """

        InputQuestions.start_timer()    
        
        self.option_selected = pyip.inputMenu(
            prompt = "\nDo you want to create a new SmartSheet or import data into an existing one?\n\n",
            choices = ['Create New SmartSheet', 'Import Data into Existing Sheet', 'End'],
            numbered = True,
            allowRegexes=['End']
        )

        if self.option_selected == 'End':
            self.end = True
            self.end_script()
        
        return self.option_selected
    

    def name_new_smartsheet(self):
        """Prompt the user to enter the name of the new smartsheet.

        Returns:
            str: The name of the new smartsheet.
        """

        if self.option_selected == 'Create New SmartSheet':
            self.new_sheet_name = pyip.inputStr(
                prompt = '\nEnter new sheet name (or "End" to exit): ', 
                allowRegexes=['End']
            )
            
            if self.new_sheet_name == "End":
                self.end = True
                self.end_script()

            return self.new_sheet_name
        
        else:
            print(f"Cannot call [name_new_smartsheet] method. The option_selected variable was '{self.option_selected}' and not 'Create New SmartSheet'.")
            self.end_script()
            return None


    def select_excel_sheet(self):
        """Prompt the user to enter the name of the excel sheet they want to import data from.

        Returns:
            str: The name of the excel sheet.
        """

        if self.option_selected == 'Import Data into Existing Sheet':
            self.excel_sheet_name = pyip.inputStr(
                prompt='\nEnter excel sheet to grab data from (or "End" to exit): ', 
                allowRegexes=['End']
            )

            if self.excel_sheet_name == "End":
                self.end = True
                self.end_script()
            
            print(f"You entered: {self.excel_sheet_name}")
            return self.excel_sheet_name
        
        else:
            print(f"Cannot call [select_excel_sheet] method. The option_selected variable was '{self.option_selected}' and not 'Import Data into Existing Sheet'.")
            self.end_script()
            return None


    def select_smartsheet_name(self):
        """Prompt the user to enter the name of the smartsheet they want to import data into.

        Returns:
            str: The name of the smartsheet.
        """

        if self.option_selected == 'Import Data into Existing Sheet':

            self.smartsheet_name = pyip.inputStr(prompt= '\nEnter name of the SmartSheet you want to upload excel data to (or "End" to exit): ', allowRegexes=['End'])
            if self.smartsheet_name == "End":
                self.end = True
                self.end_script()
            
            print(f"Name of Selected SmartSheet: {self.smartsheet_name}")
            return self.smartsheet_name
        
        else:
            print(f"Cannot call [select_smartsheet_name] method. The option_selected variable was '{self.option_selected}' and not 'Import Data into Existing Sheet'.")
            self.end_script()
            return None
