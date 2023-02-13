import sys
import time
import pyinputplus as pyip

class InputQuestions:
    start_time = None

    def __init__(self):
        self.option_selected = None
        self.new_sheet_name = None
        self.excel_sheet_name = None
        self.smartsheet_name = None
        self.end = False
        InputQuestions.instance = self
    

    @classmethod
    def start_timer(cls):
        cls.start_time = time.time()


    @classmethod
    def print_run_time(cls):
        if cls.start_time is not None:
            print("Script run time: {:.2f} seconds".format(time.time() - cls.start_time))


    @classmethod
    def end_script(cls):
        if cls.instance.end:
            cls.print_run_time()
            sys.exit("User has ended the program...")


    def main_script_options(self):
        InputQuestions.start_timer()    
        
        self.option_selected = pyip.inputMenu(
            prompt = "Do you want to create a new SmartSheet or import data into an existing one?\n",
            choices = ['Create New SmartSheet', 'Import Data into Existing Sheet', 'End'],
            numbered = True,
            allowRegexes=['End']
        )

        if self.option_selected == 'End':
            self.end = True
            self.end_script()
        
        return self.option_selected
    

    def name_new_smartsheet(self):
        if self.option_selected == 'Create New SmartSheet':
            self.new_sheet_name = pyip.inputStr(
                prompt = 'Enter new sheet name (or "End" to exit): ', 
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
        if self.option_selected == 'Import Data into Existing Sheet':
            self.excel_sheet_name = pyip.inputStr(
                prompt='Enter excel sheet to grab data from (or "End" to exit): ', 
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
        if self.option_selected == 'Import Data into Existing Sheet':

            self.smartsheet_name = pyip.inputStr(prompt= 'Enter name of the SmartSheet you want to upload excel data to (or "End" to exit): ', allowRegexes=['End'])
            if self.smartsheet_name == "End":
                self.end = True
                self.end_script()
            
            print(f"Name of Selected SmartSheet: {self.smartsheet_name}")
            return self.smartsheet_name
        
        else:
            print(f"Cannot call [select_smartsheet_name] method. The option_selected variable was '{self.option_selected}' and not 'Import Data into Existing Sheet'.")
            self.end_script()
            return None
