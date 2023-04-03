import logging
import os

class ErrorLogger:
    # Initialize the class with the folder name and the file name
    def __init__(self, folder_name, file_name):
        # Get the absolute path of the folder
        self.folder_path = os.path.abspath(folder_name)
        # Create the file name by joining the folder path and the file name
        self.file_name = os.path.join(self.folder_path, file_name)
        # Configure the logger to write to the file name
        logging.basicConfig(filename=self.file_name, filemode='a', format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S', level=logging.ERROR)
        # Create a logger object
        self.logger = logging.getLogger('ErrorLogger')


    # A method to log an error message
    def log_error(self, message):
        # Use the logger object to log the error message
        self.logger.error(message)


    # A method to log an exception with traceback
    def log_exception(self, exception):
        # Use the logger object to log the exception with traceback
        self.logger.exception(exception)


    # A method to change the logging level
    def set_level(self, level):
        # Use the logger object to set the logging level
        self.logger.setLevel(level)


    # A method to change the logging format
    def set_format(self, format):
        # Create a formatter object with the given format
        formatter = logging.Formatter(format)
        # Get the handler object from the logger object
        handler = self.logger.handlers[0]
        # Set the formatter for the handler object
        handler.setFormatter(formatter)


    # A method to change the logging date format
    def set_datefmt(self, datefmt):
        # Create a formatter object with the given date format
        formatter = logging.Formatter(datefmt=datefmt)
        # Get the handler object from the logger object
        handler = self.logger.handlers[0]
        # Set the formatter for the handler object
        handler.setFormatter(formatter)