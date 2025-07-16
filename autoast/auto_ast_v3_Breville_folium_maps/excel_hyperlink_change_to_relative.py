#description: 
# this script updates a path in an excel document 
import openpyxl 
import os
import sys
import ctypes
import arcpy
import logging
from datetime import datetime as dt
from ctypes import wintypes
from argparse import ArgumentParser
from util.environment import Environment

def run_app() -> None:
    fld, logger = get_input_parameters()
    replace = ReplaceHyperlinks(folders=fld, logger=logger)
    replace.run_replacements()
    del replace

def get_input_parameters() -> tuple:
    try:
        parser = ArgumentParser(description='This script is used to replace hyperlinks within excel documents '\
                                'with their relative paths.  Specifically meant to be used with AST ')
        parser.add_argument('fld', type=str, help='Semi colon separated list of folders to run')
        parser.add_argument('--log_level', default='INFO', choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'], 
                            help='Log level for message output')
        parser.add_argument('--log_dir', help='Path to the log file directory')

        args = parser.parse_args()
        logger = Environment.setup_logger(args)

        return args.fld, logger

    except Exception as e:
        logging.error(f'Unexpected exception, program terminating: {e.message}')
        raise Exception('Errors exist')


class ReplaceHyperlinks:
    """
    ------------------------------------------------------------------------------------------------------------
        CLASS: Used to contain the methods for replacing hyperlinks
    ------------------------------------------------------------------------------------------------------------
    """
    def __init__(self, folders: str, logger: logging.Logger) -> None:
        """
        ------------------------------------------------------------------------------------------------------------
            FUNCTION: Class method for initializing the object

            Parameters:
                None

            Return: None
        ------------------------------------------------------------------------------------------------------------
        """
        self.folders = folders.split(';')
        self.xl_files = []
        self.logger = logger

        self.logger.info('Gathering excel files')
        # Get semi colon separated list of excel files from parameter
        for fld in self.folders:
            for root, dirs, files in os.walk(fld):
                for f in files:
                    if f.endswith('xls') or f.endswith('xlsx'):
                        self.xl_files.append(os.path.join(root, f))
    
    def run_replacements(self) -> None:
        """
        ------------------------------------------------------------------------------------------------------------
            FUNCTION: Run through the list of excel files and replace al hyperlinks with relative paths

            Parameters:
                None

            Return: None
        ------------------------------------------------------------------------------------------------------------
        """

        # Loop through all Excel files in list
        for xl in self.xl_files:
        
            # Get the full unc path if its a named drive
            xl = Environment.get_full_path(str_file=xl)
            xl_head = os.path.dirname(xl)

            # Open the workbook
            self.logger.info(f'Opening workbook: {xl}')
            wb = openpyxl.load_workbook(filename=xl)

            # Loop through the sheets within the workbook
            for sheet in wb.worksheets:
            
                # iterate through all cells containing hyperlinks
                for row in sheet.iter_rows():
                    for cell in row:
                        # Check if the cell is hyperlinked
                        if cell.hyperlink:
                            # Remove the built-in file prefix
                            target = cell.hyperlink.target.replace('file:///','')

                            # If the hyperlink is to a website, ignore and continue to the next item
                            if target.startswith('https://'):
                                continue
                            
                            target = Environment.get_full_path(str_file=target)

                            # Extract the parts of the hyperlink 
                            target_head = os.path.dirname(target)
                            target_file = os.path.basename(target)

                            # Find the common directories and return the relative path
                            try:
                                rel_head = os.path.relpath(target_head, xl_head)
                            except:
                                rel_head = target_head

                            # Set the new target hyperlink ot the cell
                            new_target = os.path.join(rel_head, target_file)
                            cell.hyperlink.target = new_target
                            self.logger.info(f'Replaced {target} with {new_target}')

            self.logger.info('Saving workbook')
            wb.save(xl)
            wb.close()


if __name__ == '__main__':
    run_app()
