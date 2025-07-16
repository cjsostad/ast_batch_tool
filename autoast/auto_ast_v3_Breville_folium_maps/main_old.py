# autoast is a script for batch processing the automated status tool
# author: csostad and wburt
# copyright Governent of British Columbia
# Copyright 2019 Province of British Columbia

# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at 

#    http://www.apache.org/licenses/LICENSE-2.0

# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.


import os
from dotenv import load_dotenv
from logging_setup import setup_logging
from database_connection import setup_bcgw
from toolbox_import import import_ast
from ast_factory import AST_FACTORY

# snippet to run multiple terminal windows & "P:\corp\python_ast\python.exe" "W:\srm\nel\Local\Geomatics\Workarea\csostad\GitHubAutoAST\gss_authorizations\autoast\auto_ast_v3_Breville_folium_maps\main.py"


## *** INPUT YOUR EXCEL FILE NAME HERE ***
excel_file = 'july_7.xlsx'



#################################################################################################################################################################################
if __name__ == '__main__':
    

    # Call the setup_logging function to log the messages
    logger = setup_logging()

    # Load the default environment
    load_dotenv()

    # Call the import_ast function to import the AST toolbox
    template = import_ast(logger)

    # Call the setup_bcgw function to set up the database connection
    # secrets = setup_bcgw(logger)
    secrets, sde_connection, sde_path = setup_bcgw(logger)
    # username, password = secrets[0], secrets[1]
    
    # Set the SDE path environment variable for easy access by workers
    os.environ["SDE_FILE_PATH"] = sde_path
    logger.info(f"SDE Connection established at: {sde_path}")
    
    
    # Create the path for the queuefile
    current_path = os.path.dirname(os.path.realpath(__file__))
    qf = os.path.join(current_path, excel_file)

    # Create an instance of the Ast Factory class, assign the queuefile path and the bcgw username and passwords to the instance
    ast = AST_FACTORY(qf, secrets[0], secrets[1], logger, current_path)

    if not os.path.exists(qf):
        print("Main: Queuefile not found, creating new queuefile")
        logger.info("Main: Queuefile not found, creating new queuefile")
        ast.create_new_queuefile()
        
    # Load the jobs using the load_jobs method. This will scan the excel sheet and assign to "jobs"    
    jobs = ast.load_jobs()
    
    ast.batch_ast()
    
    ast.re_load_failed_jobs_V2()
    
    ast.batch_ast()
    
    print("Main: AST Factory COMPLETE")
    logger.info("Main: AST Factory COMPLETE")

