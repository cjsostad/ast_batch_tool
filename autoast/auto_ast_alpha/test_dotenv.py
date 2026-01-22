from dotenv import load_dotenv
import os

###############################################################################################################################################################
#
# Calling file paths
#
###############################################################################################################################################################


# Load the default environment
print("Test Dotenv: Loading .env file...")

# load_dotenv() (with no arguments) loads .env from the current working directory so if you run the script from a different folder it might not find it.
load_dotenv()

# Tell dotenv to get the variable SDE_FILE_PATH from the .env file and assign it to a variable
sde_path = os.getenv("SDE_FILE_PATH")


print(f"Test Dotenv: SDE_FILE_PATH is set to: {sde_path}")


###############################################################################################################################################################
#
# Accessing Credentials
#
###############################################################################################################################################################





# Load the file path to your secrets file into a a variable. At this point it is still just a path
secret_file = os.getenv("SECRET_FILE")

# A little bit of error handling
if not secret_file:
    raise RuntimeError("SECRET_FILE is not set. Check your project .env.")

# The path to the secret file is loaded into the load_dotenv function and dotenv opens it and loads 
# the dictionary that contains the username and password into environment variables
# Different from when we called file paths that were sitting in a .env file alongside the rest of the script
load_dotenv(secret_file)

# Now you can use getenv to access to your username and password and set those to variables to be used in a database connection

db_user   = os.getenv("BCGW_USER")
db_pass   = os.getenv("BCGW_PASS")

print("Username is:", db_user)

print("Have DB access? :", bool(db_user and db_pass))
