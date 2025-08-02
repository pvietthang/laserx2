from dotenv import load_dotenv
import os
def get_env(variable):
    load_dotenv(override=True)
    return os.getenv(variable)