"""Load application configurations to environment variables."""

from os.path import join, dirname
from dotenv import load_dotenv

def load(path='.env'):
    """Globally loads the configuration for the
    project in environmenv variables.
        
    :path: The argument containing the path for the
           env file. Defaults to PROJECT_ROOT/.env
    """
    dotenv_path = join(dirname(dirname(__file__)), path)
    load_dotenv(dotenv_path)
