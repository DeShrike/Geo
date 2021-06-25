from dotenv import load_dotenv
from pathlib import Path
import os

# https://positionstack.com/documentation

def main():
    load_dotenv()
    env_path = Path('.')
    env_file = '.env'
    load_dotenv(dotenv_path=os.path.join(env_path, env_file))
    api_key = os.getenv("API_KEY")




if __name__ == "__main__":
    main()

