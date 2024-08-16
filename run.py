# Load environment variables from .env file
import os
from dotenv import load_dotenv

from o365.o365 import SharepointService


load_dotenv()

# Main function


def main():

    tenant_id = os.getenv('TENANT_ID')
    client_id = os.getenv('CLIENT_ID')
    client_secret = os.getenv('CLIENT_SECRET')
    sc = SharepointService(
        tenant_id=tenant_id,
        client_id=client_id,
        client_secret=client_secret
    )

    sc.run(year=2024)

    sc.dropDataFolder()


# Entry point of the script
if __name__ == '__main__':
    main()
