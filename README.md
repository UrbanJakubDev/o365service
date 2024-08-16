# o365service

This script is designed to handle the o365 service.

## Installation

1. Clone the repository:

   ```shell
   https://github.com/UrbanJakubDev/o365service.git
   ```

2. Install the required dependencies:

   ```shell
   # make enviroment and activate that
   python3 venv .venv
   source .venv/bin/activate

   # install dependencies
   pip install -r requirements.txt

   
   ```

## Usage

To use this script, follow these steps:

1. Import the module:

   ```javascript
   const o365service = require('o365service');
   ```

2. Initialize the service:

   ```javascript
   const service = new o365service();
   ```

3. Call the desired methods to interact with the o365 service:

   ```javascript
   service.login('username', 'password');
   service.sendEmail('recipient@example.com', 'Hello', 'This is a test email.');
   ```

## Contributing

Contributions are welcome! If you find any issues or have suggestions for improvements, please open an issue or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE).
