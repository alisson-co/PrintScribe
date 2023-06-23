## PrintScribe - Printer Data Collection Automation

PrintScribe is a powerful automation tool designed to collect data from printers of various models. It leverages cutting-edge libraries to extract relevant information from printers and generates comprehensive reports based on the collected data. This automation process is made possible by utilizing Puppeteer, a versatile Node.js library that enables seamless control of headless Chrome or Chromium browsers.

### Key Features

- Efficiently collects data from printers of different models.
- Extracts essential information, including serial numbers and print counters.
- Generates detailed reports in the form of Excel files (.xlsx) based on the printer's serial number.

### Installation

1. Clone the PrintScribe repository:

```shell
git clone https://github.com/alisson-co/PrintScribe.git
```

2. Install the necessary dependencies:

```shell
cd PrintScribe
npm install
```

### Usage

1. Customize the printer models and define the data to be extracted by modifying the `printers.json` file.

2. Run the automation script:

```shell
npm run build
node dist/index.js
```

3. The extracted data will be automatically saved in separate Excel files, each corresponding to the respective printer's serial number.

### Dependencies

PrintScribe relies on the following libraries:

- Puppeteer: A Node.js library for automating browser actions.
- ExcelJS: A powerful library for generating Excel files.

### Contributing

Contributions are highly appreciated! If you encounter any issues or have ideas for improvement, please feel free to open an issue or submit a pull request.
