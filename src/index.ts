import puppeteer, { Page } from "puppeteer";
import * as ExcelJS from "exceljs";
import * as fs from "fs";
import { readFile } from "fs/promises";

/**
 * Represents the data of a printer.
 * Contains the IP address, serial number, and model of the printer.
 */

interface PrinterData {
  ipAddress: string; // The ip address of the printer
  serialNumber: string; // The printer's serial number
  model: string; //The model of the printer
}

// ====== Data collection functions specific to each printer model ====== //

// ====== Samsung M4080FX ====== //
/**
 * Performs the scraping process for the "For Samsung" printer model.
 * @param ipAddress The IP address of the printer.
 * @param serialNumber The serial number of the printer.
 */

async function runScrapingForSamsung(ipAddress: string, serialNumber: string) {
  // Lauches a Puppeteer browser instace
  const browser = await puppeteer.launch({ headless: "new" });
  const page = await browser.newPage();

  try {
    // Accesses the printer's counters information page
    await page.goto(
      `http://${ipAddress}/sws.application/information/countersView.sws`
    );

    // Extract the serial number of the printer page
    const serialNumberText = await page.evaluate(() => {
      const serialNumberElement = document.querySelector("#snValue");
      return serialNumberElement?.textContent?.trim() || "";
    });

    if (serialNumberText === serialNumber) {
      // Extract the print data from the table
      const counterTotalData: string[][] = await page.evaluate(() => {
        const rows = Array.from(
          document.querySelectorAll("#counterTotalList tr")
        );
        const data: string[][] = [];

        rows.forEach((row) => {
          const columns = Array.from(row.querySelectorAll("td"));
          const counterNameElement = columns[0];
          const counterName = counterNameElement?.textContent?.trim();

          if (
            counterName === "Simplex Mono" ||
            counterName === "Frente e verso" ||
            counterName === "Total de impressões"
          ) {
            const rowData = [
              counterName,
              ...columns.slice(1).map((cell) => cell.textContent?.trim() || ""),
            ];
            data.push(rowData);
          }
        });

        return data;
      });

      // Create a new workbook
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Data");

      // Add data to worksheet
      addDataToWorksheet(counterTotalData, worksheet, "Counter Total");

      // Adjust column width
      worksheet.columns.forEach((column) => {
        column.width = 20;
      });

      // Save the .xlsx file
      const filePath = `${serialNumber}.xlsx`;
      await workbook.xlsx.writeFile(filePath);
      console.log(`Arquivo salvo: ${filePath}`);
    } else {
      console.log(`Número de série não corresponde para o IP ${ipAddress}`);
    }
  } catch (error) {
    console.error(`Ocorreu um erro ao processar o IP ${ipAddress}: ${error}`);
  } finally {
    // Add a timeout before closing the browser
    await page.waitForTimeout(2000); // Wait for 2 seconds

    await browser.close();
  }
}

// ====== HP Laser 408 ====== //
/**
 * Performs the scraping process for the "Model 408" printer model.
 * @param ipAddress The IP address of the printer.
 * @param serialNumber The serial number of the printer.
 */

async function runScrapingForModel408(ipAddress: string, serialNumber: string) {
  //  Lauches a Puppeteer browser instance
  const browser = await puppeteer.launch({
    headless: "new",
    ignoreHTTPSErrors: true,
  });

  // Create a new page
  const page = await browser.newPage();

  try {
    // Accesses the printer's counters information page
    await page.goto(
      `view-source:https://${ipAddress}/sws/app/information/counters/counters.json`
    );

    // Extracts the serial number from printer's page
    const [serialNumberElementHandle] = await page.$x(
      "/html/body/table/tbody/tr[2]/td[2]"
    );
    const property = await page.evaluate(
      (Element) => Element.textContent,
      serialNumberElementHandle
    );

    const serialNumberText = property?.split('"')[1];

    if (serialNumber === serialNumberText) {
      // Function to extract "Mono Simples" data
      const getMonoSimplesData = async (page: Page, xpath: string) => {
        const element = await page.$x(xpath);
        const property = await page.evaluate(
          (Element) => Element.textContent,
          element[0]
        );
        return property?.split(":")[1].trim().replace(",", "");
      };

      // XPaths for "Simples Mono" data
      const monoSimplesXpath = [
        "/html/body/table/tbody/tr[13]/td[2]",
        "/html/body/table/tbody/tr[16]/td[2]",
        "/html/body/table/tbody/tr[17]/td[2]",
      ];

      // Extracts "Simple Mono" data
      const [
        monoSimplesPrintData,
        monoSimplesReportsData,
        monoSimplesTotalData,
      ] = await Promise.all(
        monoSimplesXpath.map((xpath) => getMonoSimplesData(page, xpath))
      );

      //Function to extract "Duplex" data
      const getDuplexData = async (page: Page, xpath: string) => {
        const element = await page.$x(xpath);
        const property = await page.evaluate(
          (Element) => Element.textContent,
          element[0]
        );
        return property?.split(":")[1].trim().replace(",", "");
      };

      // XPaths for "Duplex" data
      const duplexXpath = [
        "/html/body/table/tbody/tr[23]/td[2]",
        "/html/body/table/tbody/tr[26]/td[2]",
        "/html/body/table/tbody/tr[27]/td[2]",
      ];

      // Extracts "Duplex" data
      const [duplexPrintData, duplexReportsData, duplexTotalData] =
        await Promise.all(
          duplexXpath.map((xpath) => getDuplexData(page, xpath))
        );

      // Function to extract "Total Prints" data
      const getTotalPrintsData = async (page: Page, xpath: string) => {
        const element = await page.$x(xpath);
        const property = await page.evaluate(
          (Element) => Element.textContent,
          element[0]
        );
        return property?.split(":")[1].trim().replace(",", "");
      };

      // XPaths for "Total Prints" data
      const totalPrintsPath = [
        "/html/body/table/tbody/tr[33]/td[2]",
        "/html/body/table/tbody/tr[36]/td[2]",
        "/html/body/table/tbody/tr[37]/td[2]",
      ];

      // Extracts "Total Prints" data
      const [
        totalPrintsPrintData,
        totalPrintsReportsData,
        totalPrintsTotalData,
      ] = await Promise.all(
        totalPrintsPath.map((xpath) => getTotalPrintsData(page, xpath))
      );

      // Creates a new Excel workbook using the ExcelJS library
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Sheet 1");

      const lines = ["Imprimir", "Relatório", "Total"];
      const rows = ["Mono Simples", "Duplex", "Total Impressões"];

      // Adds data columns
      worksheet.addRow(["", ...lines]);

      // Fill the data in the correct cells
      rows.forEach((row, rowIndex) => {
        const rowData: (number | undefined)[] = [];

        lines.forEach((line, lineIndex) => {
          let data: number | undefined;

          if (row === "Mono Simples") {
            data = [
              monoSimplesPrintData,
              monoSimplesReportsData,
              monoSimplesTotalData,
            ][lineIndex] as number | undefined;
          } else if (row === "Duplex") {
            data = [duplexPrintData, duplexReportsData, duplexTotalData][
              lineIndex
            ] as number | undefined;
          } else if (row === "Total Impressões") {
            data = [
              totalPrintsPrintData,
              totalPrintsReportsData,
              totalPrintsTotalData,
            ][lineIndex] as number | undefined;
          }

          rowData.push(data);
        });

        worksheet.addRow([row, ...rowData]);
      });

      // Writes the worksheet data to a buffer
      const buffer = await workbook.xlsx.writeBuffer();

      // Converts the buffer to a Unit8Array array
      const dataArray = new Uint8Array(buffer);

      // Save the .xlsx file
      fs.writeFileSync(`${serialNumber}.xlsx`, dataArray);
    } else {
      console.log("Impressora não corresponde ao serial number");
    }

    // Closes the browser
    await browser.close();
  } catch (error) {
    console.error("Error:", error);

    // Closes the browser
    await browser.close();

    // Re-throws the error
    throw error;
  }
}

// ====== HP E57540DN ====== //
/**
 * Performs the scraping process for the "Model 57" printer model.
 * @param ipAddress The IP address of the printer.
 * @param serialNumber The serial number of the printer.
 */

async function runScrapingForModel57(ipAddress: string, serialNumber: string) {
  // Lauches a Puppeteer browser instance
  const browser = await puppeteer.launch({
    headless: "new",
    ignoreHTTPSErrors: true,
  });

  // Create a new page
  const page = await browser.newPage();

  try {
    // Accesses the printer's counters information page
    await page.goto(
      `http://${ipAddress}/hp/device/InternalPages/Index?id=UsagePage`
    );

    // Extract the serial number of the printer page
    const serialNumberText = await page.evaluate(() => {
      const xpathExpression =
        '//*[@id="UsagePage.DeviceInformation.DeviceSerialNumber"]';
      const serialNumberElement = document.evaluate(
        xpathExpression,
        document,
        null,
        XPathResult.FIRST_ORDERED_NODE_TYPE,
        null
      ).singleNodeValue;
      return serialNumberElement?.textContent || "";
    });

    if (serialNumberText === serialNumber) {
      // Extract the print data from the table
      const counterTotalData: string[][] = await page.evaluate(() => {
        const data: string[][] = [];

        // Select the desired rows using the correct selectors
        const rows = Array.from(
          document.querySelectorAll(
            "#UsagePage\\.EquivalentImpressionsTable tbody tr, #UsagePage\\.EquivalentImpressionsTable tfoot tr"
          )
        );

        rows.forEach((row) => {
          const columns = Array.from(row.querySelectorAll("td"));

          // Check if the line contains the desired data
          if (columns.length === 4) {
            const rowData = columns.map(
              (cell) => cell.textContent?.trim() || ""
            );
            data.push(rowData);
          }
        });

        return data;
      });

      // Create a new workbook
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Data");

      // Add data to worksheet
      addDataToWorksheet(counterTotalData, worksheet, "Counter Total");

      // Adjust column widths
      worksheet.columns.forEach((column) => {
        column.width = 20;
      });

      // Save the .xlsx file
      const filePath = `${serialNumber}.xlsx`;
      await workbook.xlsx.writeFile(filePath);
      console.log(`Arquivo salvo: ${filePath}`);
    } else {
      console.log(`Número de série não corresponde para o IP ${ipAddress}`);
    }
  } catch (error) {
    console.error(`Ocorreu um erro ao processar o IP ${ipAddress}: ${error}`);
  } finally {
    // Add a timeout before closing the browser
    await page.waitForTimeout(2000); // Wait for 2 seconds

    // Close the browser
    await browser.close();
  }
}

// ====== HP E52645DN ====== //
/**
 * Performs the scraping process for the "Model 52" printer model.
 * @param ipAddress The IP address of the printer.
 * @param serialNumber The serial number of the printer.
 */

async function runScrapingForModel52(ipAddress: string, serialNumber: string) {
  const browser = await puppeteer.launch({
    headless: "new",
    ignoreHTTPSErrors: true,
  });
  const page = await browser.newPage();

  try {
    await page.goto(
      `http://${ipAddress}/hp/device/InternalPages/Index?id=UsagePage`
    );

    // Extract the serial number of the printer page
    const serialNumberText = await page.evaluate(() => {
      const xpathExpression =
        '//*[@id="UsagePage.DeviceInformation.DeviceSerialNumber"]';
      const serialNumberElement = document.evaluate(
        xpathExpression,
        document,
        null,
        XPathResult.FIRST_ORDERED_NODE_TYPE,
        null
      ).singleNodeValue;
      return serialNumberElement?.textContent || "";
    });

    if (serialNumberText === serialNumber) {
      // Extract the print data from the table
      const counterTotalData: string[][] = await page.evaluate(() => {
        const data: string[][] = [];

        // Select the desired rows using the correct selectors
        const rows = Array.from(
          document.querySelectorAll(
            "#UsagePage\\.EquivalentImpressionsTable tbody tr, #UsagePage\\.EquivalentImpressionsTable tfoot tr"
          )
        );

        rows.forEach((row) => {
          const columns = Array.from(row.querySelectorAll("td"));

          // Check if the line contains the desired data
          if (columns.length === 2) {
            const rowData = columns.map(
              (cell) => cell.textContent?.trim() || ""
            );
            data.push(rowData);
          }
        });

        return data;
      });

      // Create a new workbook
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Data");

      // Add data to worksheet
      function addDataToWorksheet(
        data: string[][],
        worksheet: ExcelJS.Worksheet,
        tableName: string
      ) {
        // Add the table header
        worksheet.addRow([tableName]);
        worksheet.addRow(["Type", "Total"]);

        // Add the table data
        data.forEach((row) => {
          worksheet.addRow(row);
        });

        // Add an empty row after the table
        worksheet.addRow([]);
      }

      addDataToWorksheet(counterTotalData, worksheet, "Counter Total");

      // Adjust column widths
      worksheet.columns.forEach((column) => {
        column.width = 20;
      });

      // Save the .xlsx file
      const filePath = `${serialNumber}.xlsx`;
      await workbook.xlsx.writeFile(filePath);
      console.log(`Arquivo salvo: ${filePath}`);
    } else {
      console.log(`Número de série não corresponde para o IP ${ipAddress}`);
    }
  } catch (error) {
    console.error(`Ocorreu um erro ao processar o IP ${ipAddress}: ${error}`);
  } finally {
    // Add a timeout before closing the browser
    await page.waitForTimeout(2000); // Wait for 2 seconds

    // Close the browser
    await browser.close();
  }
}

/**
 * Adds data to an Excel sheet.
 *
 * @param {string[][]} data - The data to be added to the worksheet. It is a 2D array of strings, where each element represents a cell in the worksheet.
 * @param {ExcelJS.Worksheet} worksheet - The worksheet where the data will be added.
 * @param {string} header - The header that will be added before the data.
 */

function addDataToWorksheet(
  data: string[][],
  worksheet: ExcelJS.Worksheet,
  header: string
) {
  worksheet.addRow([header]);
  data.forEach((row) => {
    worksheet.addRow(row);
  });
  worksheet.addRow([]);
}

/**
 * Reads the file printers.json and returns the printers' data.
 * @returns A Promise that resolves on an array of PrinterData objects containing information about the printers.
 */

async function readPrintersDataFromFile(): Promise<PrinterData[]> {
  const jsonData = await readFile("src/printers.json", "utf-8");
  const data = JSON.parse(jsonData);
  const printersData: PrinterData[] = data.printers.flatMap((printer: any) => {
    const printerName = Object.keys(printer)[0];
    const printerDetailsArray = printer[printerName];
    if (Array.isArray(printerDetailsArray)) {
      return printerDetailsArray.map((printerDetails: any) => ({
        // Return the data to be used to define the model, serial number and address of the printers
        ipAddress: printerDetails.ipAddress,
        serialNumber: printerDetails.serialNumber,
        model: printerDetails.model,
      }));
    }
    return [];
  });

  /**
   * - Reads the content of the "printers.json" file and parses it as JSON data.
   * - The data is then mapped to an array of PrinterData objects, containing the IP address, serial number, and model of the printers.
   * - The array of PrinterData objects is returned as the result.
   */

  return printersData;
}

/**
 * - Main function
 * - It reads the printers' data from the printers.json file and scraps it according to the template.
 */

async function main() {
  // Calls the file printers.json with the information about the printers
  const printersData = await readPrintersDataFromFile();

  // modelScrapingFunctions gets the models from the printers
  const modelScrapingFunctions = {
    "Samsung M4080FX": runScrapingForSamsung,
    "HP Laser 408": runScrapingForModel408,
    "HP E57540DN": runScrapingForModel57,
    "HP E52645DN": runScrapingForModel52,
  } as {
    [key: string]: (ipAddress: string, serialNumber: string) => Promise<void>;
  };

  for (const printerData of printersData) {
    const { ipAddress, serialNumber, model } = printerData;
    const scrapingFunction = modelScrapingFunctions[model];

    if (scrapingFunction) {
      await scrapingFunction(ipAddress, serialNumber);
    } else {
      console.log(`Printer model not supported for IP: ${ipAddress}`);
    }
  }

  /**
   * - This function is the entry point of the program.
   * - It retrieves the printers' data using the `readPrintersDataFromFile` function.
   * - It defines the `modelScrapingFunctions` object which maps printer models to corresponding scraping functions.
   * - It iterates over each printer in `printersData` and performs scraping based on the printer model.
   * - If a scraping function is not available for a printer model, it displays a message indicating the lack of support.
   */
}

main();
