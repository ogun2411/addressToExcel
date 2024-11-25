const { chromium } = require("playwright");

const XLSX = require("xlsx");

async function main() {
  // Load excel file with addresses
  const workbook = XLSX.readFile("SolarApiLocations.xlsx");

  //access the first sheet
  const sheetName = workbook.SheetNames[0];
  sheet = workbook.Sheets[sheetName];

  // Convert to JSON
  const data = XLSX.utils.sheet_to_json(sheet);

  for (const row of data) {
    try {
      console.log(`Processing address: ${row.Address}`);

      // Get longitude and latitude
      const lonLat = await getLongLat(row.Address);

      // Add them to the row
      row.Latitude = lonLat[0];
      row.Longitude = lonLat[1];

      console.log(`Processed row:`, row);
      process.exit();
    } catch (error) {
      console.error(
        `Error processing row with address "${row.Address}":`,
        error,
      );
      process.exit();
    }
  }

  sheet = XLSX.utils.json_to_sheet(data);

  workbook.Sheets[sheetName] = sheet;

  XLSX.writeFile(wb, "SolarApi Locations.xlsx");
}

// retreives longitude and latitude
async function getLongLat(address) {
  // Launch the browser and open a new blank page
  const browser = await chromium.launch({ headless: false });
  const page = await browser.newPage();

  // Navigate the page to a URL.
  await page.goto("https://www.itilog.com/");

  // Set screen size.
  await page.setViewportSize({ width: 1080, height: 1024 });

  // Clear the input field for address
  await page.locator("#address").clear();

  // Wait for the input field to be cleared
  await page.waitForSelector("#address");

  // Type into address box.
  await page.locator("#address").fill(address);

  // Wait then click on Find GPS Coordinates button.
  await page.waitForSelector("#address");
  await page.locator("#address_to_map").click();

  // Wait for the latitude and longitude fields to have non-empty values
  await page.waitForFunction(() => {
    const latitudeField = document.querySelector("#latitude");
    const longitudeField = document.querySelector("#longitude");
    return (
      latitudeField?.value.trim() !== "" && longitudeField?.value.trim() !== ""
    );
  });

  // wait for the .lat and .lng field to load
  // await page.waitForSelector("#latitude", { timeout: 60000 });
  //await page.waitForSelector("#longitude", { timeout: 60000 });

  // Record the values
  const latitude = await page
    .locator("#latitude")
    .inputValue({ timeout: 60000 });
  const longitude = await page
    .locator("#longitude")
    .inputValue({ timeout: 60000 });

  // transform string to float
  const lat = parseFloat(latitude);
  const lng = parseFloat(longitude);

  // Print the Longitude and Latitude.
  console.log({ latitude: latitude, longitude: longitude });

  const direction = [lat, lng];

  await browser.close();

  return direction;
}

// Check if this is the main module
if (require.main === module) {
  main().catch((error) => {
    console.error("An error occurred:", error);
    process.exit(1);
  });
}
