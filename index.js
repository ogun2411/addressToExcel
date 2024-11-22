import puppeteer from "puppeteer";

const XLSX = require("xlsx");

// Load excel file with addresses
const workbook = XLSX.readFile("SolarApi Locations.xlsx");

//access the first sheet
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

// Convert to JSON
const data = XLSX.utils.sheet_to_json(sheet);

data.forEach((row) => {
  console.log(row.Name + " " + row.Address + "\n");
  lonLat = getLongLat(JSON.stringify(row.Address));
});

// retreives longitude and latitude
async function getLongLat(address) {
  // Launch the browser and open a new blank page
  const browser = await puppeteer.launch();
  const page = await browser.newPage();

  // Navigate the page to a URL.
  await page.goto("https://www.latlong.net/convert-address-to-lat-long.html");

  // Set screen size.
  await page.setViewport({ width: 1080, height: 1024 });

  // Type into address box.
  await page.locator(".width70").fill(address);

  // Wait and click on find button.
  await page.locator(".btnfind").click();

  // wait for the .lat and .lng field to load
  await page.waitForSelector(".lat");

  // Record the values
  latitude = await page.$eval(".lat", (input) => input.value);
  longitude = await page.$eval(".lng", (input) => input.value);

  // Print the Longitude and Latitude.
  console.log(
    "The Latitude is $f and the latitude is %f\n.",
    latitude,
    longitude,
  );

  await browser.close();
}
