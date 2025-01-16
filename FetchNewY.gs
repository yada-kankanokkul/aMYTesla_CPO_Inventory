function rawTeslaModelYInventory() {
  const mainSheetName = 'inventoryNewY';  // Specify the main sheet name

  const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mainSheetName);

  if (!mainSheet) {
    Logger.log('Sheet not found!');
    return;
  }

  const baseUrl = 'https://www.tesla.com/inventory/api/v4/inventory-results';  // Tesla API endpoint

  // Define the base query parameters
  const baseQueryPayload = {
    "query": {
      "model": "my",
      "condition": "new",
      "arrangeby": "Price",
      "order": "asc",
      "market": "US",
      "language": "en",
      "super_region": "north america",
      "PaymentType": "cash",
      "lng": -122.2031,
      "lat": 47.8948,
      "zip": "98208",
      "range": 0,
      "region": "WA"
    },
    "count": 50,
    "isFalconDeliverySelectionEnabled": false,
    "version": null
  };

  // Queries to execute
  const queries = [
    { ...baseQueryPayload, offset: 0, outsideOffset: 0, outsideSearch: false },
    { ...baseQueryPayload, offset: 0, outsideOffset: 0, outsideSearch: true },
    { ...baseQueryPayload, offset: 0, outsideOffset: 50, outsideSearch: true },
    { ...baseQueryPayload, offset: 0, outsideOffset: 100, outsideSearch: true }
  ];

  // Function to encode the query object into a URL parameter
  function encodeQueryParams(query) {
    return encodeURIComponent(JSON.stringify(query));
  }

  // Clear existing content in the main sheet
  mainSheet.clear();

  const outputHeaders = [
    'InventoryPrice', 
    'Discount',
    'TrimName',
    'TrimCode',
    'IsDemo',
    'Odometer',
    'VIN',
    'Link'
  ];

  // Fetch data and append to the sheet
  function fetchDataAndAppend(query) {
    const encodedQuery = encodeQueryParams(query);
    const url = `${baseUrl}?query=${encodedQuery}`;

    const response = UrlFetchApp.fetch(url);
    const jsonData = JSON.parse(response.getContentText());

    const cars = jsonData.results;  // Adjust if the actual data structure is different

    if (cars.length === 0) {
      Logger.log('No data found');
      return;
    }

    // Append headers to the sheet if not already done
    if (mainSheet.getLastRow() === 0) {
      mainSheet.appendRow(outputHeaders);
    }

    // Append car data to the sheet
    cars.forEach(car => {
      const link = `https://www.tesla.com/my/order/${car['VIN']}`;
      const inventoryPrice = car['Price'] || 0;
      const row = outputHeaders.map(header => {
        switch(header) {
          case 'Link':
            return link;
          case 'InventoryPrice':
            return inventoryPrice;
          case 'Discount':
            return car['Discount'] || ''; // Assuming 'Discount' is the key for Discount
          default:
            return car[header] || '';
        }
      });

      // Append the row to the sheet
      mainSheet.appendRow(row);

      // Highlight row in light green if 'InventoryPrice' is less than 45,000
      if (inventoryPrice < 45000) {
        const lastRow = mainSheet.getLastRow();  // Get the last row number
        mainSheet.getRange(lastRow, 1, 1, outputHeaders.length).setBackground('#A8E6A1');  // Light green color
      }
    });
  }

  // Execute all queries and append the results
  queries.forEach(query => fetchDataAndAppend(query));
}
