function fetchTeslaModelYInventory() {
  const mainSheetName = 'TestScript';  // Specify the main sheet name
  const recordSheetName = 'MyRecord';   // Specify the record sheet name

  const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mainSheetName);
  const recordSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(recordSheetName);

  if (!mainSheet || !recordSheet) {
    Logger.log('Sheet(s) not found!');
    return;
  }
  
  const baseUrl = 'https://www.tesla.com/inventory/api/v4/inventory-results';  // Tesla API endpoint

  // Define the base query parameters
  const baseQueryPayload = {
    "query": {
      "model": "my",
      "condition": "new",
      //"condition": "used",
      "options": {
        "TRIM": ["PAWD", "LRAWD", "LRRWD"],
        //"AUTOPILOT": ["AUTOPILOT_FULL_SELF_DRIVING"],
        //"Year": ["2024", "2023", "2022"]
      },
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

  // Clear existing content and set up headers in main sheet
  mainSheet.clear();

  // Define the headers to output in main sheet, including Discount
  const outputHeaders = [
    'TotalPrice', // Calculated as Price + TransportationFee
    'Discount',   // New Discount column
    'YearBuild',
    'WarrantyLeft', 
    'Warranty',
    'Odometer',
    'ProperMiles', 
    'TrimName',
    'TrimCode',
    'TransportationFee',
    'MetroName',
    'TitleSubStatus',
    'VehicleHistory',
    'IsChargingConnectorIncluded',
    'IsDemo',
    'OnConfiguratorPricePercentage',
    'OwnerShipTransferCount',
    'RefurbishmentEstimateStatus',
    'VIN',
    'added',
    'PricingDate',
    'PriceBookName',
    'Link'  // New Link column
  ];
  
  // Append headers to the main sheet
  mainSheet.appendRow(outputHeaders);

  // Function to fetch and append data
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

    // Helper function to format dates
    function formatDate(dateString, format) {
      if (!dateString) return '';
      const date = new Date(dateString);
      const year = date.getFullYear();
      const month = ("0" + (date.getMonth() + 1)).slice(-2);
      const day = ("0" + date.getDate()).slice(-2);
      switch (format) {
        case 'YearMonth':
          return `${year}-${month}`;
        case 'YearMonthDay':
          return `${year}-${month}-${day}`;
        default:
          return dateString;
      }
    }

    // Helper function to calculate the difference in months between two dates
    function monthsDiff(date1, date2) {
      const year1 = date1.getFullYear();
      const month1 = date1.getMonth();
      const year2 = date2.getFullYear();
      const month2 = date2.getMonth();
      return (year2 - year1) * 12 + (month2 - month1);
    }

    // Populate main sheet with car data
    cars.forEach(car => {
      const totalPrice = (car['Price'] || 0) + (car['TransportationFee'] || 0);
      // const discount = car['Discount'] || 0;  // Get the discount value

      const warrantyVehicleExpDate = new Date(car['WarrantyVehicleExpDate']);
      const monthWarrantyVehicleExpDate = ("0" + (warrantyVehicleExpDate.getMonth() + 1)).slice(-2);
      const yearWarrantyVehicleExpDate = warrantyVehicleExpDate.getFullYear();
      
      const warrantyBatteryExpDate = new Date(car['WarrantyBatteryExpDate']);
      const yearWarrantyBatteryExpDate = warrantyBatteryExpDate.getFullYear();

      const warranty = `${monthWarrantyVehicleExpDate}-${yearWarrantyVehicleExpDate}/${yearWarrantyBatteryExpDate}`;

      const today = new Date();
      const warrantyLeft = monthsDiff(today, warrantyVehicleExpDate);

      const properMiles = (48 - warrantyLeft) * 1042;

      const yearBuild = formatDate(car['ManufacturingYear'], 'YearMonth');
      const added = formatDate(car['RefurbishmentCompletionETA'], 'YearMonthDay');

      const link = `https://www.tesla.com/my/order/${car['VIN']}`;

      const row = outputHeaders.map(header => {
        switch(header) {
          case 'TotalPrice':
            return totalPrice;
          // case 'Discount':
          //   return discount;  // Add Discount value
          case 'Warranty':
            return warranty;
          case 'WarrantyLeft':
            return warrantyLeft >= 0 ? warrantyLeft : '';
          case 'ProperMiles':
            return properMiles >= 0 ? properMiles : '';
          case 'YearBuild':
            return yearBuild;
          case 'added':
            return added;
          case 'Link':
            return link;
          default:
            return car[header] || '';
        }
      });

      // Add data to the main sheet
      const mainLastRow = mainSheet.getLastRow();
      mainSheet.appendRow(row);

      // Highlight row if ProperMiles and Odometer difference is within 10%
      if (properMiles >= 0 && Math.abs(properMiles - car['Odometer']) <= 0.1 * properMiles) {
        mainSheet.getRange(mainLastRow + 1, 1, 1, outputHeaders.length).setBackground('#A8C6EA');
        const properMilesColumnIndex = outputHeaders.indexOf('ProperMiles') + 1;
        const odometerColumnIndex = outputHeaders.indexOf('Odometer') + 1;
        mainSheet.getRange(mainLastRow + 1, properMilesColumnIndex).setBackground('#6699FF');
        mainSheet.getRange(mainLastRow + 1, odometerColumnIndex).setBackground('#6699FF');
      }

      // Highlight row if WarrantyLeft >= 30 and TotalPrice <= 43000
      if (warrantyLeft >= 30 && totalPrice <= 43000) {
        mainSheet.getRange(mainLastRow + 1, 1, 1, outputHeaders.length).setBackground('#AED49B');
        const totalPriceMilesColumnIndex = outputHeaders.indexOf('TotalPrice') + 1;
        const warrantyLeftColumnIndex = outputHeaders.indexOf('WarrantyLeft') + 1;
        mainSheet.getRange(mainLastRow + 1, totalPriceMilesColumnIndex).setBackground('#C0F9A5');
        mainSheet.getRange(mainLastRow + 1, warrantyLeftColumnIndex).setBackground('#C0F9A5');

        // Check if VIN already exists in record sheet
        const existingRow = recordSheet.createTextFinder(car['VIN']).findNext();
        if (existingRow) {
          const existingRowIndex = existingRow.getRow();
          recordSheet.deleteRow(existingRowIndex);
        } 
        // Append the row to the record sheet
        recordSheet.appendRow(row);
        
      }
    });
    
  }

  // Execute both queries and append the results
  queries.forEach(query => fetchDataAndAppend(query));
}
