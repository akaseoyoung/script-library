function calculateActivityBycolor() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('yoursheetname');
  const startRow = 8;
  const endRow = 103;
  const numRows = endRow - startRow + 1;
  const numCols = sheet.getLastColumn();
  const dataRange = sheet.getRange(startRow, 1, numRows, numCols);
  const dataValues = dataRange.getValues();
  const colorsCount = {}; // Object to store color count for each column
  
  // Loop through each column
  for (let col = 2; col <= numCols; col++) { // Exclude A column
    const columnColors = {};
    // Loop through each row
    for (let row = 0; row < numRows; row++) {
      const cell = dataValues[row][col - 1];
      const color = cell ? sheet.getRange(startRow + row, col).getBackground() : null;
      if (color && color !== '#b7e1cd' && color !== '#000000') { // Exclude specific colors
        columnColors[color] = (columnColors[color] || 0) + 1;
      }
    }
    // Sort colors by count in descending order
    const sortedColors = Object.keys(columnColors).sort((a, b) => columnColors[b] - columnColors[a]);
    // Limit to 6 colors
    const topColors = sortedColors.slice(0, 6);
    // Fill cells in rows 105 to 110 with top colors
    for (let i = 0; i < topColors.length; i++) {
      sheet.getRange(105 + i, col).setBackground(topColors[i]);
      sheet.getRange(105 + i, col).setValue(columnColors[topColors[i]] / (numRows-1) * 100 + '%');
    }
    // Store color count for each column
    colorsCount[col] = columnColors;
  }
  
  Logger.log(colorsCount); // You can check the color count in the Logs
  
  // Notify user
  SpreadsheetApp.getUi().alert('활동 색상별 백분율 계산 완료!');
}
