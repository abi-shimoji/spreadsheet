/**
 * @customfunction
 */
function getFunctionName() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const fileName = spreadSheet.getName();
  const fileNameArr = fileName.split('_');

  let result = '';
  for (let i = 1; i < fileNameArr.length; i++) {
    if (i === 1) {
      result += `${fileNameArr[i]}:`;
      continue;
    }
    result += `${fileNameArr[i]}`;
    if (i !== fileNameArr.length - 1) {
      result += '_'
    }
  }

  return result;
}