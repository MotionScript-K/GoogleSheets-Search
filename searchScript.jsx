const source = SpreadsheetApp.getActive();
 
function runScript() {
 
var lastRow = source.getLastRow();
 
for(let x = 2712; x <= lastRow; x++){
  let targetObj = source.getRange(`A${x}`);
  let targetVal = targetObj.getValues();
 
  try{
    var textFinder = source.getRange("D:D").createTextFinder(targetVal).matchEntireCell(true).matchCase(false);
    var firstOccurrence = textFinder.findNext().getValues();
    let result = source.getRange(`B${x}`).setValue("YES");
 
  } catch {
    let result = source.getRange(`B${x}`).setValue("NO");
  }
}
 
}
