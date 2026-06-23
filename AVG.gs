/**
 * Calculates dynamic time averages for RMs.
 * @param {Array<Array<String>>|String} namesRange The column of RM names.
 * @param {Array<Array<String>>|String} dataRange The matrix of times.
 * @customfunction
 */
function GET_RM_AVERAGES(namesRange, dataRange) {
  // 1. Handle single-cell inputs (Sheets passes single cells as raw values, not arrays)
  if (!Array.isArray(namesRange)) namesRange = [[namesRange]];
  if (!Array.isArray(dataRange)) dataRange = [[dataRange]];

  let output = [];

  // Loop through every row in the provided range
  for (let r = 0; r < namesRange.length; r++) {
    let name = namesRange[r][0];
    
    // 2. Fix Alignment: Don't skip empty rows entirely. 
    // Push an empty row to keep the output aligned with the input rows in the Sheet.
    if (!name || name.toString().trim() === "") {
      output.push(["", ""]); 
      continue; 
    }

    // 3. Prevent undefined errors if dataRange is somehow shorter than namesRange
    let rowData = dataRange[r] || [];
    let totalSeconds = 0;
    let count = 0;

    // Loop horizontally through the time entries for this specific RM
    for (let c = 0; c < rowData.length; c++) {
      let cell = rowData[c] ? rowData[c].toString().trim() : "";
      
      // Skip blanks and dashes immediately
      if (cell === "" || cell === "—" || cell === "-") continue;

      // If the cell contains a number, extract the time
      if (/\d/.test(cell)) {
        let hMatch = cell.match(/(\d+)h/);
        let mMatch = cell.match(/(\d+)m/);
        let sMatch = cell.match(/(\d+)s/);
        let dMatch = cell.match(/(\d+)d/);

        let d = dMatch ? parseInt(dMatch[1]) : 0;
        let h = hMatch ? parseInt(hMatch[1]) : 0;
        let m = mMatch ? parseInt(mMatch[1]) : 0;
        let s = sMatch ? parseInt(sMatch[1]) : 0;

        // Convert everything to seconds
        let seconds = (d * 86400) + (h * 3600) + (m * 60) + s;
        totalSeconds += seconds;
        count++; // Only count actual entries
      }
    }

    let avgStr = "—"; // Default if they have zero logged time
    
    // Calculate the average and convert back to d h m s
    if (count > 0) {
      let avgSeconds = Math.round(totalSeconds / count);
      let d = Math.floor(avgSeconds / 86400);
      avgSeconds %= 86400;
      let h = Math.floor(avgSeconds / 3600);
      avgSeconds %= 3600;
      let m = Math.floor(avgSeconds / 60);
      let s = avgSeconds % 60;

      let timeParts = [];
      if (d > 0) timeParts.push(d + "d");
      if (h > 0) timeParts.push(h + "h");
      if (m > 0) timeParts.push(m + "m");
      if (s > 0) timeParts.push(s + "s");
      avgStr = timeParts.join(" ") || "0s";
    }

    // Add the Name and the calculated Average to our final output list
    output.push([name, avgStr]);
  }

  // Fallback if no valid names exist
  if (output.length === 0) return [["—", "—"]];

  return output;
}