/**
 * GetEvaluationStatusAndTopInstructor 
 *  ✔ Reads evaluation rows from RatingsData (Excel table) * GetEvaluationStatusAndTopInstructor
 *  ✔ Loads admin-configured ExpectedSubmissions and DeadlineDate
 *  ✔ Determines whether deadline has passed
 *  ✔ Computes the top instructor and highest score
 *  ✔ Returns a workflow status for Power Automate:
 *      - "Complete"          → All submissions received before deadline
 *      - "DeadlineReached"   → Deadline passed (send results anyway)
 *      - "InProgress"        → Waiting for more submissions
 */

function main(workbook: ExcelScript.Workbook) {

  // -------------------------------------------
  // 1. Load evaluation data
  // -------------------------------------------
  const dataSheet = workbook.getWorksheet("RatingsData");
  const ratingsTable = dataSheet.getTable("Ratings");

  const ratingsValues = ratingsTable.getRange().getValues();
  const actualSubmissions = ratingsValues.length - 1; // minus header row


  // -------------------------------------------
  // 2. Load admin settings (expected submissions + deadline)
  // -------------------------------------------
  const settingsSheet = workbook.getWorksheet("Settings");

  // Expected submissions count
  const expectedSubmissions = settingsSheet
    .getRange("B1")
    .getValue() as number;

  // Deadline value from Excel (stored as text or date)
  const deadlineCellValue = settingsSheet.getRange("B2").getValue();

  // Convert deadline into a usable Date object
  // This handles both Excel serial and ISO string formats
  let deadlineDate: Date;

  if (typeof deadlineCellValue === "number") {
    // Excel serial date → JS date
    deadlineDate = new Date((deadlineCellValue - 25569) * 86400 * 1000);
  } else {
    deadlineDate = new Date(deadlineCellValue.toString());
  }

  const currentDate = new Date();
  const deadlineReached = currentDate >= deadlineDate;


  // -------------------------------------------
  // 3. Compute top instructor
  // -------------------------------------------
  let topInstructor = "";
  let topScore = -Infinity;

  // ratingsValues structure:
  // Row 0 = header
  // Columns: [ParticipantEmail, Instructor, Score, Timestamp]
  for (let i = 1; i < ratingsValues.length; i++) {
    const instructor = ratingsValues[i][1];
    const score = ratingsValues[i][2];

    if (typeof score === "number" && score > topScore) {
      topScore = score;
      topInstructor = instructor;
    }
  }

  // -------------------------------------------
  // 4. Determine workflow status
  // -------------------------------------------
  let status: string;

  if (actualSubmissions >= expectedSubmissions) {
    status = "Complete"; // All ratings received
  } else if (deadlineReached) {
    status = "DeadlineReached"; // Deadline reached, send with partial data
  } else {
    status = "InProgress"; // Still waiting for more submissions
  }

  // -------------------------------------------
  // 5. Return everything as JSON to Power Automate
  // -------------------------------------------
  return {
    expectedSubmissions,
    actualSubmissions,
    deadline: deadlineDate.toISOString(),
    deadlineReached,
    topInstructor,
    topScore,
    status
  };
}
 * -----------------------------------
 * This script:
