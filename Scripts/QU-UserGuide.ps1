###############################################################################
# SCRIPT NAME: QU-UserGuide
#
# AUTHOR:
#   John Buxton
#
# PURPOSE:
#   Provides an operator-friendly User Guide viewer for the QU application.
#   Displays individual sections or the entire guide through a simple menu.
#
# PARAMETER (INPUT):
#   None
#
# PARAMETER (OUTPUT):
#   None
#
# RETURN:
#   None
#
# HELPER FUNCTIONS:
#   Show-UG-Overview
#   Show-UG-Lifecycle
#   Show-UG-Workflow
#   Show-UG-Scripts
#   Show-UG-StatusCodes
#   Show-UG-Safety
#   Show-UG-FileNaming
#   Show-UG-HighLevelDiagram
#   Show-UG-FileMessagesFlow
#   Show-UG-Troubleshooting
#   Show-UG-All
#
# CHANGE LOG:
#   v1.0 (2026-01-11) - Initial release of User Guide viewer.
#   v1.1 (2026-01-12) - Added High-Level Overview Diagram section.
###############################################################################

$version = "v1.1";


###############################################################################
# SECTION DISPLAY FUNCTIONS
###############################################################################

function Show-UG-Overview {
  Clear-Host
  Write-Host "==================== USER GUIDE: OVERVIEW ====================" -ForegroundColor Cyan
  Write-Host ""
  Write-Host "The QU Application processes backlog MSMQ messages by:"
  Write-Host "  • Moving messages from MSMQ → SQL quarantine table"
  Write-Host "  • Parsing DSP, timestamp, and SMI fields"
  Write-Host "  • Removing nonessential messages (SMI='DFD' or DBLINK1)"
  Write-Host "  • Displaying message counts by status"
  Write-Host "  • Writing Ready messages to the Aircom input folder"
  Write-Host ""
  Write-Host "This ensures only essential messages are filed for Aircom processing."
  Write-Host ""
  Pause
}

function Show-UG-Lifecycle {
  Clear-Host
  Write-Host "==================== USER GUIDE: MESSAGE LIFECYCLE ====================" -ForegroundColor Cyan
  Write-Host ""
  Write-Host "Message Status Flow:"
  Write-Host ""
  Write-Host "  Pending   → After MSMQ messages are moved into SQL"
  Write-Host "  Parsed    → After DSP/SMI/timestamp fields are extracted"
  Write-Host "  Ready     → After nonessential messages are removed"
  Write-Host "  Processed → After messages are written to Aircom input folder"
  Write-Host ""
  Write-Host "Each stage represents a clean checkpoint in the pipeline."
  Write-Host ""
  Pause
}

function Show-UG-Workflow {
  Clear-Host
  Write-Host "==================== USER GUIDE: OPERATOR WORKFLOW ====================" -ForegroundColor Cyan
  Write-Host ""
  Write-Host "Correct order of operations:"
  Write-Host ""
  Write-Host "  1. Move QU Messages Into Quarantine"
  Write-Host "  2. Parse DSP DayTime SMI"
  Write-Host "  3. Remove Nonessential Messages"
  Write-Host "  4. Display Message Status Summary"
  Write-Host "  5. Move QU Messages to Aircom Input Folder"
  Write-Host ""
  Write-Host "Following this order ensures clean, predictable processing."
  Write-Host ""
  Pause
}

function Show-UG-Scripts {
  Clear-Host
  Write-Host "==================== USER GUIDE: SCRIPT DESCRIPTIONS ====================" -ForegroundColor Cyan
  Write-Host ""
  Write-Host "QU-Move-AllMsmqMessagesToDatabase:"
  Write-Host "  • Reads MSMQ messages"
  Write-Host "  • Inserts into SQL quarantine table"
  Write-Host "  • Sets Status='Pending'"
  Write-Host ""
  Write-Host "QU-Update-MsmqParsedFields:"
  Write-Host "  • Extracts DSP, SMI, timestamp"
  Write-Host "  • Regenerates BodyBinary"
  Write-Host "  • Sets Status='Parsed'"
  Write-Host ""
  Write-Host "QU-Remove-NonessentialMessages:"
  Write-Host "  • Removes SMI='DFD' and DBLINK1 messages"
  Write-Host "  • Sets remaining Parsed → Ready"
  Write-Host ""
  Write-Host "QU-Display-MessageStatusSummary:"
  Write-Host "  • Shows counts by Status"
  Write-Host ""
  Write-Host "QU-File-AllMessagesAdaptive:"
  Write-Host "  • Writes Ready messages to Aircom input folder"
  Write-Host "  • Deletes SQL rows after written to Aircom input folder"
  Write-Host "  • Applies queue-depth safety throttling"
  Write-Host ""
  Pause
}

function Show-UG-StatusCodes {
  Clear-Host
  Write-Host "==================== USER GUIDE: STATUS CODES ====================" -ForegroundColor Cyan
  Write-Host ""
  Write-Host "Pending:"
  Write-Host "  Message has been moved from MSMQ into SQL but not parsed yet."
  Write-Host ""
  Write-Host "Parsed:"
  Write-Host "  DSP, SMI, and timestamp fields extracted."
  Write-Host ""
  Write-Host "Ready:"
  Write-Host "  Message passed filtering and is ready to be filed."
  Write-Host ""
  Write-Host "Processed:"
  Write-Host "  Message has been written to Aircom input folder."
  Write-Host "  The 'Processed' messages are immediately deleted from the database table"
  Write-Host ""
  Pause
}

function Show-UG-Safety {
  Clear-Host
  Write-Host "==================== USER GUIDE: SAFETY CONTROLS ====================" -ForegroundColor Cyan
  Write-Host ""
  Write-Host "Queue-depth safety logic:"
  Write-Host ""
  Write-Host "  • Queue depth is checked once per batch, immediately AFTER:"
  Write-Host "        – All files in the current batch have been written"
  Write-Host "        – All corresponding SQL rows have been deleted"
  Write-Host ""
  Write-Host "  • A fixed 6-second pause occurs BEFORE each queue-depth check."
  Write-Host "        This ensures Aircom has time to drain messages from MSMQ"
  Write-Host "        before the next batch is evaluated."
  Write-Host ""
  Write-Host "  • Warning threshold:"
  Write-Host "        – If queue depth ≥150 → Display WARNING but continue processing"
  Write-Host ""
  Write-Host "  • Hard-stop threshold:"
  Write-Host "        – If queue depth >200 → STOP all processing immediately"
  Write-Host "        – System enters a wait loop"
  Write-Host "        – Queue depth is rechecked every 6 seconds"
  Write-Host "        – Processing resumes ONLY when queue depth <50"
  Write-Host ""
  Pause
}

function Show-UG-FileNaming {
  Clear-Host
  Write-Host "==================== USER GUIDE: FILE NAMING ====================" -ForegroundColor Cyan
  Write-Host ""
  Write-Host "Requeue file naming format:"
  Write-Host ""
  Write-Host "  Requeue_<SMI>_<DDHHMM>_<Sequence>.txt"
  Write-Host ""
  Write-Host "Example:"
  Write-Host "  Requeue_DFD_051230_00042.txt"
  Write-Host ""
  Write-Host "Fields:"
  Write-Host "  • SMI      = Service Message Indicator"
  Write-Host "  • DDHHMM   = Day, Hour, Minute from ParsedTimestamp"
  Write-Host "  • Sequence = Rolling counter (0–99999)"
  Write-Host ""
  Pause
}


###############################################################################
# NEW SECTION — HIGH-LEVEL OVERVIEW DIAGRAM
###############################################################################

function Show-UG-HighLevelDiagram {
  Clear-Host
  Write-Host "==================== USER GUIDE: HIGH-LEVEL OVERVIEW DIAGRAM ====================" -ForegroundColor Cyan
  Write-Host ""

  Write-Host "                          ┌──────────────────────────────┐"
  Write-Host "                          │        QU-Application        │"
  Write-Host "                          │     (Main Menu Dispatcher)   │"
  Write-Host "                          └──────────────┬───────────────┘"
  Write-Host "                                         │"
  Write-Host "                                         │"
  Write-Host "     ┌───────────────────────────────────┼───────────────────────────────────┐"
  Write-Host "     │                                   │                                   │"
  Write-Host "     ▼                                   ▼                                   ▼"
  Write-Host ""
  Write-Host "┌──────────────────────────┐   ┌──────────────────────────┐   ┌──────────────────────────┐"
  Write-Host "│ 1. Move MSMQ Messages    │   │ 2. Parse DSP/SMI Fields  │   │ 3. Remove Nonessential    │"
  Write-Host "│ QU-Move-AllMsmqMessages  │   │ QU-Update-MsmqParsed...  │   │ QU-Remove-Nonessential... │"
  Write-Host "└──────────────┬───────────┘   └──────────────┬───────────┘   └──────────────┬───────────┘"
  Write-Host "               │                               │                               │"
  Write-Host "               │                               │                               │"
  Write-Host "               ▼                               ▼                               ▼"
  Write-Host ""
  Write-Host "      ┌──────────────────┐          ┌──────────────────┐          ┌──────────────────┐"
  Write-Host "      │ MSMQ Queue       │          │ SQL Quarantine   │          │ SQL Quarantine   │"
  Write-Host "      │ .\private$\...   │          │ Status: Pending  │          │ Status: Parsed   │"
  Write-Host "      └─────────┬────────┘          └─────────┬────────┘          └─────────┬────────┘"
  Write-Host "                │                               │                               │"
  Write-Host "                │                               │                               │"
  Write-Host "                ▼                               ▼                               ▼"
  Write-Host ""
  Write-Host "      ┌──────────────────┐          ┌──────────────────┐          ┌──────────────────┐"
  Write-Host "      │ Insert into SQL  │          │ Extract DSP/SMI  │          │ Delete or mark   │"
  Write-Host "      │ Status=Pending   │          │ Timestamp        │          │ SMI='DFD','*D*'  │"
  Write-Host "      └──────────────────┘          │ Status=Parsed    │          │ Status=Ready     │"
  Write-Host "                                    └──────────────────┘          └──────────────────┘"
  Write-Host ""
  Write-Host "                                         ▼"
  Write-Host "                                         ▼"
  Write-Host "                                         ▼"
  Write-Host ""
  Write-Host "                           ┌──────────────────────────┐"
  Write-Host "                           │ 4. Display Status Summary│"
  Write-Host "                           │ QU-Display-MessageStatus │"
  Write-Host "                           └──────────────┬───────────┘"
  Write-Host "                                          │"
  Write-Host "                                          ▼"
  Write-Host "                           ┌──────────────────────────┐"
  Write-Host "                           │ Pending / Parsed / Ready │"
  Write-Host "                           │ Processed counts         │"
  Write-Host "                           └──────────────────────────┘"
  Write-Host ""
  Write-Host "                                          ▼"
  Write-Host "                                          ▼"
  Write-Host "                                          ▼"
  Write-Host ""
  Write-Host "                           ┌──────────────────────────┐"
  Write-Host "                           │ 5. File Ready Messages   │"
  Write-Host "                           │ QU-File-AllMessages...   │"
  Write-Host "                           └──────────────┬───────────┘"
  Write-Host "                                          │"
  Write-Host "                                          ▼"
  Write-Host "                           ┌──────────────────────────┐"
  Write-Host "                           │ Write BodyBinary to      │"
  Write-Host "                           │   D:\input (Aircom)      │"
  Write-Host "                           │ Status=Processed         │"
  Write-Host "                           └──────────────────────────┘"
  Write-Host ""
  Write-Host "                                          ▼"
  Write-Host "                                          ▼"
  Write-Host "                                          ▼"
  Write-Host ""
  Write-Host "                           ┌──────────────────────────┐"
  Write-Host "                           │ 6. User Guide System     │"
  Write-Host "                           │ QU-UserGuide (new)       │"
  Write-Host "                           └──────────────────────────┘"
  Write-Host ""
  Pause
}

function Show-UG-FileMessagesFlow {
  Clear-Host
  Write-Host "==================== FILE MESSAGES TO AIRCOM FLOW DIAGRAM ====================" -ForegroundColor Cyan
  Write-Host ""

  Write-Host "┌───────────────────────────────────────────────────────────────┐" -ForegroundColor Cyan
  Write-Host "│                QU-File-AllMessagesAdaptive                    │" -ForegroundColor Cyan
  Write-Host "│        Adaptive File Output + Safety + Batch Control          │" -ForegroundColor Cyan
  Write-Host "└───────────────────────────────────────────────────────────────┘" -ForegroundColor Cyan

  Write-Host "                               │"
  Write-Host "                               ▼"

  Write-Host "                 ┌──────────────────────────┐" -ForegroundColor Cyan
  Write-Host "                 │ INITIALIZATION           │" -ForegroundColor Cyan
  Write-Host "                 │ - Open SQL connection    │" -ForegroundColor White
  Write-Host "                 │ - Ensure D:\input exists │" -ForegroundColor White
  Write-Host "                 │ - batchSize = 20         │" -ForegroundColor White
  Write-Host "                 │ - RequeueSeq = 0         │" -ForegroundColor White
  Write-Host "                 └───────────────┬──────────┘" -ForegroundColor Cyan

  Write-Host "                                 │"
  Write-Host "                                 ▼"

  Write-Host "                 ┌──────────────────────────┐" -ForegroundColor Cyan
  Write-Host "                 │ MAIN LOOP (while \$true) │" -ForegroundColor Cyan
  Write-Host "                 └───────────────┬──────────┘" -ForegroundColor Cyan

  Write-Host "                                 │"
  Write-Host "                                 ▼"

  Write-Host "                 ┌──────────────────────────┐" -ForegroundColor Cyan
  Write-Host "                 │ Reset per-batch counter  │" -ForegroundColor White
  Write-Host "                 │ filesWrittenThisBatch=0  │" -ForegroundColor White
  Write-Host "                 └───────────────┬──────────┘" -ForegroundColor Cyan

  Write-Host "                                 │"
  Write-Host "                                 ▼"

  Write-Host "                 ┌────────────────────────────────────────────┐" -ForegroundColor Cyan
  Write-Host "                 │ Pull TOP(batchSize) FROM SQL WHERE Ready   │" -ForegroundColor White
  Write-Host "                 │ ORDER BY Id                                 │" -ForegroundColor White
  Write-Host "                 └───────────────┬────────────────────────────┘" -ForegroundColor Cyan

  Write-Host "                                 │"
  Write-Host "                     ┌───────────┴───────────┐"
  Write-Host "                     │ No rows?               │" -ForegroundColor Yellow
  Write-Host "                     │ (dt.Rows.Count = 0)    │" -ForegroundColor Yellow
  Write-Host "                     └───────────┬───────────┘"
  Write-Host "                                 │Yes" -ForegroundColor Green
  Write-Host "                                 ▼"

  Write-Host "                 ┌──────────────────────────┐" -ForegroundColor Cyan
  Write-Host "                 │  All messages processed  │" -ForegroundColor Green
  Write-Host "                 │  BREAK main loop         │" -ForegroundColor Green
  Write-Host "                 └──────────────────────────┘" -ForegroundColor Cyan

  Write-Host ""
  Write-Host "                                 │No"
  Write-Host "                                 ▼"

  Write-Host "        ┌────────────────────────────────────────────────────────────┐" -ForegroundColor Cyan
  Write-Host "        │ PROCESS EACH MESSAGE IN BATCH                              │" -ForegroundColor Cyan
  Write-Host "        └────────────────────────────────────────────────────────────┘" -ForegroundColor Cyan

  Write-Host "                                 │"
  Write-Host "                                 ▼"

  Write-Host "        ┌────────────────────────────────────────────────────────────┐" -ForegroundColor Cyan
  Write-Host "        │ For each row:                                              │" -ForegroundColor White
  Write-Host "        │   - Extract SMI, ParsedTimestamp                           │" -ForegroundColor White
  Write-Host "        │   - Build DayHHMM                                          │" -ForegroundColor White
  Write-Host "        │   - Increment RequeueSeq                                   │" -ForegroundColor White
  Write-Host "        │   - Build filename: Requeue_<SMI>_<DDHHMM>_<Seq>.txt       │" -ForegroundColor White
  Write-Host "        │   - Write file to D:\input                                 │" -ForegroundColor White
  Write-Host "        │   - Increment filesWrittenThisBatch                        │" -ForegroundColor White
  Write-Host "        │   - DELETE SQL row (Id)                                    │" -ForegroundColor White
  Write-Host "        └────────────────────────────────────────────────────────────┘" -ForegroundColor Cyan

  Write-Host ""
  Write-Host "                                 │"
  Write-Host "                                 ▼"

  Write-Host "        ┌────────────────────────────────────────────────────────────┐" -ForegroundColor Cyan
  Write-Host "        │ 6-second pause (Aircom drain window)                       │" -ForegroundColor DarkYellow
  Write-Host "        └────────────────────────────────────────────────────────────┘" -ForegroundColor Cyan

  Write-Host "                                 │"
  Write-Host "                                 ▼"

  Write-Host "        ┌────────────────────────────────────────────────────────────┐" -ForegroundColor Cyan
  Write-Host "        │ Count remaining Ready messages in SQL                      │" -ForegroundColor White
  Write-Host "        └────────────────────────────────────────────────────────────┘" -ForegroundColor Cyan

  Write-Host "                                 │"
  Write-Host "                                 ▼"

  Write-Host "        ┌────────────────────────────────────────────────────────────┐" -ForegroundColor Cyan
  Write-Host "        │ queueDepth = Get-QueueDepth()                              │" -ForegroundColor White
  Write-Host "        │ Display: queueDepth | filesWritten | remainingMessages     │" -ForegroundColor White
  Write-Host "        └────────────────────────────────────────────────────────────┘" -ForegroundColor Cyan

  Write-Host "                                 │"
  Write-Host "                                 ▼"

  Write-Host "        ┌────────────────────────────────────────────────────────────┐" -ForegroundColor Cyan
  Write-Host "        │ SAFETY CHECKS                                              │" -ForegroundColor Cyan
  Write-Host "        │                                                            │"
  Write-Host "        │ If queueDepth ≥150 and ≤200 → WARNING                      │" -ForegroundColor DarkYellow
  Write-Host "        │ If queueDepth >200 → HARD STOP                             │" -ForegroundColor Red
  Write-Host "        │     Enter wait loop:                                       │" -ForegroundColor White
  Write-Host "        │       Every 6 sec: check queueDepth                        │" -ForegroundColor White
  Write-Host "        │       Resume only when queueDepth < 50                     │" -ForegroundColor Green
  Write-Host "        └────────────────────────────────────────────────────────────┘" -ForegroundColor Cyan

  Write-Host "                                 │"
  Write-Host "                                 ▼"

  Write-Host "        ┌────────────────────────────────────────────────────────────┐" -ForegroundColor Cyan
  Write-Host "        │ ADAPTIVE BATCH SIZE LOGIC                                  │" -ForegroundColor Cyan
  Write-Host "        │                                                            │"
  Write-Host "        │ If queueDepth ≤100 → batchSize += 10                       │" -ForegroundColor Yellow
  Write-Host "        │ Else → batchSize = 20                                      │" -ForegroundColor White
  Write-Host "        └────────────────────────────────────────────────────────────┘" -ForegroundColor Cyan

  Write-Host "                                 │"
  Write-Host "                                 ▼"

  Write-Host "                 ┌──────────────────────────┐" -ForegroundColor Cyan
  Write-Host "                 │ Loop back to top         │" -ForegroundColor White
  Write-Host "                 │ (next batch)             │" -ForegroundColor White
  Write-Host "                 └──────────────────────────┘" -ForegroundColor Cyan

  Pause
}


###############################################################################
# TROUBLESHOOTING SECTION (unchanged)
###############################################################################

function Show-UG-Troubleshooting {
  Clear-Host
  Write-Host "==================== USER GUIDE: TROUBLESHOOTING ====================" -ForegroundColor Cyan
  Write-Host ""

  Write-Host "Common issues and how to resolve them:" -ForegroundColor White
  Write-Host ""

  #############################################################################
  # ISSUE 1 — Queue depth stays high (never drops below 150–200)
  #############################################################################
  Write-Host "1. Queue depth remains high (≥150 or >200):" -ForegroundColor Yellow
  Write-Host ""
  Write-Host "  Possible causes:"
  Write-Host "    • Aircom server is not draining messages fast enough"
  Write-Host "    • Aircom input service is paused or offline"
  Write-Host "    • Network latency or server load is slowing MSMQ processing"
  Write-Host ""
  Write-Host "  What the operator should do:"
  Write-Host "    • Check Aircom service health"
  Write-Host "    • Confirm D:\input files are being consumed"
  Write-Host "    • Wait for queue depth to fall below 50"
  Write-Host "    • Resume processing when the script automatically continues"
  Write-Host ""
  Write-Host ""

  #############################################################################
  # ISSUE 2 — Messages remain in 'Pending' and never become 'Parsed'
  #############################################################################
  Write-Host "2. Messages stuck in 'Pending':" -ForegroundColor Yellow
  Write-Host ""
  Write-Host "  Possible causes:"
  Write-Host "    • Parsing script was not run"
  Write-Host "    • Message body does not contain 'QU ' or '<?xml '"
  Write-Host "    • BodyText is empty or corrupted"
  Write-Host ""
  Write-Host "  What the operator should do:"
  Write-Host "    • Run: Parse DSP DayTime SMI"
  Write-Host "    • Review the Display Message Status Summary"
  Write-Host "    • If messages still fail to parse, inspect BodyText manually"
  Write-Host ""
  Write-Host ""

  #############################################################################
  # ISSUE 3 — Messages remain in 'Parsed' and never become 'Ready'
  #############################################################################
  Write-Host "3. Messages stuck in 'Parsed':" -ForegroundColor Yellow
  Write-Host ""
  Write-Host "  Possible causes:"
  Write-Host "    • Remove Nonessential Messages script was not run"
  Write-Host "    • SQL update failed"
  Write-Host ""
  Write-Host "  What the operator should do:"
  Write-Host "    • Run: Remove Nonessential Messages"
  Write-Host "    • Re-run Display Message Status Summary"
  Write-Host ""
  Write-Host ""

  #############################################################################
  # ISSUE 4 — Messages remain in 'Ready' and never get written to files
  #############################################################################
  Write-Host "4. Messages stuck in 'Ready':" -ForegroundColor Yellow
  Write-Host ""
  Write-Host "  Possible causes:"
  Write-Host "    • File-AllMessagesAdaptive script was not run"
  Write-Host "    • Output directory D:\input does not exist or is locked"
  Write-Host "    • Queue depth exceeded 200 and script is waiting"
  Write-Host ""
  Write-Host "  What the operator should do:"
  Write-Host "    • Verify D:\input exists and is writable"
  Write-Host "    • Check if script is in safety wait mode"
  Write-Host "    • Wait for queue depth to fall below 50"
  Write-Host ""
  Write-Host ""

  #############################################################################
  # ISSUE 5 — File writing errors (I/O errors, access denied, etc.)
  #############################################################################
  Write-Host "5. File writing errors:" -ForegroundColor Yellow
  Write-Host ""
  Write-Host "  Possible causes:"
  Write-Host "    • D:\input is full or read-only"
  Write-Host "    • Antivirus or backup software is locking files"
  Write-Host "    • Filename collision (extremely rare due to sequence logic)"
  Write-Host ""
  Write-Host "  What the operator should do:"
  Write-Host "    • Ensure D:\input has free space"
  Write-Host "    • Ensure no external process is locking the folder"
  Write-Host "    • Re-run the batch after resolving the issue"
  Write-Host ""
  Write-Host ""

  #############################################################################
  # ISSUE 6 — SQL connection failures
  #############################################################################
  Write-Host "6. SQL connection failures:" -ForegroundColor Yellow
  Write-Host ""
  Write-Host "  Possible causes:"
  Write-Host "    • Incorrect SQL username or password"
  Write-Host "    • SQL Server is offline or unreachable"
  Write-Host "    • Network authentication expired"
  Write-Host ""
  Write-Host "  What the operator should do:"
  Write-Host "    • Re-run the main script and re-enter credentials"
  Write-Host "    • Verify SQL Server is online"
  Write-Host "    • Check network connectivity"
  Write-Host ""
  Write-Host ""

  #############################################################################
  # ISSUE 7 — Messages missing DSP, SMI, or timestamp
  #############################################################################
  Write-Host "7. Missing DSP/SMI/timestamp fields:" -ForegroundColor Yellow
  Write-Host ""
  Write-Host "  Possible causes:"
  Write-Host "    • Message format does not match expected QU or XML structure"
  Write-Host "    • DSP line not found"
  Write-Host "    • SMI line missing or malformed"
  Write-Host ""
  Write-Host "  What the operator should do:"
  Write-Host "    • Inspect BodyText for malformed messages"
  Write-Host "    • Confirm message begins with 'QU ' or '<?xml '"
  Write-Host "    • If malformed, message may need manual review or deletion"
  Write-Host ""
  Write-Host ""

  #############################################################################
  # ISSUE 8 — Script appears to “hang” or pause for long periods
  #############################################################################
  Write-Host "8. Script appears paused or unresponsive:" -ForegroundColor Yellow
  Write-Host ""
  Write-Host "  This is normal during:"
  Write-Host "    • 6-second batch pauses"
  Write-Host "    • Queue-depth safety wait loop"
  Write-Host "    • Large batches of file writes"
  Write-Host ""
  Write-Host "  What the operator should do:"
  Write-Host "    • Check for queue-depth messages on screen"
  Write-Host "    • Verify Aircom is draining MSMQ"
  Write-Host "    • Allow script to continue automatically"
  Write-Host ""
  Write-Host ""

  Pause
}


###############################################################################
# DISPLAY ALL SECTIONS
###############################################################################

function Show-UG-All {
  Show-UG-Overview
  Show-UG-Lifecycle
  Show-UG-Workflow
  Show-UG-Scripts
  Show-UG-StatusCodes
  Show-UG-Safety
  Show-UG-FileNaming
  Show-UG-HighLevelDiagram
  Show-UG-FileMessagesFlow
  Show-UG-Troubleshooting
}


###############################################################################
# USER GUIDE MENU
###############################################################################

while ($true) {

  Clear-Host
  Write-Host "==================== QU USER GUIDE MENU $version ====================" -ForegroundColor Cyan
  Write-Host "1. Overview"
  Write-Host "2. Message Lifecycle"
  Write-Host "3. Operator Workflow"
  Write-Host "4. Script Descriptions"
  Write-Host "5. Status Codes"
  Write-Host "6. Safety Controls"
  Write-Host "7. File Naming Convention"
  Write-Host "8. High-Level Overview Diagram"
  Write-Host "9. Fle Messages Flow Diagram"
  Write-Host "T. Trouble Shooting"
  Write-Host "A. Display ALL Sections"
  Write-Host "X. Return to Main Menu"
  Write-Host "===================================================================="
  Write-Host ""

  $choice = Read-Host "Select an option"

  switch ($choice) {

    "1" { Show-UG-Overview }
    "2" { Show-UG-Lifecycle }
    "3" { Show-UG-Workflow }
    "4" { Show-UG-Scripts }
    "5" { Show-UG-StatusCodes }
    "6" { Show-UG-Safety }
    "7" { Show-UG-FileNaming }
    "8" { Show-UG-HighLevelDiagram }
    "9" { Show-UG-FileMessagesFlow }
    "T" { Show-UG-Troubleshooting }
    "A" { Show-UG-All }
    "X" { return }

    default {
      Write-Host "Invalid selection." -ForegroundColor Red
      Start-Sleep -Seconds 1
    }
  }
}

###############################################################################
# END OF SCRIPT
###############################################################################





