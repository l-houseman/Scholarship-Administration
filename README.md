# USRA Administration Protocol 2026
These scripts are used to automate a manual scholarship research administration process of naming and sorting large-volume files; creating reviewer assignments; consolidating hiring information; creating award letters and result notifications. This saved ~40 hours of work time in 2026

## Table of Contents
- [Application Sorting, File Naming, and Validation](#application-sorting-file-naming-and-validation)
- [Reviewer Assignment Creation](#reviewer-assignment-creation)
- [Hiring Information Document Naming and Consolidation](#hiring-information-document-naming-and-consolidation)
- [Award Document Creation](#award-document-creation)
- [Award Notification Creation](#award-notification-creation)


## Application Sorting, File Naming, and Validation
### Stage 1: Application Sorting and File Naming Protocol
#### How to run this protocol (Application Naming and Sorting Script) - PowerShell 7
**1.	Confirm the three input folders exist and are populated. This step takes three sets of incoming files:**
- USRA 2026 Applications/
  - Student Application Forms
  - Faculty Application Forms
  - Student Transcripts (Non-USask)
And
- Creates one folder per student under an organized root
- Renames each file into a standard naming convention
- Moves each file into the correct student folder
- Produces a CSV log of every action and a “missing documents” report
**2.	Confirm file naming follows the “final underscore student name” rule.**
**3.	Set:**
- $DryRun = $true
**4.	Run script in PowerShell7.**
**5.	Review:** 
- organize log (look for skips)
- missing report (look for gaps)
**6.	If results look correct:** 
- set $DryRun = $false
**7.	Run again to execute moves and folder creation.**
**8.	Archive the CSV log + missing report for traceability.**


#### Details of what this script does:
**File naming prerequisite (critical)**
!The student name must appear after the final underscore _ in the filename.!
Example:
- something_something_Lastname Firstname.pdf
The script’s Get-StudentNameFromBase() function:\
- takes the base filename (no extension)
- finds the last underscore
- treats the text after that underscore as the student name
If a file does not follow this pattern, it is skipped and logged as Skip_NoStudentSuffix.

**Standardized titles used (what files get renamed to)**
The script maps each source folder to a standardized document title prefix:
- Student application forms → 2026 USRA Student Application Form
- Faculty application forms → 2026 USRA Faculty Application Form
- Non USask transcripts → 2026 USRA Non-USask Transcript

**The final filename format becomes:**
{TitlePrefix}_{StudentName}{Extension}
Examples:
- 2026 USRA Student Application Form_Jane Smith.pdf
- 2026 USRA Faculty Application Form_Jane Smith.pdf
- 2026 USRA Non-USask Transcript_Jane Smith.pdf
  
**Collision handling (duplicate filenames)**
If $AllowOverwrite = $false and a target filename already exists, the script does not overwrite. It instead appends:
- (1), (2), etc.
Example:
- 2026 USRA Student Application Form_Jane Smith.pdf
- 2026 USRA Student Application Form_Jane Smith (1).pdf
  
**Execution mode controls**
The script has two major switches:
***Dry run***
- $DryRun = $true
  - no files are moved
  - the script writes a “Plan” row to the log for each file, showing what would happen
- $DryRun = $false
-   the script creates folders as needed and moves files
- 
**Optional subset processing**
- $ProcessOnly = @('Student Transcripts')
  - limits execution to specific source folders by display name
 
**What the script outputs (for audit and troubleshooting)**
Two CSVs are generated in the parent directory of $OrganizedRoot:
***1) Action log***
USRA_2026_organize_log_YYYYMMDD_HHMMSS.csv
Contains, per file:
- timestamp
- source folder + original filename
- extracted student name
- document type/title prefix
- new filename + destination path
- action: Plan, Moved, or a Skip_* reason
***2) Missing document report***
USRA_2026_missing_report_YYYYMMDD_HHMMSS.csv
Contains, per student folder detected:
- whether they have:
  - Student application
  - Faculty application
  - Transcript
- a “Missing” column listing what is absent
- the expected folder path
This report is built from a hashtable ($StudentDocs) that tracks which doc categories were successfully processed for each student.

### Stage 2: Splitting, naming, and sorting USask transcript files (Transcript Splitter) - Python
*Note: Because USask transcript files may contain full legal names that many students do not use, this still requires some manual consolidation (e.g., legal name on transcript is Susan Anna Joyce Double Lastname but student actually uses Susan Lastname in practice; you will need to consolidate “Susan Anna Joyce Double Lastname” file with “Susan Lastname” file). 

**This step takes one consolidated PDF containing many USask transcripts and produces:**
- One output PDF per student
- Named consistently
- Routed into the student’s folder under the organized root (created by Stage 1)
- With a fallback mechanism when a matching student folder is not found
- With preflight and diagnostics CSVs for audit

#### How to run this protocol (USask Transcript SplitNameSort Automation Script.py)
***Required inputs***
You need three inputs (paths set in the CONFIG section of the script):
1.	Consolidated transcript PDF
CONSOLIDATED_PDF = ...\2026 Tri-Agency USRA USask Transcripts.pdf
2.	Organized root (the student folders created in Step A)
ORGANIZED_ROOT = ...\USRA 2026_Organized
3.	Excel “roster” file mapping student number → first/last name
ROSTER_XLSX = ...\2026 Tri-Agency USRA USask Student Transcript Info.xlsx

**Confirm Step A is completed and ORGANIZED_ROOT exists with student folders.**
1.	Confirm CONSOLIDATED_PDF exists at the configured path.
2.	Confirm ROSTER_XLSX exists and contains student number + first/last name columns.
3.	Set in CONFIG:
   - PREFLIGHT_ONLY = True
Run script.
4.	Review:
  - preflight report
  - diagnostics CSVs
  - confirm folder existence issues
5.	Set:
  - PREFLIGHT_ONLY = False
  - DRY_RUN = True
6.	Run again to see planned outputs in console + log.
7.	Set: 
  - DRY_RUN = False
8.	Run to write per student transcript PDFs.
9.	Review USRA_2026_split_log_*.csv and confirm: 
  - counts align with expected students
  - fallback routing volume is reasonable
  - no errors in the log

#### Details of What this script does:
**Dependencies (local Python packages)**
The script imports:
- pypdf
- pdfplumber
- pandas + openpyxl
- rapidfuzz
  
**How transcript boundaries are detected (the splitting logic)**
The script finds the start of each student transcript by scanning each page’s extracted text for BOTH:
- Student Number: #####
- Name: Last, First
These are detected using regex:
- RE_STUDENT_NUM
- RE_NAME
Every page that contains both patterns is treated as a transcript start page.

**Segment creation**
Once start pages are identified, the script creates segments:
- transcript i starts at start_page[i]
- transcript i ends at start_page[i+1] - 1
- last transcript ends at the final page
  
**Legend handling**
The script treats page 1 of the consolidated PDF as a legend:
- it is excluded from the student transcript content
- it is appended to the end of each student’s output PDF
  
**Canonical naming of students (how final names are chosen)**
The script produces a final student folder/name using this priority order:
1.	Roster mapping by Student Number (preferred)
   - If the student number exists in the Excel roster, use: FirstName LastName from the roster as canonical.
2.	If not in roster:
  - Use the detected First Last from the PDF header.
3.	If canonical names exist and a match is unclear:
  - Use fuzzy matching (rapidfuzz) to pick the closest canonical roster name above a cutoff.
    
**Case normalization behavior**
The script only normalizes capitalization when the extracted name is clearly all caps or all lower. Mixed/intentional casing is preserved.

**Name sanitization**
Names are made filesystem safe:
- illegal filename characters replaced
- extra spaces trimmed
- trailing punctuation removed
  
**Preflight + diagnostics (what happens before writing PDFs)**
Before splitting, the script:
- reads the roster and builds:
  - a student number → canonical name mapping
- scans the PDF to detect transcript start pages
- produces preflight CSVs that document:
  - which student numbers are in the PDF
  - which student numbers are in the roster
  - which canonical folders exist under ORGANIZED_ROOT  
Outputs include:
- preflight_report_TIMESTAMP.csv
- diagnostics_pdf_not_in_excel_TIMESTAMP.csv
- diagnostics_excel_not_in_pdf_TIMESTAMP.csv
  
**PREFLIGHT_ONLY mode**
If PREFLIGHT_ONLY = True, the script runs those checks and exits without writing any PDFs.

**Destination routing rules (where outputs are written)**
For each student transcript segment, the script chooses a destination folder:
  **Normal routing**
  - If a folder exists at: ORGANIZED_ROOT/<Canonical Student Name>/ then output goes there.
  **Missing folder routing (fallback)**
  If the canonical folder does not exist:
- If ROUTE_MISSING_TO_FALLBACK = True, output goes to: ORGANIZED_ROOT/_NO_STUDENT_FOLDER_FOUND/<StudentNumber_CanonicalName>/    

This ensures transcripts still get created and nothing is silently dropped
If fallback routing is disabled and REQUIRE_EXISTING_FOLDERS = True, the segment is skipped and logged.

**Output file naming convention**
For each student, the script writes:
{YEAR} USRA USask Transcript_{student}.pdf
Where {student} is either:
Firstname Lastname (normal routing), or
Firstname Lastname_StudentNumber (fallback routing to reduce ambiguity)

File collisions are handled by appending (1), (2), etc.

**Split run log (audit record)**
A CSV log is written under ORGANIZED_ROOT:
- USRA_2026_split_log_TIMESTAMP.csv
Each row captures:
- start/end pages
- student number
- detected name (last, first)
- final canonical name used
- destination folder + output filename
- action: planned / written / fallback / skipped / error
- notes (including “legend appended”)

### Stage 3: Validating that each student folder exists, and has all  required documents
This script verifies that all USRA applicants have a folder and that each folder contains the required application and transcript documents for adjudication.

**It cross checks:**
- An official Excel roster of student names
- Against the organized USRA application folder structure
- And produces a single CSV audit report identifying:
  - missing folders
  - missing documents
  - extra folders not tied to a student in the roster
This script is designed to be repeatable year to year with minimal changes.

#### **What the script checks:**
For each student listed in the Excel roster, the script verifies:
Required documents (must exist)
- Faculty Application Form
  2026 USRA Faculty Application Form_Firstname Lastname.pdf
- Student Application Form
  2026 USRA Student Application Form_Firstname Lastname.pdf
Transcript rules
- At least one of the following must exist:
  - USask Transcript
    2026 USRA USask Transcript_Firstname Lastname.pdf
  - Non USask Transcript
    2026 USRA Non-USask Transcript_Firstname Lastname.pdf
- Some students may have both transcripts; this is valid.
  
**Folder level checks**
- Every student in the Excel roster should have exactly one folder
- Folders present on disk but not in the roster are flagged
- The fallback folder
_NO_STUDENT_FOLDER_FOUND
is recorded but not validated, since it is an intentional holding location

**Folder structure assumed**
USRA 2026_Organized\
- ├─ Firstname Lastname\
  - │  ├─ 2026 USRA Faculty Application Form_Firstname Lastname.pdf
  - │  ├─ 2026 USRA Student Application Form_Firstname Lastname.pdf
  - │  ├─ 2026 USRA USask Transcript_Firstname Lastname.pdf
  - │  └─ 2026 USRA Non-USask Transcript_Firstname Lastname.pdf   (optional)
  - │
- ├─ _NO_STUDENT_FOLDER_FOUND\
- │  └─ ...

Document matching is done by filename prefix, not full filename, so minor differences after the student name will not break the script.

**Inputs**
1. Excel roster
An .xlsx file containing at least:
- Student first name
- Student last name
The script is flexible about column headers (e.g., “First Name”, “Student First Name”, etc.).
2. Organized root folder
The directory containing one folder per student.

**Outputs**
The script generates one CSV file in the organized root, named:
USRA_2026_folder_and_file_audit_YYYYMMDD_HHMMSS.csv

**Output columns**
| Column | Meaning |
| ------ | ------- |
| NameOrFolder |	Student name (from Excel) or folder name |
| FolderExists |	YES / NO |
| FacultyApplicationForm |	YES / NO |
| StudentApplicationForm |	YES / NO |
| USaskTranscript |	YES / NO |
| NonUSaskTranscript |	YES / NO |
| Status |	Complete or specific issue(s) |

**Status values you may see**
- Complete
- MissingFolder
- MissingFacultyApplication
- MissingStudentApplication
- MissingTranscript
- ExtraFolder_NotInRoster
- FallbackFolder
Multiple issues are listed as semicolon separated values.

#### How to run the script
1.	Open Command Prompt (or PowerShell).
2.	Navigate to the project directory
3.	Run python Scripts\audit_usra_2026_folders_against_roster.py
4.	Open the generated CSV in Excel and filter by the status column

**Recommended workflow**
1.	Run the transcript splitting script first
2.	Manually resolve items in _NO_STUDENT_FOLDER_FOUND
3.	Run this audit script
4.	Filter the CSV to:
    - Status ≠ Complete
    - Status = MissingFolder
    - Status = ExtraFolder_NotInRoster
5.	Resolve issues and re run until clean

## Reviewer Assignment Creation
This protocol creates a list of which applications a reviewer can be assigned to, and creates a first take at creating assignments. Human audit is necessary, given that there are year-to-year intricacies that are not accounted for in the script. Python.

### Stage 1: Build eligibility possibilities (who can review what) 
Stage 1 creates a conflict‑free eligibility universe by excluding own‑department, own‑student, student‑department, and declared department conflicts, fully supporting cross‑listed departments and co‑supervision; it produces USRA_eligibility_pairs.xlsx for transparent auditing. 

**Purpose:** Construct a conflict‑free set of all allowed (application, reviewer) pairs before any assignment. This ensures the adjudicators are assigned eligible application packages. 
**Script name:** stage1_build_eligibility.py 
**Requirements:**
- All input spreadsheets, cleaned, ensuring that all departments/names are uniform
- Info about stream review eligibility 

**Inputs:** 
- Applicants workbook with one row per application, including columns:
  - AppID, StudentDepartment, SupervisorNSID, and the award stream in Which award are you applying for?
- Reviewers workbook with one row per reviewer, including columns:
  - ReviewerNSID, ReviewerDept, HasStudentInCompetition, StudentDepartmentList, ConflictDepartmentList, and StreamReview Eligibility  

**Normalization:**
- Convert NSIDs to lowercase for exact matching (avoids case‑related misses).
- Split any semicolon‑separated fields into arrays (e.g., multi‑department values, multi‑supervisor NSIDs).
- Treat department fields as lists and check for any overlap (supports cross‑listed student departments and multi‑appointment reviewers).  

**Hard exclusion rules (all must be satisfied to be eligible):** 
- Own department rule: Reject if the reviewer’s department list intersects the application’s student department list.
- Own student rule: Reject if the reviewer’s NSID appears in the application’s SupervisorNSID list (handles co‑supervisors).
- Student‑department rule: If the reviewer has a student in the competition, reject if the application’s student department appears in the reviewer’s StudentDepartmentList.
- Declared conflicts rule: Reject if the application’s student department appears in the reviewer’s ConflictDepartmentList.  

**Outputs:**
- USRA_eligibility_pairs.xlsx with:
  - EligibilityPairs: one row per permissible (AppID, ReviewerNSID) («Eligible = Yes»)
  - AppDiagnostics: count of eligible reviewers per AppID
  - ReviewerDiagnostics: count of eligible applications per ReviewerNSID
  - Apps_LT2_Eligible: applications with fewer than two eligible reviewers (should be empty before proceeding to ensure there are enough reviewers with the conflict rules applied) 

### Stage 2: Create assignments (who will review what) 
Stage 2 assigns two reviewers to each application, held reviewer loads to 10–12, preferred CIHR/SSHRC‑eligible reviewers for those streams (flagging any necessary fallbacks), and exported a ready‑to‑use workbook. 

**Purpose:** Assign two reviewers to each application, while holding reviewer workloads to 10–12 applications and respecting stream preferences for CIHR/SSHRC. 
**Script name:** stage2_make_assignments.py 

**Inputs:**
- Stage 1 output: USRA_eligibility_pairs.xlsx → EligibilityPairs (the allowed pair universe).
- Applicants workbook: used for each app’s award stream and to add the StudentName to outputs.
- Reviewers workbook: used for stream eligibility (StreamReview Eligibility) and to add ReviewerName to outputs.  

**Stream rule:**
- CIHR and SSHRC apps: the script first tries to assign both reviewers who list the relevant stream in StreamReview Eligibility.
  - If the eligible pool is too thin to fill both slots under the load caps, the script fills the remaining slot(s) from the general COI‑eligible pool, and logs the case in Exceptions as a “fallback used” for that CIHR/SSHRC app.
- NSERC apps: any eligible reviewer is acceptable (this may change year to year); there is no stream filtering here  

**Assignment process:** 
1. Order applications hardest‑first. CIHR/SSHRC apps are ordered by (a) count of stream‑eligible reviewers available, then (b) total COI‑eligible reviewers; NSERC apps are assigned later. This protects scarce expertise.
2. Pick two reviewers per app from the allowed pool under MAX_LOAD = 12, prioritizing the lowest current load and the stream‑eligible subset where required. The selection is deterministic with a fixed random seed to break ties.
3. Repair to minimum load. After the first pass, a swap routine uses only eligible alternatives to raise any reviewer below MIN_LOAD = 10 into range, while avoiding breaking a perfect CIHR/SSHRC stream match when possible. (If an app already needed a stream fallback, the script is more flexible.)
4. Validate.
   - Confirm total assignments = 195 apps × 2 = 390
   - Flag any reviewer outside the 10–12 range
   - Flag any CIHR/SSHRC app that required a stream fallback
   - All flags go to an Exceptions sheet for co-chair review.  

**Outputs:** 
- USRA_assignments.xlsx with four tabs:
  - Assignments_ByApp — AppID, StudentName, AwardStream, Reviewer1NSID, Reviewer1Name, Reviewer2NSID, Reviewer2Name (names are now merged automatically from your source files).
  - Assignments_ByReviewer — ReviewerNSID, ReviewerName, AssignedAppCount, list of AppIDs.
  - LoadSummary — one line per reviewer with assigned count and a boolean Within10to12 check.
  - Exceptions — any CIHR/SSHRC fallbacks (i.e., when the pool didn’t allow both reviewers to be stream‑eligible) and any load out‑of‑range or feasibility anomalies. 

## Hiring Information Document Naming and Consolidation 
This step uses 1) a PowerShell script to rename and file student hiring forms, 2) PowerQuery to consolidate ~200 individual Excel forms into one spreadsheet, then 3) Excel formulas + award result data from another spreadsheet to verify which to copy to the awarded, alternate list, and unawarded list tabs. 

### Stage 1 - Rename Individual Excel Sheets for Easy Reference Later
Script file: rename_move student hiring forms.ps1
This stage replicates the naming process. [Application Sorting, File Naming, and Validation](#application-sorting-file-naming-and-validation) to standardize file names and append the student's name so the individual file can be easily found later if necessary. 
**Purpose:** This PowerShell 7 script renames and moves completed “Information for Student Hiring” Excel forms into a single destination folder using a standardized filename format. It is intended for one‑time batch processing of hiring forms after submission.
The script does not modify file contents. It only renames and moves files.
**Maintenance Notes:**
This script is designed for:
- One‑time annual use
- Predictable file naming
- Small‑to‑moderate batch sizes
 
**What This Script Does:** For each file in the source folder:
- Extracts the student name from the filename after the final underscore (_)
- Renames the file to a standardized format:
- 2026 USRA Hiring Form_<Student Name>.xlsx
- Moves the renamed file into a single destination folder
- Prevents overwriting by appending (1), (2), etc. if a filename already exists
- Optionally runs in DryRun mode to preview changes without making them

**Required Folder Structure:**
Input (source)
All hiring forms must already exist in one folder:
- C:\Users\ljh440\USRA 2026 Applications\Inputs\Information for Student Hiring Forms\

Output (destination)
The script moves files into:
- C:\Users\ljh440\USRA 2026 Applications\USRA 2026_Hiring Forms\
- The destination folder will be created automatically if it does not exist.

**Filename Requirements:**
Each file must end with the student’s name after a final underscore. This structure is automatically created from our application portal. If the default structure differs year-to-year, the code would need to be updated. If no final underscore‑suffix name is found, the file is skipped with a warning.

Valid examples:
- faculty-information-for-student-hirin_Mary Cherneske.xlsx
- FacultyHiringForm_Jordan Lee.xlsx
- HiringForm_Final_Emily Zhao.xlsx
Invalid examples:
- HiringForm.xlsx
- HiringForm_Mary_Cherneske.xlsx   ← ambiguous underscore usage

**Configuration SettingsL**
At the top of the script:
- $SourceDir = '...Inputs\Information for Student Hiring Forms'
- $DestDir   = '...USRA 2026_Hiring Forms'
- $YearLabel = '2026' (controls the year shown in the renamed file)
- $DryRun    = $false ($true = preview only, $false = execute changes | always run once with $true to make sure it will work at intended)

**How to Run:**
1. Save the script as a .ps1 file (e.g., Rename-Hiring-Forms.ps1)
2. Open PowerShell 7
3. Run: PowerShellSet-ExecutionPolicy -Scope Process -ExecutionPolicy BypassShow more lines
4. Run the script
5. Review output in DryRun
6. Set $DryRun = $false
7. Run again to apply changes

### Stage 2 - Consolidate Individual Files with Powery Query
Power Query: One-Time Extraction of 1 row of hiring data per student

**Step-by-Step:**
1. Put all the individual forms in one folder
Make sure they are the only Excel files in that folder.
2. In your master workbook
Go to:
Data → Get Data → From File → From Folder
Select the folder.
Power Query will show you the file list.
Click Combine & Transform Data. 
3. In the preview window
Power Query shows a sample file.
Important:
On the left, choose the sheet named Info for Student Hiring (this matches your file).
Click OK.
This loads the entire sheet, but we will trim it.
4. Transform the sheet in Power Query
Once you're inside Power Query:
Step A — Remove the first row
This gets rid of the block of instructions/header text.
Home → Remove Rows → Remove Top Rows → enter 1.

Step B — Promote the new first row to headers
Home → Use First Row as Headers
This makes Row 2 (your labels) the official headers.

Step C — Filter to keep only Row 1 (the data row)
Your remaining table now has:

Row 1 → student’s actual values
Row 2+ → blank rows or unused fields

Apply a filter on any column to remove blank rows.
Example (choose a column with full data, like “Student Full Name”):
Filter → Remove null / blanks

Step D — Keep only Columns B–L
Select the columns you want to keep:

Student Full Name
Student NSID
Faculty Supervisor Name
Faculty NSID
Faculty Email
Length of Project
Proposed Start Date
Student Hourly Wage
CFOAPAL 1
CFOAPAL 2
Onboarding Contact NSID

Right‑click → Remove Other Columns
Now you have exactly your desired row.

Step E — Close & Load
Home → Close & Load
Power Query combines all the files using the same steps and appends the rows.

### Stage 3 - Sort into Awarded, Alternate List, and Unawarded Tabs
1. Sort hiring and result spreadsheets the same way (by NSID is usually best)
2. Insert that data beside the NSID column in the hiring spreadsheet. 
3. If everything is sorted and named correctly, the rows should match up. To check, run a quick IF statement in Excel to see if they actually match:
Allow for leading/trailing spaces | =IF(TRIM(A1)=TRIM(B1),"Match","No match")
4. Solve any problems, and copy awarded, alternate list, and unawarded info to individual tabs


## Award Document Creation
This step uses VBA code created by Imnoss Ltd to amplify Word MailMerge to create, save as PDF, and name individual award documents. 
1. Prepare spreadsheet of info and template letter as you usually do for a Word MailMerge
   - Spreadsheet must contain **EXACT** columns DocFolderPath and DocFileName (for Word copies_ and PdfFolderPath and PdfFileName (for the PDF copies)
3. Go through usual MailMerge steps to create individual letters, but stop before "save letters" step.
4a. Save the Macro for one time use (save for future use with 4b) 
- Open the VBA Editor
  - Press Alt + F11
  - The Visual Basic for Applications editor opens
- Insert a new module
  - In the left pane, find:
    - VBAProject (YourDocumentName)
  - Right‑click on it
  - Choose Insert → Module
- Paste the macro
  - Paste the entire macro exactly as‑is into the code window. Script Name: Imnross Ltd MailMerge to PDF
- Save Word doc as a macro-enabled file
4b. Save the Macro for all Word documents (recommended)
- Open Microsoft Word
  - Open Word normally (no document is required).
- Open the VBA Editor
  - Press: Alt + F11
  - This opens the Visual Basic for Applications editor.
- Locate the global template
  - In the left pane (Project Explorer), find: VBAProject (Normal)
  - If the pane isn’t visible: Click View → Project Explorer
- Insert a new module
  - Right‑click VBAProject (Normal)
  - Select Insert → Module
    - A blank code window opens.
- Paste the macro code
  - Paste the full macro exactly as provided (see 4a for script location)
- Save the macro
  - Close the VBA editor (Alt + Q), or close Word
    - When prompted "Save changes to Normal.dotm?", Click Yes
- The macro is now permanently stored on that machine.
5. Run the MailMergeToPdfBasic Macro, and let it run to create your individual award letters 

**To Use the Macro Later**
- Once saved to Normal.dotm: Open any Word mail‑merge document, select Macros and choose MailMergToPdfBasic, click Run
- The macro will execute using the currently open document.
  
## Award Notification Creation
This step uses VBA code to automate the creation of personalized email notifications with unique attachments from an Excel spreadsheet
- Process documentation to be provided
